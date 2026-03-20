using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace ZoomLocalMute
{
    public class AudioEngine : IDisposable
    {
        public bool IsRunning { get; private set; }
        public event Action<string>? StatusChanged;
        public event Action<List<SpeakerStatus>>? SpeakersUpdated;
        public event Action<EnrollResult>? EnrollCompleted;
        public event Action<string, string?>? AudioSampleReceived; // name, base64wav
        public event Action<string>? AutoEnrolled;

        private ClientWebSocket? _ws;
        private CancellationTokenSource? _cts;
        private readonly ConcurrentDictionary<string, bool> _mutedLocally = new();
        private const string WsUri = "ws://localhost:5799";

        public async Task ConnectAsync()
        {
            _cts = new CancellationTokenSource();
            _ws  = new ClientWebSocket();
            try
            {
                await _ws.ConnectAsync(new Uri(WsUri), _cts.Token);
                StatusChanged?.Invoke("✅ Connected to audio backend.");
                IsRunning = true;
                _ = ReceiveLoopAsync();
            }
            catch (Exception ex)
            {
                StatusChanged?.Invoke($"⚠ Could not connect to Python backend: {ex.Message}");
            }
        }

        public void Start(object? _ = null)
        {
            SendMessage(new { type = "start_audio" });
            IsRunning = true;
            StatusChanged?.Invoke("▶ Audio filtering active.");
        }

        public void Stop()
        {
            SendMessage(new { type = "stop_audio" });
            IsRunning = false;
            StatusChanged?.Invoke("⏹ Stopped.");
        }

        public void UpdateParticipants(string[] names)
        {
            SendMessage(new { type = "update_participants", names });
        }

        public void SetMuted(string name, bool muted)
        {
            _mutedLocally[name] = muted;
            SendMessage(new { type = "set_muted", name, muted });
        }

        public bool IsMuted(string name) =>
            _mutedLocally.TryGetValue(name, out var m) && m;

        public void NotifyActiveSpeaker(string? name)
        {
            SendMessage(new { type = "set_active_speaker", name });
        }

        public void StartEnroll(string name)
        {
            SendMessage(new { type = "start_enroll", name });
        }

        public void StopEnroll()
        {
            SendMessage(new { type = "stop_enroll" });
        }

        public void RequestSample(string name)
        {
            SendMessage(new { type = "get_sample", name });
        }

        public void DeleteEnrollment(string name)
        {
            SendMessage(new { type = "delete_enrollment", name });
        }

        public void LoadSavedVoice(string participantName, string savedVoiceName, string base64Embedding)
        {
            SendMessage(new { type = "load_saved_voice", name = participantName, saved_name = savedVoiceName, embedding = base64Embedding });
        }

        private async Task ReceiveLoopAsync()
        {
            var buffer = new byte[131072]; // 128KB for audio samples
            try
            {
                while (_ws?.State == WebSocketState.Open)
                {
                    // Handle large messages by accumulating chunks
                    var sb = new StringBuilder();
                    WebSocketReceiveResult result;
                    do
                    {
                        result = await _ws.ReceiveAsync(
                            new ArraySegment<byte>(buffer), _cts!.Token
                        );
                        if (result.MessageType == WebSocketMessageType.Close) return;
                        sb.Append(Encoding.UTF8.GetString(buffer, 0, result.Count));
                    } while (!result.EndOfMessage);

                    HandleMessage(sb.ToString());
                }
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                StatusChanged?.Invoke($"⚠ Backend disconnected: {ex.Message}");
                IsRunning = false;
            }
        }

        private void HandleMessage(string json)
        {
            try
            {
                using var doc  = JsonDocument.Parse(json);
                var root = doc.RootElement;
                var type = root.GetProperty("type").GetString();

                switch (type)
                {
                    case "speakers":
                    {
                        var speakers = new List<SpeakerStatus>();
                        foreach (var el in root.GetProperty("data").EnumerateArray())
                        {
                            speakers.Add(new SpeakerStatus
                            {
                                Name          = el.GetProperty("name").GetString() ?? "",
                                Muted         = el.GetProperty("muted").GetBoolean(),
                                Enrolled      = el.GetProperty("enrolled").GetBoolean(),
                                EnrollmentPct = el.GetProperty("enrollment_pct").GetInt32(),
                                IsSpeaking    = el.GetProperty("is_speaking").GetBoolean(),
                                Confidence    = el.GetProperty("confidence").GetInt32(),
                                Interrupted   = el.GetProperty("interrupted").GetBoolean(),
                                HasSample     = el.GetProperty("has_sample").GetBoolean(),
                            });
                        }
                        SpeakersUpdated?.Invoke(speakers);
                        break;
                    }
                    case "status":
                    {
                        var msg = root.GetProperty("message").GetString();
                        if (msg != null) StatusChanged?.Invoke(msg);
                        break;
                    }
                    case "enroll_result":
                    {
                        // Extract all data BEFORE using block disposes
                        var data        = root.GetProperty("data");
                        var result      = new EnrollResult
                        {
                            Success     = data.GetProperty("success").GetBoolean(),
                            Name        = data.GetProperty("name").GetString() ?? "",
                            Interrupted = data.GetProperty("interrupted").GetBoolean(),
                            Duration    = data.GetProperty("duration").GetDouble(),
                            Message     = data.GetProperty("message").GetString() ?? "",
                            Embedding   = data.TryGetProperty("embedding", out var embEl)
                                          ? embEl.GetString() : null
                        };
                        EnrollCompleted?.Invoke(result);
                        break;
                    }
                    case "audio_sample":
                    {
                        var name   = root.GetProperty("name").GetString() ?? "";
                        var found  = root.GetProperty("found").GetBoolean();
                        var b64    = found && root.TryGetProperty("data", out var d)
                                     ? d.GetString() : null;
                        AudioSampleReceived?.Invoke(name, b64);
                        break;
                    }
                    case "auto_enrolled":
                    {
                        var name = root.GetProperty("name").GetString() ?? "";
                        AutoEnrolled?.Invoke(name);
                        break;
                    }
                }
            }
            catch { }
        }

        private void SendMessage(object msg)
        {
            if (_ws?.State != WebSocketState.Open) return;
            try
            {
                var json  = JsonSerializer.Serialize(msg);
                var bytes = Encoding.UTF8.GetBytes(json);
                Task.Run(() => _ws.SendAsync(
                    new ArraySegment<byte>(bytes),
                    WebSocketMessageType.Text,
                    true,
                    CancellationToken.None
                ));
            }
            catch { }
        }

        public void Dispose()
        {
            _cts?.Cancel();
            _ws?.Dispose();
        }
    }

    public class SpeakerStatus
    {
        public string Name          { get; set; } = "";
        public bool   Muted         { get; set; }
        public bool   Enrolled      { get; set; }
        public int    EnrollmentPct { get; set; }
        public bool   IsSpeaking    { get; set; }
        public int    Confidence    { get; set; }
        public bool   Interrupted   { get; set; }
        public bool   HasSample     { get; set; }
    }

    public class EnrollResult
    {
        public bool   Success     { get; set; }
        public string Name        { get; set; } = "";
        public bool   Interrupted { get; set; }
        public double Duration    { get; set; }
        public string Message     { get; set; } = "";
        public string? Embedding  { get; set; }
    }
}