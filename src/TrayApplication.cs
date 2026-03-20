using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ZoomLocalMute
{
    public class TrayApplication : IDisposable
    {
        private readonly NotifyIcon         _trayIcon;
        private readonly ZoomSpeakerTracker _tracker;
        private readonly AudioEngine        _engine;
        private readonly VoiceGallery       _gallery;
        private ControlWindow?              _window;

        public TrayApplication()
        {
            _tracker = new ZoomSpeakerTracker();
            _engine  = new AudioEngine();
            _gallery = new VoiceGallery();

            _engine.StatusChanged += msg =>
                _trayIcon?.ShowBalloonTip(2000, "ZoomLocalMute.ai", msg, ToolTipIcon.Info);

            _engine.SpeakersUpdated += statuses =>
                _window?.UpdateFromBackend(statuses);

            _engine.EnrollCompleted += result =>
                _window?.OnEnrollCompleted(result);

            _engine.AudioSampleReceived += (name, b64) =>
                _window?.OnAudioSampleReceived(name, b64);

            _engine.AutoEnrolled += name =>
                _window?.OnAutoEnrolled(name);

            _tracker.ParticipantsUpdated += names =>
            {
                _engine.UpdateParticipants(names);
                _window?.UpdateParticipants(names);
            };

            _tracker.ActiveSpeakerChanged += name =>
            {
                _engine.NotifyActiveSpeaker(name);
                _window?.HighlightSpeaker(name);
            };

            _trayIcon = new NotifyIcon
            {
                Icon    = SystemIcons.Application,
                Visible = true,
                Text    = "ZoomLocalMute.ai"
            };

            var menu = new ContextMenuStrip();
            menu.Items.Add("Open",         null, (_, _) => ShowWindow());
            menu.Items.Add("Voice Gallery", null, (_, _) => ShowGalleryWindow());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Exit",         null, (_, _) => ExitApp());
            _trayIcon.ContextMenuStrip = menu;
            _trayIcon.DoubleClick += (_, _) => ShowWindow();

            _tracker.Start();

            Task.Run(async () =>
            {
                await Task.Delay(1000);
                await _engine.ConnectAsync();
            });

            ShowWindow();
        }

        private void ShowWindow()
        {
            if (_window == null || _window.IsDisposed)
            {
                _window = new ControlWindow(_engine, _tracker, _gallery);
                _window.FormClosed += (_, _) => _window = null;
            }
            _window.Show();
            _window.BringToFront();
        }

        private void ShowGalleryWindow()
        {
            var gw = new GalleryWindow(_gallery, _engine);
            gw.Show();
        }

        private void ExitApp()
        {
            _engine.Stop();
            _tracker.Stop();
            _trayIcon.Visible = false;
            Application.Exit();
        }

        public void Dispose()
        {
            _engine.Dispose();
            _tracker.Dispose();
            _trayIcon.Dispose();
            _window?.Dispose();
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    //  Voice Gallery Window
    // ════════════════════════════════════════════════════════════════════════

    public class GalleryWindow : Form
    {
        private readonly VoiceGallery _gallery;
        private readonly AudioEngine  _engine;
        private readonly FlowLayoutPanel _panel;

        public GalleryWindow(VoiceGallery gallery, AudioEngine engine)
        {
            _gallery = gallery;
            _engine  = engine;

            Text            = "🎙 Voice Gallery";
            Size            = new Size(420, 500);
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition   = FormStartPosition.CenterScreen;
            BackColor       = Color.FromArgb(22, 22, 30);
            ForeColor       = Color.White;
            Font            = new Font("Segoe UI", 9.5f);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1,
                Padding = new Padding(12), BackColor = Color.Transparent
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(layout);

            layout.Controls.Add(new Label
            {
                Text      = "Saved voice profiles — click Apply to load into current meeting",
                Dock      = DockStyle.Fill,
                ForeColor = Color.FromArgb(140, 140, 170),
                Font      = new Font("Segoe UI", 8.5f),
                TextAlign = ContentAlignment.MiddleLeft
            }, 0, 0);

            var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            _panel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents  = false,
                AutoSize      = true,
                AutoSizeMode  = AutoSizeMode.GrowAndShrink,
                BackColor     = Color.Transparent,
                Padding       = new Padding(2)
            };
            scroll.Controls.Add(_panel);
            layout.Controls.Add(scroll, 0, 1);

            Refresh();
        }

        public new void Refresh()
        {
            _panel.Controls.Clear();
            foreach (var voice in _gallery.Voices.Values.OrderBy(v => v.Name))
                _panel.Controls.Add(new GalleryRow(voice, _gallery, _engine, Refresh));

            if (!_gallery.Voices.Any())
                _panel.Controls.Add(new Label
                {
                    Text      = "No saved voices yet.\nEnroll someone in a meeting and click \"Save Voice\".",
                    ForeColor = Color.FromArgb(120, 120, 150),
                    Font      = new Font("Segoe UI", 9f),
                    AutoSize  = true,
                    Padding   = new Padding(10)
                });
        }
    }

    public class GalleryRow : Panel
    {
        public GalleryRow(SavedVoice voice, VoiceGallery gallery, AudioEngine engine, Action refresh)
        {
            Height    = 48;
            Width     = 370;
            BackColor = Color.FromArgb(36, 36, 50);
            Margin    = new Padding(0, 0, 0, 2);

            var nameLabel = new Label
            {
                Text      = voice.Name,
                Left      = 10, Top = 6,
                Width     = 160, Height = 20,
                ForeColor = Color.White,
                Font      = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                AutoEllipsis = true
            };

            var dateLabel = new Label
            {
                Text      = voice.SavedAt.ToString("MMM dd, yyyy"),
                Left      = 10, Top = 26,
                Width     = 160, Height = 16,
                ForeColor = Color.FromArgb(120, 120, 150),
                Font      = new Font("Segoe UI", 7.5f)
            };

            // Apply to participant dropdown
            var applyBtn = new Button
            {
                Text      = "Apply to...",
                Left      = 178, Top = 10,
                Width     = 90, Height = 26,
                BackColor = Color.FromArgb(55, 100, 190),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 7.5f),
                Cursor    = Cursors.Hand
            };
            applyBtn.FlatAppearance.BorderSize = 0;
            applyBtn.Click += (_, _) =>
            {
                // Show a dropdown of current participants to apply to
                var menu = new ContextMenuStrip();
                // We'll populate with a placeholder — in practice the
                // participant list comes from the tracker via the engine
                menu.Items.Add("(join a meeting first)", null, null).Enabled = false;
                menu.Show(applyBtn, new System.Drawing.Point(0, applyBtn.Height));
            };

            var deleteBtn = new Button
            {
                Text      = "🗑",
                Left      = 274, Top = 10,
                Width     = 34, Height = 26,
                BackColor = Color.FromArgb(100, 40, 40),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 8f),
                Cursor    = Cursors.Hand
            };
            deleteBtn.FlatAppearance.BorderSize = 0;
            deleteBtn.Click += (_, _) =>
            {
                if (MessageBox.Show($"Delete voice profile for \"{voice.Name}\"?",
                    "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    gallery.Delete(voice.Name);
                    refresh();
                }
            };

            Controls.Add(nameLabel);
            Controls.Add(dateLabel);
            Controls.Add(applyBtn);
            Controls.Add(deleteBtn);
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    //  Main control window
    // ════════════════════════════════════════════════════════════════════════

    public class ControlWindow : Form
    {
        private readonly AudioEngine        _engine;
        private readonly ZoomSpeakerTracker _tracker;
        private readonly VoiceGallery       _gallery;

        private readonly Panel           _scrollPanel;
        private readonly FlowLayoutPanel _participantPanel;
        private readonly Button          _toggleBtn;
        private readonly Label           _statusLabel;
        private readonly Label           _speakerLabel;
        private readonly Label           _countLabel;
        private readonly Label           _backendLabel;

        private readonly Dictionary<string, ParticipantRow> _rows = new();

        public ControlWindow(AudioEngine engine, ZoomSpeakerTracker tracker, VoiceGallery gallery)
        {
            _engine  = engine;
            _tracker = tracker;
            _gallery = gallery;

            Text            = "ZoomLocalMute.ai";
            Size            = new Size(540, 640);
            MinimumSize     = new Size(480, 480);
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition   = FormStartPosition.CenterScreen;
            BackColor       = Color.FromArgb(22, 22, 30);
            ForeColor       = Color.White;
            Font            = new Font("Segoe UI", 9.5f);

            var layout = new TableLayoutPanel
            {
                Dock        = DockStyle.Fill,
                RowCount    = 8,
                ColumnCount = 1,
                Padding     = new Padding(14, 12, 14, 10),
                BackColor   = Color.Transparent
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            Controls.Add(layout);

            layout.Controls.Add(new Label
            {
                Text      = "🔇  ZoomLocalMute.ai",
                Dock      = DockStyle.Fill,
                Font      = new Font("Segoe UI", 13f, FontStyle.Bold),
                ForeColor = Color.FromArgb(100, 190, 255),
                TextAlign = ContentAlignment.MiddleLeft
            }, 0, 0);

            _backendLabel = new Label
            {
                Text      = "⏳ Connecting to Python backend...",
                Dock      = DockStyle.Fill,
                ForeColor = Color.FromArgb(200, 160, 80),
                TextAlign = ContentAlignment.MiddleLeft,
                Font      = new Font("Segoe UI", 8.5f)
            };
            layout.Controls.Add(_backendLabel, 0, 1);

            _speakerLabel = new Label
            {
                Text      = "🎙  Active speaker: —",
                Dock      = DockStyle.Fill,
                ForeColor = Color.FromArgb(160, 220, 160),
                TextAlign = ContentAlignment.MiddleLeft,
                Font      = new Font("Segoe UI", 9f)
            };
            layout.Controls.Add(_speakerLabel, 0, 2);

            _countLabel = new Label
            {
                Text      = "Participants: 0",
                Dock      = DockStyle.Fill,
                ForeColor = Color.FromArgb(120, 120, 150),
                TextAlign = ContentAlignment.MiddleLeft,
                Font      = new Font("Segoe UI", 8.5f)
            };
            layout.Controls.Add(_countLabel, 0, 3);

            _scrollPanel = new Panel
            {
                Dock       = DockStyle.Fill,
                BackColor  = Color.FromArgb(28, 28, 40),
                AutoScroll = false
            };
            _participantPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents  = false,
                AutoSize      = true,
                AutoSizeMode  = AutoSizeMode.GrowAndShrink,
                BackColor     = Color.Transparent,
                Padding       = new Padding(4)
            };
            _scrollPanel.Controls.Add(_participantPanel);
            _scrollPanel.AutoScroll = true;
            layout.Controls.Add(_scrollPanel, 0, 4);
            _scrollPanel.Resize += (_, _) => ResizeRows();

            _toggleBtn = new Button
            {
                Text      = "⏹  Stop Filtering",
                Dock      = DockStyle.Fill,
                BackColor = Color.FromArgb(190, 55, 55),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 10f, FontStyle.Bold),
                Cursor    = Cursors.Hand,
                Margin    = new Padding(0, 4, 0, 0)
            };
            _toggleBtn.FlatAppearance.BorderSize = 0;
            _toggleBtn.Click += OnToggle;
            layout.Controls.Add(_toggleBtn, 0, 5);

            _statusLabel = new Label
            {
                Text      = "Ready. Start Python backend first, then join a Zoom meeting.",
                Dock      = DockStyle.Fill,
                ForeColor = Color.FromArgb(110, 110, 140),
                TextAlign = ContentAlignment.MiddleLeft,
                Font      = new Font("Segoe UI", 8f)
            };
            layout.Controls.Add(_statusLabel, 0, 6);

            layout.Controls.Add(new Label
            {
                Text      = "ℹ️  Hold 🎤 to enroll • ▶ to hear sample • 💾 to save to gallery",
                Dock      = DockStyle.Fill,
                ForeColor = Color.FromArgb(90, 90, 120),
                TextAlign = ContentAlignment.MiddleLeft,
                Font      = new Font("Segoe UI", 7.5f)
            }, 0, 7);

            _engine.StatusChanged += msg => InvokeIfRequired(() =>
            {
                _statusLabel.Text = msg;
                if (msg.Contains("Connected"))
                {
                    _backendLabel.Text      = "✅ Python backend connected";
                    _backendLabel.ForeColor = Color.FromArgb(100, 200, 100);
                    _toggleBtn.Text         = "⏹  Stop Filtering";
                    _toggleBtn.BackColor    = Color.FromArgb(190, 55, 55);
                }
                else if (msg.Contains("filtering active"))
                {
                    _toggleBtn.Text      = "⏹  Stop Filtering";
                    _toggleBtn.BackColor = Color.FromArgb(190, 55, 55);
                }
                else if (msg.Contains("disconnected") || msg.Contains("Could not"))
                {
                    _backendLabel.Text      = "❌ Backend not running — start audio_backend.py first";
                    _backendLabel.ForeColor = Color.FromArgb(220, 80, 80);
                    _toggleBtn.Text         = "▶  Start Filtering";
                    _toggleBtn.BackColor    = Color.FromArgb(34, 150, 90);
                }
                else if (msg.Contains("Stopped"))
                {
                    _toggleBtn.Text      = "▶  Start Filtering";
                    _toggleBtn.BackColor = Color.FromArgb(34, 150, 90);
                }
            });
        }

        // ── Engine events ────────────────────────────────────────────────────

        public void OnEnrollCompleted(EnrollResult result)
        {
            InvokeIfRequired(() =>
            {
                if (!result.Success) return;
                if (result.Interrupted)
                    MessageBox.Show(
                        "Another voice was detected during enrollment.\n" +
                        "The dominant voice was used.\n\n" +
                        "Listen to the sample to verify it's the right person.",
                        "Interruption Detected",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning
                    );

                // Offer to save to gallery
                if (result.Embedding != null)
                {
                    var save = MessageBox.Show(
                        $"Enrollment complete for {result.Name}.\n\nSave this voice to your gallery for future meetings?",
                        "Save Voice Profile",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question
                    );
                    if (save == DialogResult.Yes)
                    {
                        string galleryName = result.Name;
                        // Allow renaming
                        using var dlg = new RenameDialog(result.Name);
                        if (dlg.ShowDialog() == DialogResult.OK)
                            galleryName = dlg.NewName;
                        _gallery.Save(galleryName, result.Embedding);
                        _statusLabel.Text = $"💾 Saved \"{galleryName}\" to voice gallery.";
                    }
                }
            });
        }

        public void OnAudioSampleReceived(string name, string? b64)
        {
            InvokeIfRequired(() =>
            {
                if (b64 == null)
                {
                    MessageBox.Show($"No audio sample available for {name}.",
                        "No Sample", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                try
                {
                    var tmpPath = System.IO.Path.Combine(
                        System.IO.Path.GetTempPath(),
                        $"zlm_sample_{name.Replace(" ", "_")}.wav"
                    );
                    System.IO.File.WriteAllBytes(tmpPath, Convert.FromBase64String(b64));
                    new System.Media.SoundPlayer(tmpPath).Play();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Playback error: {ex.Message}",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            });
        }

        public void OnAutoEnrolled(string name)
        {
            InvokeIfRequired(() =>
            {
                _statusLabel.Text = $"✅ Auto-enrolled: {name}";
            });
        }

        // ── Backend speaker updates ──────────────────────────────────────────

        public void UpdateFromBackend(List<SpeakerStatus> statuses)
        {
            InvokeIfRequired(() =>
            {
                foreach (var s in statuses)
                    if (_rows.TryGetValue(s.Name, out var row))
                        row.UpdateStatus(s);

                var speaking = statuses.FirstOrDefault(s => s.IsSpeaking);
                if (speaking != null)
                    _speakerLabel.Text = $"🎙  Active speaker: {speaking.Name}";
            });
        }

        // ── Participant list ─────────────────────────────────────────────────

        public void UpdateParticipants(string[] names)
        {
            InvokeIfRequired(() =>
            {
                _participantPanel.SuspendLayout();

                foreach (var name in names)
                {
                    if (!_rows.ContainsKey(name))
                    {
                        // Check if we have a saved voice for this name
                        var saved = _gallery.Voices.Values
                            .FirstOrDefault(v => v.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

                        var row = new ParticipantRow(name, _engine, _gallery, GetRowWidth(), saved);
                        _rows[name] = row;
                        _participantPanel.Controls.Add(row);

                        // Auto-apply saved voice if exact name match
                        if (saved != null)
                        {
                            _engine.LoadSavedVoice(name, saved.Name, saved.Base64Embedding);
                            _statusLabel.Text = $"🎙 Auto-applied saved voice for {name}";
                        }
                    }
                }

                foreach (var name in _rows.Keys.Except(names).ToList())
                {
                    _participantPanel.Controls.Remove(_rows[name]);
                    _rows[name].Dispose();
                    _rows.Remove(name);
                }

                _participantPanel.ResumeLayout();
                _countLabel.Text = $"Participants: {_rows.Count}";

                if (!_rows.Any())
                    _statusLabel.Text = "No participants found. Open Zoom's Participants panel.";
            });
        }

        public void HighlightSpeaker(string? name)
        {
            InvokeIfRequired(() =>
            {
                if (name != null)
                    _speakerLabel.Text = $"🎙  Active speaker: {name}";
                foreach (var (n, row) in _rows)
                    row.SetSpeakingFallback(n == name);
            });
        }

        private int GetRowWidth() =>
            Math.Max(200, _scrollPanel.ClientSize.Width - 20);

        private void ResizeRows()
        {
            var w = GetRowWidth();
            foreach (var row in _rows.Values)
                row.SetWidth(w);
        }

        private void OnToggle(object? sender, EventArgs e)
        {
            if (_engine.IsRunning)
            {
                _engine.Stop();
                _toggleBtn.Text      = "▶  Start Filtering";
                _toggleBtn.BackColor = Color.FromArgb(34, 150, 90);
            }
            else
            {
                _engine.Start();
                _toggleBtn.Text      = "⏹  Stop Filtering";
                _toggleBtn.BackColor = Color.FromArgb(190, 55, 55);
            }
        }

        private void InvokeIfRequired(Action action)
        {
            if (InvokeRequired) BeginInvoke(action);
            else action();
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    //  Rename dialog for saving to gallery
    // ════════════════════════════════════════════════════════════════════════

    public class RenameDialog : Form
    {
        public string NewName { get; private set; }
        private readonly TextBox _textBox;

        public RenameDialog(string defaultName)
        {
            NewName         = defaultName;
            Text            = "Save Voice Profile";
            Size            = new Size(340, 130);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition   = FormStartPosition.CenterParent;
            MaximizeBox     = false;
            MinimizeBox     = false;
            BackColor       = Color.FromArgb(30, 30, 40);
            ForeColor       = Color.White;

            var lbl = new Label
            {
                Text     = "Save as name:",
                Left = 10, Top = 12, Width = 300,
                ForeColor = Color.FromArgb(180, 180, 200)
            };

            _textBox = new TextBox
            {
                Text      = defaultName,
                Left      = 10, Top = 34, Width = 300,
                BackColor = Color.FromArgb(44, 44, 60),
                ForeColor = Color.White,
                Font      = new Font("Segoe UI", 10f)
            };

            var okBtn = new Button
            {
                Text      = "Save",
                Left      = 150, Top = 62, Width = 80, Height = 28,
                BackColor = Color.FromArgb(55, 130, 55),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.OK
            };
            okBtn.FlatAppearance.BorderSize = 0;
            okBtn.Click += (_, _) => NewName = _textBox.Text.Trim();

            var cancelBtn = new Button
            {
                Text         = "Cancel",
                Left         = 240, Top = 62, Width = 70, Height = 28,
                BackColor    = Color.FromArgb(80, 40, 40),
                ForeColor    = Color.White,
                FlatStyle    = FlatStyle.Flat,
                DialogResult = DialogResult.Cancel
            };
            cancelBtn.FlatAppearance.BorderSize = 0;

            Controls.AddRange(new Control[] { lbl, _textBox, okBtn, cancelBtn });
            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    //  Participant row
    // ════════════════════════════════════════════════════════════════════════

    public class ParticipantRow : Panel
    {
        private readonly string       _name;
        private readonly AudioEngine  _engine;
        private readonly VoiceGallery _gallery;
        private readonly Label        _nameLabel;
        private readonly Button       _muteBtn;
        private readonly Button       _enrollBtn;
        private readonly Button       _listenBtn;
        private readonly Button       _saveBtn;
        private readonly Button       _resetBtn;
        private readonly Button       _galleryBtn;
        private readonly ProgressBar  _enrollBar;
        private readonly Label        _enrollLabel;
        private readonly Label        _confidenceLabel;
        private bool _isMuted;
        private bool _isSpeaking;
        private bool _isEnrolled;
        private bool _isEnrolling;
        private string? _currentEmbedding;

        private static readonly Color BgNormal    = Color.FromArgb(36, 36, 50);
        private static readonly Color BgSpeaking  = Color.FromArgb(28, 72, 48);
        private static readonly Color BgMuted     = Color.FromArgb(72, 28, 28);
        private static readonly Color BgMutedSpk  = Color.FromArgb(90, 40, 28);
        private static readonly Color BgEnrolling = Color.FromArgb(60, 50, 20);
        private static readonly Color BgSaved     = Color.FromArgb(28, 50, 72);

        public ParticipantRow(string name, AudioEngine engine, VoiceGallery gallery,
                              int width, SavedVoice? savedVoice = null)
        {
            _name    = name;
            _engine  = engine;
            _gallery = gallery;
            _isMuted = engine.IsMuted(name);

            Height    = 80;
            Width     = width;
            BackColor = savedVoice != null ? BgSaved : BgNormal;
            Margin    = new Padding(0, 0, 0, 3);

            // Row 1: Name + Mute
            _nameLabel = new Label
            {
                Text         = name + (savedVoice != null ? " 💾" : ""),
                Left         = 10, Top = 5,
                Width        = width - 140, Height = 20,
                ForeColor    = Color.White,
                Font         = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                AutoEllipsis = true,
                Anchor       = AnchorStyles.Left | AnchorStyles.Top
            };

            _muteBtn = new Button
            {
                Text      = _isMuted ? "Unmute for me" : "Mute for me",
                Left      = width - 128, Top = 4,
                Width     = 120, Height = 26,
                BackColor = _isMuted
                    ? Color.FromArgb(180, 55, 55)
                    : Color.FromArgb(55, 100, 190),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 8f),
                Cursor    = Cursors.Hand,
                Anchor    = AnchorStyles.Right | AnchorStyles.Top
            };
            _muteBtn.FlatAppearance.BorderSize = 0;
            _muteBtn.Click += OnMuteClick;

            // Row 2: Enroll + Listen + Save + Reset + Gallery
            _enrollBtn = new Button
            {
                Text      = "🎤 Hold to Enroll",
                Left      = 10, Top = 32,
                Width     = 118, Height = 22,
                BackColor = Color.FromArgb(55, 110, 55),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 7f),
                Cursor    = Cursors.Hand,
                Anchor    = AnchorStyles.Left | AnchorStyles.Top
            };
            _enrollBtn.FlatAppearance.BorderSize = 0;
            _enrollBtn.MouseDown += OnEnrollDown;
            _enrollBtn.MouseUp   += OnEnrollUp;

            _listenBtn = new Button
            {
                Text      = "▶",
                Left      = 134, Top = 32,
                Width     = 28, Height = 22,
                BackColor = Color.FromArgb(40, 80, 120),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 8f),
                Cursor    = Cursors.Hand,
                Visible   = false,
                Anchor    = AnchorStyles.Left | AnchorStyles.Top
            };
            _listenBtn.FlatAppearance.BorderSize = 0;
            _listenBtn.Click += (_, _) => _engine.RequestSample(_name);

            _saveBtn = new Button
            {
                Text      = "💾",
                Left      = 168, Top = 32,
                Width     = 28, Height = 22,
                BackColor = Color.FromArgb(50, 100, 50),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 8f),
                Cursor    = Cursors.Hand,
                Visible   = false,
                Anchor    = AnchorStyles.Left | AnchorStyles.Top
            };
            _saveBtn.FlatAppearance.BorderSize = 0;
            _saveBtn.Click += OnSaveClick;

            _resetBtn = new Button
            {
                Text      = "🗑",
                Left      = 202, Top = 32,
                Width     = 28, Height = 22,
                BackColor = Color.FromArgb(100, 40, 40),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 8f),
                Cursor    = Cursors.Hand,
                Visible   = false,
                Anchor    = AnchorStyles.Left | AnchorStyles.Top
            };
            _resetBtn.FlatAppearance.BorderSize = 0;
            _resetBtn.Click += OnResetClick;

            _galleryBtn = new Button
            {
                Text      = "📂 Gallery",
                Left      = 236, Top = 32,
                Width     = 80, Height = 22,
                BackColor = Color.FromArgb(60, 60, 110),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 7f),
                Cursor    = Cursors.Hand,
                Anchor    = AnchorStyles.Left | AnchorStyles.Top
            };
            _galleryBtn.FlatAppearance.BorderSize = 0;
            _galleryBtn.Click += OnGalleryClick;

            // Row 3: Progress / confidence
            _enrollBar = new ProgressBar
            {
                Left    = 10, Top = 60,
                Width   = width - 140, Height = 6,
                Minimum = 0, Maximum = 100, Value = 0,
                Style   = ProgressBarStyle.Continuous,
                Anchor  = AnchorStyles.Left | AnchorStyles.Top,
                Visible = true
            };

            _enrollLabel = new Label
            {
                Text      = "Waiting for voice...",
                Left      = 10, Top = 59,
                Width     = width - 140, Height = 14,
                ForeColor = Color.FromArgb(140, 140, 170),
                Font      = new Font("Segoe UI", 7f),
                Visible   = false,
                Anchor    = AnchorStyles.Left | AnchorStyles.Top
            };

            _confidenceLabel = new Label
            {
                Text      = "",
                Left      = 10, Top = 59,
                Width     = width - 140, Height = 14,
                ForeColor = Color.FromArgb(120, 200, 120),
                Font      = new Font("Segoe UI", 7f),
                Visible   = false,
                Anchor    = AnchorStyles.Left | AnchorStyles.Top
            };

            Controls.Add(_nameLabel);
            Controls.Add(_muteBtn);
            Controls.Add(_enrollBtn);
            Controls.Add(_listenBtn);
            Controls.Add(_saveBtn);
            Controls.Add(_resetBtn);
            Controls.Add(_galleryBtn);
            Controls.Add(_enrollBar);
            Controls.Add(_enrollLabel);
            Controls.Add(_confidenceLabel);
        }

        public void UpdateStatus(SpeakerStatus status)
        {
            if (InvokeRequired) { BeginInvoke(() => UpdateStatus(status)); return; }

            _isMuted    = status.Muted;
            _isSpeaking = status.IsSpeaking;
            _isEnrolled = status.Enrolled;

            _muteBtn.Text      = _isMuted ? "Unmute for me" : "Mute for me";
            _muteBtn.BackColor = _isMuted
                ? Color.FromArgb(180, 55, 55)
                : Color.FromArgb(55, 100, 190);

            if (_isEnrolled)
            {
                _enrollBar.Visible       = false;
                _enrollLabel.Visible     = false;
                _listenBtn.Visible       = status.HasSample;
                _saveBtn.Visible         = true;
                _resetBtn.Visible        = true;
                _confidenceLabel.Visible = true;
                _confidenceLabel.Text    = $"Confidence: {status.Confidence}%"
                    + (status.Interrupted ? "  ⚠" : "");
                _confidenceLabel.ForeColor = status.Confidence >= 70
                    ? Color.FromArgb(100, 200, 100)
                    : status.Confidence >= 50
                        ? Color.FromArgb(200, 180, 60)
                        : Color.FromArgb(200, 80, 80);
            }
            else
            {
                _enrollBar.Value         = status.EnrollmentPct;
                _enrollBar.Visible       = !_isEnrolling;
                _enrollLabel.Visible     = !_isEnrolling;
                _enrollLabel.Text        = status.EnrollmentPct == 0
                    ? "Waiting for voice..."
                    : $"Auto-enrolling... {status.EnrollmentPct}%";
                _listenBtn.Visible       = false;
                _saveBtn.Visible         = false;
                _resetBtn.Visible        = false;
                _confidenceLabel.Visible = false;
            }

            UpdateColor();
        }

        public void SetSpeakingFallback(bool speaking)
        {
            _isSpeaking = speaking;
            UpdateColor();
        }

        public void SetWidth(int width)
        {
            Width                  = width;
            _nameLabel.Width       = width - 140;
            _enrollBar.Width       = width - 140;
            _enrollLabel.Width     = width - 140;
            _confidenceLabel.Width = width - 140;
            _muteBtn.Left          = width - 128;
        }

        private void UpdateColor()
        {
            if (InvokeRequired) { BeginInvoke(UpdateColor); return; }
            if (_isEnrolling)
                BackColor = BgEnrolling;
            else
                BackColor = (_isMuted && _isSpeaking) ? BgMutedSpk
                          : _isMuted                  ? BgMuted
                          : _isSpeaking               ? BgSpeaking
                                                      : BgNormal;
        }

        private void OnMuteClick(object? sender, EventArgs e)
        {
            _isMuted = !_isMuted;
            _engine.SetMuted(_name, _isMuted);
            _muteBtn.Text      = _isMuted ? "Unmute for me" : "Mute for me";
            _muteBtn.BackColor = _isMuted
                ? Color.FromArgb(180, 55, 55)
                : Color.FromArgb(55, 100, 190);
            UpdateColor();
        }

        private void OnEnrollDown(object? sender, MouseEventArgs e)
        {
            _isEnrolling         = true;
            _enrollBtn.Text      = "🔴 Recording...";
            _enrollBtn.BackColor = Color.FromArgb(160, 40, 40);
            _enrollBar.Visible   = false;
            UpdateColor();
            _engine.StartEnroll(_name);
        }

        private void OnEnrollUp(object? sender, MouseEventArgs e)
        {
            _isEnrolling         = false;
            _enrollBtn.Text      = "🎤 Hold to Enroll";
            _enrollBtn.BackColor = Color.FromArgb(55, 110, 55);
            UpdateColor();
            _engine.StopEnroll();
        }

        private void OnSaveClick(object? sender, EventArgs e)
        {
            // Request the embedding from backend via enroll_result
            // For now, show save dialog — embedding comes back via EnrollCompleted
            MessageBox.Show(
                "To save this voice profile, hold the 🎤 button briefly to trigger a re-enrollment,\n" +
                "then choose 'Yes' when asked to save to gallery.",
                "Save Voice",
                MessageBoxButtons.OK, MessageBoxIcon.Information
            );
        }

        private void OnResetClick(object? sender, EventArgs e)
        {
            if (MessageBox.Show($"Delete enrollment for {_name}?",
                "Confirm Reset", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                _engine.DeleteEnrollment(_name);
        }

        private void OnGalleryClick(object? sender, EventArgs e)
        {
            // Show saved voices and let user pick one to apply
            var voices = _gallery.Voices.Values.OrderBy(v => v.Name).ToList();
            if (!voices.Any())
            {
                MessageBox.Show(
                    "No saved voices in your gallery yet.\nEnroll someone and save their voice first.",
                    "Gallery Empty", MessageBoxButtons.OK, MessageBoxIcon.Information
                );
                return;
            }

            var menu = new ContextMenuStrip();
            menu.BackColor = Color.FromArgb(36, 36, 50);
            menu.ForeColor = Color.White;

            foreach (var voice in voices)
            {
                var v    = voice;
                var item = new ToolStripMenuItem($"Apply \"{v.Name}\" ({v.SavedAt:MMM dd})")
                {
                    ForeColor = Color.White,
                    BackColor = Color.FromArgb(36, 36, 50)
                };
                item.Click += (_, _) =>
                {
                    _engine.LoadSavedVoice(_name, v.Name, v.Base64Embedding);
                    _nameLabel.Text = _name + " 💾";
                    BackColor       = BgSaved;
                };
                menu.Items.Add(item);
            }

            menu.Show(_galleryBtn, new System.Drawing.Point(0, _galleryBtn.Height));
        }
    }
}