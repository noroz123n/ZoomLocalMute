using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace ZoomLocalMute
{
    /// <summary>
    /// Persists enrolled voice profiles to disk so they can be
    /// re-applied to participants in future meetings.
    /// Stored as JSON in %AppData%\ZoomLocalMute\voices.json
    /// </summary>
    public class VoiceGallery
    {
        private static readonly string _dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ZoomLocalMute"
        );
        private static readonly string _file = Path.Combine(_dir, "voices.json");

        private Dictionary<string, SavedVoice> _voices = new();

        public IReadOnlyDictionary<string, SavedVoice> Voices => _voices;

        public VoiceGallery()
        {
            Load();
        }

        public void Save(string name, string base64Embedding, string? base64Sample = null)
        {
            _voices[name] = new SavedVoice
            {
                Name           = name,
                Base64Embedding = base64Embedding,
                Base64Sample   = base64Sample,
                SavedAt        = DateTime.Now
            };
            Persist();
        }

        public void Delete(string name)
        {
            _voices.Remove(name);
            Persist();
        }

        public SavedVoice? Get(string name) =>
            _voices.TryGetValue(name, out var v) ? v : null;

        private void Load()
        {
            try
            {
                if (!File.Exists(_file)) return;
                var json = File.ReadAllText(_file);
                _voices  = JsonSerializer.Deserialize<Dictionary<string, SavedVoice>>(json)
                           ?? new();
            }
            catch { _voices = new(); }
        }

        private void Persist()
        {
            try
            {
                Directory.CreateDirectory(_dir);
                File.WriteAllText(_file,
                    JsonSerializer.Serialize(_voices,
                        new JsonSerializerOptions { WriteIndented = true }));
            }
            catch { }
        }
    }

    public class SavedVoice
    {
        public string   Name            { get; set; } = "";
        public string   Base64Embedding { get; set; } = "";
        public string?  Base64Sample    { get; set; }
        public DateTime SavedAt         { get; set; }
    }
}