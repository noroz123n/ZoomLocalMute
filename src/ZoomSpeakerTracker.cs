using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;

namespace ZoomLocalMute
{
    public class ZoomSpeakerTracker : IDisposable
    {
        public event Action<string?>? ActiveSpeakerChanged;
        public event Action<string[]>? ParticipantsUpdated;

        private string? _lastSpeaker;
        private string[] _lastParticipants = Array.Empty<string>();
        private System.Threading.Timer? _pollTimer;
        private readonly UIA3Automation _automation = new();
        private const int PollIntervalMs = 600;

        private static readonly HashSet<string> _blocklist = new(StringComparer.OrdinalIgnoreCase)
        {
            "mute", "unmute", "start video", "stop video", "participants",
            "chat", "share screen", "record", "reactions", "more", "leave",
            "end", "security", "polls", "q&a", "breakout rooms", "whiteboard",
            "apps", "captions", "raise hand", "lower hand", "yes", "no",
            "go slower", "go faster", "thumbs up", "thumbs down", "clap",
            "need a break", "away", "admit", "remove", "waiting room",
            "rename", "pin", "spotlight", "host", "co-host", "mute all",
            "unmute all", "allow unmute", "lock meeting", "settings",
            "audio settings", "video settings", "view", "gallery view",
            "speaker view", "minimize", "maximize", "close", "zoom",
            "everyone", "all panelists", "panelists and attendees"
        };

        private static readonly string[] _suffixesToStrip = new[]
        {
            ", Computer audio muted,Video off",
            ", Computer audio unmuted,Video off",
            ", Computer audio muted,Video on",
            ", Computer audio unmuted,Video on",
            ", Computer audio muted",
            ", Computer audio unmuted",
            ",Video off",
            ",Video on",
            " (Host)",
            " (Co-host)",
            " (me)",
            " is speaking"
        };

        public void Start()
        {
            _pollTimer = new System.Threading.Timer(_ => Poll(), null, 0, PollIntervalMs);
        }

        public void Stop()
        {
            _pollTimer?.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite);
        }

        private void Poll()
        {
            try
            {
                var zoomProcs = Process.GetProcessesByName("Zoom")
                    .Where(p => p.MainWindowHandle != IntPtr.Zero)
                    .ToList();

                if (zoomProcs.Count == 0) return;

                var proc = zoomProcs.FirstOrDefault(p =>
                               p.MainWindowTitle.Contains("Meeting", StringComparison.OrdinalIgnoreCase))
                           ?? zoomProcs.First();

                var app     = FlaUI.Core.Application.Attach(proc);
                var windows = app.GetAllTopLevelWindows(_automation);

                var meetingWindow = windows.FirstOrDefault(w =>
                    w.ClassName is "ZPContentViewWndClass"
                                or "VideoFrameWnd"
                                or "ZPFloatVideoWnd"
                                or "ZPMainFrame"
                                or "ConfMultiTabContentWndClass"
                );

                var window = meetingWindow ?? app.GetMainWindow(_automation);
                if (window == null) return;
                if (window.Properties.ProcessId != proc.Id) return;

                DetectActiveSpeaker(window);
                DetectParticipants(window);
            }
            catch { }
        }

        private void DetectActiveSpeaker(AutomationElement window)
        {
            try
            {
                string? speaker = null;
                var all = window.FindAllDescendants();
                foreach (var el in all)
                {
                    var name = el.Name ?? "";
                    if (name.EndsWith(" is speaking", StringComparison.OrdinalIgnoreCase))
                    {
                        speaker = StripSuffixes(name[..^" is speaking".Length].Trim());
                        break;
                    }
                }
                if (speaker != _lastSpeaker)
                {
                    _lastSpeaker = speaker;
                    ActiveSpeakerChanged?.Invoke(speaker);
                }
            }
            catch { }
        }

        private void DetectParticipants(AutomationElement window)
        {
            try
            {
                var plistPanel = window.FindFirstDescendant(
                    cf => cf.ByClassName("zPlistWndClass")
                );

                var searchRoot = plistPanel ?? window;

                var lists = searchRoot.FindAllDescendants(
                    cf => cf.ByControlType(ControlType.List)
                );

                List<string> candidates = new();

                foreach (var list in lists)
                {
                    var items = list.FindAllChildren(
                        cf => cf.ByControlType(ControlType.ListItem)
                    );
                    var itemNames = items
                        .Select(i => NormalizeParticipantName(i.Name?.Trim() ?? ""))
                        .Where(n => n != null)
                        .Select(n => n!)
                        .ToList();

                    if (itemNames.Count >= 1)
                        candidates.AddRange(itemNames);
                }

                var names = candidates
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(n => n)
                    .ToArray();

                if (names.Length == 0)
                    names = DetectFromVideoTiles(plistPanel ?? window);

                if (!names.SequenceEqual(_lastParticipants, StringComparer.OrdinalIgnoreCase))
                {
                    _lastParticipants = names;
                    ParticipantsUpdated?.Invoke(names);
                }
            }
            catch { }
        }

        private string[] DetectFromVideoTiles(AutomationElement root)
        {
            try
            {
                var textEls = root.FindAllDescendants(
                    cf => cf.ByControlType(ControlType.Text)
                );
                return textEls
                    .Select(e => NormalizeParticipantName(e.Name?.Trim() ?? ""))
                    .Where(n => n != null)
                    .Select(n => n!)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(n => n)
                    .ToArray();
            }
            catch { return Array.Empty<string>(); }
        }

        private static string StripSuffixes(string name)
        {
            bool changed = true;
            while (changed)
            {
                changed = false;
                foreach (var suffix in _suffixesToStrip)
                {
                    if (name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                    {
                        name    = name[..^suffix.Length].Trim();
                        changed = true;
                    }
                }
                // Also strip anything after a comma that looks like status info
                var commaIdx = name.LastIndexOf(",(", StringComparison.Ordinal);
                if (commaIdx > 0)
                {
                    name    = name[..commaIdx].Trim();
                    changed = true;
                }
            }
            return name;
        }

        private static string? NormalizeParticipantName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;

            name = StripSuffixes(name.Trim());

            if (name.Length < 2) return null;
            if (name.Length > 80) return null;
            if (_blocklist.Contains(name)) return null;

            // Filter meeting duration strings
            if (name.StartsWith("Meeting ", StringComparison.OrdinalIgnoreCase)) return null;

            // Filter time strings like "00:02:08"
            if (Regex.IsMatch(name, @"^\d{2}:\d{2}")) return null;

            // Filter yourself — "(me)" or "(Host, me)"
            if (name.Contains(", me)", StringComparison.OrdinalIgnoreCase)) return null;
            if (name.EndsWith("(me)", StringComparison.OrdinalIgnoreCase)) return null;

            // Filter digit-only starts
            if (char.IsDigit(name[0]) && name.Split(' ').Length <= 2) return null;

            // Must contain at least one letter including Hebrew
            bool hasLetter = name.Any(c => char.IsLetter(c) || (c >= '\u0590' && c <= '\u05FF'));
            if (!hasLetter) return null;

            if (name.Contains('\n') || name.Contains('\r') || name.Contains('\t')) return null;

            return name;
        }

        public void Dispose()
        {
            Stop();
            _pollTimer?.Dispose();
            _automation.Dispose();
        }
    }
}