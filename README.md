# ZoomLocalMute.ai

**Locally mute individual Zoom participants using AI speaker identification.**

Zoom doesn't let you silence a specific person for yourself only — their audio is all-or-nothing. ZoomLocalMute.ai intercepts Zoom's audio output, identifies who is speaking using a neural speaker embedding model, and selectively silences them on your end only. Everyone else in the meeting still hears them normally.

---

## How It Works

```
Zoom ──► CABLE Input (VB-Audio) ──► Python backend ──► Your headphones
                                          │
                                    Speaker ID (AI)
                                    Per-speaker mute
```

1. Zoom's audio is routed through a virtual audio cable (VB-Audio)
2. The Python backend captures the audio stream in real time
3. pyannote's speaker embedding model identifies who is speaking every ~2 seconds
4. If a muted participant is identified as speaking, their audio is silenced locally
5. The C# UI shows all participants with enrollment and mute controls

---

## Features

| | |
|---|---|
| 🎤 Hold to Enroll | Hold button while participant speaks to build their voice profile |
| 🔇 Mute for me | Silence a participant locally — others still hear them |
| ▶ Listen | Play back recorded voice sample to verify correct enrollment |
| 💾 Save to gallery | Persist voice profiles across meetings |
| 📂 Gallery | Auto-apply saved voices when the same person joins future meetings |

---

## Requirements

- Windows 10/11
- [VB-Audio Virtual Cable](https://vb-audio.com/Cable/) (free)
- .NET 8.0 SDK
- Python 3.10+
- NVIDIA GPU recommended (runs on CPU but slower)

### Zoom Setup
In Zoom Settings → Audio, set **Speaker** to `CABLE Input (VB-Audio Virtual Cable)`.

In Windows Sound Settings, do **not** enable "Listen to this device" for CABLE Output — the Python backend handles output directly to your headphones.

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/yourusername/ZoomLocalMuteAI.git
cd ZoomLocalMuteAI
```

### 2. Set up Python environment

```bash
python -m venv venv
venv\Scripts\activate
```

Install PyTorch with GPU support (recommended):
```bash
pip install torch==2.3.0+cu121 torchaudio==2.3.0+cu121 --index-url https://download.pytorch.org/whl/cu121
```

Or CPU only:
```bash
pip install torch torchaudio
```

Install remaining dependencies:
```bash
pip install -r requirements.txt
```

### 3. Set up HuggingFace access

The speaker model requires a free HuggingFace account:

1. Create an account at [huggingface.co](https://huggingface.co)
2. Accept the license at [pyannote/embedding](https://huggingface.co/pyannote/embedding)
3. Generate a token at [huggingface.co/settings/tokens](https://huggingface.co/settings/tokens)
4. Login from within the venv:

```bash
huggingface-cli login
```

The model (~200MB) downloads automatically on first run.

### 4. Build the C# UI

```bash
dotnet build
```

---

## Usage

### Start the backend

```bash
venv\Scripts\activate
python audio_backend.py
```

### Start the UI

```bash
dotnet run
```

### Enrolling a participant

1. Join a Zoom meeting and open Zoom's **Participants panel**
2. In ZoomLocalMute, find the participant's row
3. Hold **🎤 Hold to Enroll** while they are speaking
4. Release after 3–5 seconds
5. Click **▶ Listen** to verify it recorded the right person
6. Click **💾** to save the profile for future meetings

Enroll the same person 2–3 times — each enrollment is averaged into the profile for better accuracy.

### Muting

Once enrolled, click **Mute for me**. The system silences that participant whenever their voice is identified. Click again to unmute.

---

## Project Structure

```
ZoomLocalMute/
├── audio_backend.py          # Python audio engine + WebSocket server
├── requirements.txt          # Python dependencies
├── ZoomLocalMute.csproj      # C# project file
└── src/
    ├── Program.cs            # Entry point
    ├── AudioEngine.cs        # WebSocket client + speaker state management
    ├── ZoomSpeakerTracker.cs # Zoom UI Automation (reads participant names)
    ├── TrayApplication.cs    # WinForms UI + system tray
    └── VoiceGallery.cs       # Voice profile persistence
```

`pretrained_models/` and `venv/` are excluded — see `.gitignore`. The model downloads automatically on first run.

---

## Architecture

### Python Backend

- **Audio stream**: `sounddevice` duplex stream — reads from VB-Audio CABLE Output, writes to headphones
- **Speaker identification**: [pyannote/embedding](https://huggingface.co/pyannote/embedding) — trained on diverse real-world audio including noisy/VoIP conditions
- **Passive clustering**: accumulates ~1.5s speech segments, computes embeddings, clusters unknown voices and auto-enrolls when a participant name is available
- **Manual enrollment**: records while button is held, detects interruptions to extract the dominant voice, averages embeddings across enrollments
- **WebSocket server**: port 5799, JSON protocol to/from C# UI

### C# UI

- **ZoomSpeakerTracker**: polls Zoom's accessibility tree every 600ms using FlaUI/UI Automation to read participant names
- **AudioEngine**: WebSocket client — sends participant list and mute commands, receives speaker status and enrollment results
- **ControlWindow**: WinForms UI with per-participant rows showing confidence score, enrollment progress, and controls
- **VoiceGallery**: saves voice profiles as JSON to `%AppData%\ZoomLocalMute\voices.json`

---

## Speaker Identification Notes

Zoom's adaptive bitrate codec (Opus) compresses audio differently depending on network conditions and speaker priority, which causes similarity scores to vary for the same person. Enrolling multiple times helps the model adapt.

| Condition | Typical score |
|---|---|
| Same person (good conditions) | 0.50 – 0.80 |
| Same person (Zoom compression) | 0.30 – 0.60 |
| Different people | 0.05 – 0.25 |

Current threshold: `0.45` — adjustable via `SIMILARITY_THRESHOLD` in `audio_backend.py`.

---

## Known Limitations

- **Windows only** — uses Windows UI Automation to read Zoom's participant list
- **VB-Audio required** — must be installed separately, cannot be bundled
- **Zoom compression** — Zoom's adaptive bitrate codec reduces identification accuracy vs. clean audio
- **Zoom ToS** — reads Zoom's accessibility tree (same API as screen readers). Use at your own discretion.
- **CPU latency** — ~200ms per embedding on CPU; a GPU significantly improves responsiveness

---

## Dependencies

### Python
| Package | License |
|---|---|
| pyannote.audio | MIT |
| sounddevice | MIT |
| websockets | BSD |
| torch / torchaudio | BSD |
| scipy | BSD |
| numpy | BSD |

### C#
| Package | License |
|---|---|
| NAudio | MIT |
| FlaUI | MIT |

### External
| Tool | License |
|---|---|
| VB-Audio Virtual Cable | Freeware (install separately) |

---

## License

MIT — see [LICENSE](LICENSE)

---

## Potential Improvements

- Upgrade to [pyannote/embedding-3.1](https://huggingface.co/pyannote/embedding-3.1) for higher accuracy
- Volume slider — reduce a participant to X% instead of full silence
- macOS support — replace Windows UI Automation with macOS Accessibility API
- Auto-threshold — dynamically calibrate similarity threshold per meeting
