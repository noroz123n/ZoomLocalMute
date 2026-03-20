"""
ZoomLocalMute.ai — Python Audio Backend
========================================
Captures Zoom's audio output, identifies speakers using pyannote's
speaker embedding model, and selectively mutes individual participants
locally (only you can't hear them — everyone else still can).

Architecture:
  Zoom → CABLE Input (VB-Audio) → This script → Your headphones

Features:
  - Passive voice clustering: automatically learns voices in the background
  - Manual enrollment: hold a button in the UI while someone speaks
  - Interruption detection: handles multiple voices during enrollment
  - Voice gallery: save/load voice profiles across meetings
  - Per-speaker muting with smooth fade in/out
  - WebSocket server on port 5799 for C# UI communication

Requirements:
  pip install pyannote.audio sounddevice websockets scipy torch torchaudio
  HuggingFace token required — run: huggingface-cli login
  Accept model license at: https://huggingface.co/pyannote/embedding

Usage:
  python audio_backend.py
"""

import asyncio
import json
import logging
import collections
import threading
import time
import io
import base64
import numpy as np
import sounddevice as sd
import torch
import torch.nn.functional as F
import websockets
from dataclasses import dataclass, field
from typing import Optional
from scipy.io import wavfile

# ── Logging ───────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s"
)
log = logging.getLogger("ZoomLocalMute.ai")

# ── Constants ─────────────────────────────────────────────────────────────────

# Audio settings
SAMPLE_RATE = 16000  # Hz — pyannote embedding expects 16kHz
BLOCK_SIZE  = 512    # samples per audio callback (~32ms per block)
WS_PORT     = 5799   # WebSocket port for C# UI communication

# Enrollment
ENROLL_SECONDS = 5.0   # seconds of speech needed to build voice profile
BUFFER_SECONDS = 30.0  # rolling audio buffer per enrolled speaker

# Speaker identification
# Cosine similarity threshold for voice matching.
# pyannote embedding-3.1 is more accurate than ECAPA-TDNN so scores are higher.
# Same person typically: 0.5-0.9. Different people typically: 0.0-0.3.
SIMILARITY_THRESHOLD = 0.3
FADE_STEPS           = 12   # blocks for smooth mute/unmute fade (~384ms)

# Voice Activity Detection
VAD_THRESHOLD = 0.002  # RMS energy above this = speech, below = silence

# Passive clustering
CLUSTER_MIN_SECONDS  = 1.5   # collect this many seconds before running identification
CLUSTER_MAX_CLUSTERS = 20
CLUSTER_MERGE_THRESH = 0.75  # slightly higher since pyannote embeddings are more discriminative

# Identification heartbeat — force a cluster if none fired recently
HEARTBEAT_SECONDS = 2.0

# Interruption detection during manual enrollment
INTERRUPT_WINDOW_SEC = 0.5
INTERRUPT_THRESH     = 0.35  # higher than ECAPA since pyannote is more discriminative

# How long to keep last identified speaker active before reset
IDENTIFICATION_TIMEOUT = 6.0

# ── Data structures ───────────────────────────────────────────────────────────

@dataclass
class VoiceCluster:
    """
    An unidentified voice cluster from passive listening.
    Accumulates audio until matched to a participant name,
    then becomes an enrolled SpeakerProfile.
    """
    cluster_id:     int
    embedding:      torch.Tensor
    audio_sample:   np.ndarray
    total_seconds:  float        = 0.0
    candidate_name: Optional[str] = None


@dataclass
class SpeakerProfile:
    """
    An enrolled participant with a known voice embedding.
    Created via manual enrollment or auto-enrollment from passive clustering.
    """
    name:         str
    embedding:    Optional[torch.Tensor] = None  # voice fingerprint
    audio_sample: Optional[np.ndarray]  = None  # 3-sec clip for UI playback
    audio_buffer: collections.deque = field(
        default_factory=lambda: collections.deque(
            maxlen=int(BUFFER_SECONDS * SAMPLE_RATE)
        )
    )
    speech_seconds:     float = 0.0
    muted:              bool  = False
    is_speaking:        bool  = False
    gain:               float = 1.0
    confidence:         float = 0.0
    enroll_interrupted: bool  = False

    @property
    def is_enrolled(self) -> bool:
        return self.embedding is not None

    @property
    def enrollment_pct(self) -> int:
        return min(100, int(self.speech_seconds / ENROLL_SECONDS * 100))


# ── Device detection ──────────────────────────────────────────────────────────

def find_best_input_device() -> tuple[int, str]:
    """
    Find the best audio input for capturing Zoom's output.
    Priority: VB-Audio CABLE > any virtual cable > Stereo Mix > system default
    """
    devices = sd.query_devices()
    inputs  = [(i, d) for i, d in enumerate(devices) if d["max_input_channels"] > 0]

    for i, d in inputs:
        n = d["name"].lower()
        if "cable output" in n or ("vb-audio" in n and "cable" in n):
            return i, d["name"]
    for i, d in inputs:
        if "virtual" in d["name"].lower() and "cable" in d["name"].lower():
            return i, d["name"]
    for i, d in inputs:
        n = d["name"].lower()
        if "stereo mix" in n or "what u hear" in n or "loopback" in n:
            return i, d["name"]

    idx  = sd.default.device[0]
    name = sd.query_devices(idx)["name"]
    log.warning(f"No virtual cable found. Using: {name}")
    log.warning("For best results: set Zoom Speaker to 'CABLE Input (VB-Audio)'")
    return idx, name


def find_best_output_device() -> tuple[int, str]:
    """
    Find best audio output (real headphones/speakers).
    Excludes virtual devices. Priority: headphones > Realtek > first real output.
    """
    devices = sd.query_devices()
    outputs = [(i, d) for i, d in enumerate(devices) if d["max_output_channels"] > 0]
    real    = [(i, d) for i, d in outputs
               if not any(kw in d["name"].lower()
                          for kw in ["virtual", "cable", "vb-audio"])]

    for i, d in real:
        if any(kw in d["name"].lower()
               for kw in ["headphone", "headset", "earphone", "bluetooth"]):
            return i, d["name"]
    for i, d in real:
        if "realtek" in d["name"].lower():
            return i, d["name"]
    if real:
        return real[0][0], real[0][1]["name"]

    idx  = sd.default.device[1]
    name = sd.query_devices(idx)["name"]
    return idx, name


def list_devices_for_ui() -> dict:
    """Return all audio devices with recommendations for the UI device selector."""
    devices = sd.query_devices()
    inputs  = [{"index": i, "name": d["name"]}
               for i, d in enumerate(devices) if d["max_input_channels"]  > 0]
    outputs = [{"index": i, "name": d["name"]}
               for i, d in enumerate(devices) if d["max_output_channels"] > 0]
    bi, bn  = find_best_input_device()
    bo, on  = find_best_output_device()
    return {
        "inputs":           inputs,
        "outputs":          outputs,
        "best_input_idx":   bi,
        "best_output_idx":  bo,
        "best_input_name":  bn,
        "best_output_name": on,
    }


# ── Main Backend ──────────────────────────────────────────────────────────────

class AudioBackend:
    """
    Core audio engine. Manages the pyannote embedding model, audio stream,
    speaker profiles, and WebSocket communication with the C# UI.
    """

    def __init__(self):
        self.speakers: dict[str, SpeakerProfile] = {}
        self.clients   = set()
        self._loop     = None
        self._stream   = None
        self.running   = False

        self._enroll_lock    = threading.Lock()
        self._embedding_lock = threading.Lock()

        self._active_speaker: Optional[str] = None

        self._clusters:    list[VoiceCluster] = []
        self._cluster_id   = 0

        self._enrolling_name: Optional[str]    = None
        self._enroll_buffer:  list[np.ndarray] = []

        self._block_buffer:  list[np.ndarray] = []
        self._block_buf_max  = int(CLUSTER_MIN_SECONDS * SAMPLE_RATE / BLOCK_SIZE) + 1

        self._last_identified_speaker:  Optional[str] = None
        self._last_identification_time: float         = 0.0
        self._last_cluster_fire_time:   float         = 0.0

        log.info("Loading pyannote speaker embedding model...")
        self._load_model()
        log.info(f"Model ready. Device: {'CUDA' if torch.cuda.is_available() else 'CPU'}")

    # ── Model ─────────────────────────────────────────────────────────────────

    def _load_model(self):
        from pyannote.audio import Model
        from pyannote.audio import Inference
        import torch
        import pytorch_lightning as pl
        import omegaconf

        # Add the specific classes PyTorch is blocking to the safe list
        # We include EarlyStopping and OmegaConf configs to be safe
        torch.serialization.add_safe_globals([
            pl.callbacks.early_stopping.EarlyStopping,
            pl.callbacks.model_checkpoint.ModelCheckpoint,
            omegaconf.listconfig.ListConfig,
            omegaconf.dictconfig.DictConfig
        ])

        device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
        log.info(f"Loading pyannote model on {device} with security allowlist...")

        try:
            # We wrap the load in a context manager that allows 'unsafe' globals
            # just for this specific operation.
            with torch.serialization.safe_globals([
                pl.callbacks.early_stopping.EarlyStopping,
                omegaconf.listconfig.ListConfig
            ]):
                model = Model.from_pretrained(
                    "pyannote/embedding", 
                    token=True
                )

            if model is None:
                raise ValueError("Model load returned None. Check HF token.")

            self.encoder = Inference(model, window="whole")
            self.encoder.to(device)
            self._device = device
            log.info("Model loaded successfully!")

        except Exception as e:
            log.error(f"Final loading attempt failed: {e}")
            # If even that fails, we use the 'Nuclear Option' for PyTorch 2.6
            log.info("Attempting nuclear bypass...")
            import torch.serialization
            # Manually override the default behavior for this session
            old_load = torch.load
            torch.load = lambda *args, **kwargs: old_load(*args, **{**kwargs, "weights_only": False})
            model = Model.from_pretrained("pyannote/embedding", token=True)
            self.encoder = Inference(model, window="whole")
            self.encoder.to(device)
            self._device = device
            log.info("Model loaded via bypass.")

    # ── Embeddings ────────────────────────────────────────────────────────────

    def _get_embedding(self, audio: np.ndarray, silent: bool = False) -> Optional[torch.Tensor]:
        """
        Compute a speaker embedding from audio samples.

        Args:
            audio:  float32 numpy array at SAMPLE_RATE Hz
            silent: suppress "too short" warnings (used internally)

        Returns:
            Normalized embedding tensor, or None if audio is invalid.
        """
        try:
            # 3 seconds is enough for a reliable embedding and keeps CPU fast
            max_samples = int(3.0 * SAMPLE_RATE)
            if len(audio) > max_samples:
                start = (len(audio) - max_samples) // 2
                audio = audio[start:start + max_samples]

            if len(audio) < SAMPLE_RATE:
                if not silent:
                    log.warning(f"Audio too short: {len(audio)/SAMPLE_RATE:.2f}s")
                return None

            # pyannote Inference expects a dict with waveform tensor + sample rate
            waveform = torch.tensor(audio, dtype=torch.float32).unsqueeze(0)

            with self._embedding_lock:
                with torch.no_grad():
                    emb = self.encoder({"waveform": waveform, "sample_rate": SAMPLE_RATE})

            # emb is a numpy array from pyannote — convert to normalized tensor
            emb_tensor = torch.tensor(emb, dtype=torch.float32).squeeze()
            return F.normalize(emb_tensor, dim=-1)

        except Exception as e:
            log.warning(f"Embedding error: {type(e).__name__}: {e}")
            return None

    def _similarity(self, a: torch.Tensor, b: torch.Tensor) -> float:
        """Cosine similarity between two embeddings (-1 to 1)."""
        return F.cosine_similarity(a.unsqueeze(0), b.unsqueeze(0)).item()

    def _embedding_to_b64(self, emb: torch.Tensor) -> str:
        """Serialize embedding to base64 for storage."""
        buf = io.BytesIO()
        torch.save(emb.cpu(), buf)
        return base64.b64encode(buf.getvalue()).decode("utf-8")

    def _b64_to_embedding(self, b64: str) -> torch.Tensor:
        """Deserialize embedding from base64."""
        buf = io.BytesIO(base64.b64decode(b64))
        return F.normalize(torch.load(buf, weights_only=True), dim=-1)

    # ── Passive clustering ────────────────────────────────────────────────────

    def _process_cluster(self, audio: np.ndarray):
        """
        Process a speech segment for speaker identification.
        Called from background threads every ~1.5-2 seconds during speech.

        1. Compute embedding
        2. Match against enrolled speakers
        3. If no match: merge with existing cluster or create new one
        4. Auto-enroll clusters once they have enough audio + candidate name
        """
        self._last_cluster_fire_time = time.time()

        emb = self._get_embedding(audio)
        if emb is None:
            return

        # Match against enrolled speakers
        best_name  = None
        best_score = SIMILARITY_THRESHOLD
        enrolled   = {n: s for n, s in self.speakers.items() if s.is_enrolled}

        for name, sp in enrolled.items():
            score = self._similarity(emb, sp.embedding)
            if score > best_score:
                best_score = score
                best_name  = name

        if enrolled:
            scores = {n: round(self._similarity(emb, s.embedding), 3)
                      for n, s in enrolled.items()}
            log.info(f"Scores: {scores} | match={best_name}")

        if best_name:
            for n, s in self.speakers.items():
                s.is_speaking = (n == best_name)

            sp = self.speakers[best_name]
            sp.audio_buffer.extend(audio.tolist())
            sp.speech_seconds              += len(audio) / SAMPLE_RATE
            sp.confidence                   = best_score
            self._last_identification_time  = time.time()
            self._last_identified_speaker   = best_name
            log.info(f"✅ Identified: {best_name} ({best_score:.2f})")

            if sp.speech_seconds % 15 < (len(audio) / SAMPLE_RATE):
                threading.Thread(
                    target=self._refresh_embedding, args=(sp,), daemon=True
                ).start()
            return

        # No match — keep _last_identified_speaker active.
        # The IDENTIFICATION_TIMEOUT in _process_block handles natural pauses.
        for s in self.speakers.values():
            s.is_speaking = False

        # Try existing clusters
        best_cluster = None
        best_cscore  = CLUSTER_MERGE_THRESH
        for cluster in self._clusters:
            score = self._similarity(emb, cluster.embedding)
            if score > best_cscore:
                best_cscore  = score
                best_cluster = cluster

        if best_cluster:
            best_cluster.embedding     = F.normalize(
                0.7 * best_cluster.embedding + 0.3 * emb, dim=-1
            )
            best_cluster.total_seconds += len(audio) / SAMPLE_RATE
            if len(audio) >= int(3.0 * SAMPLE_RATE):
                best_cluster.audio_sample = audio[:int(3.0 * SAMPLE_RATE)]
            if self._active_speaker and not best_cluster.candidate_name:
                best_cluster.candidate_name = self._active_speaker

            if (best_cluster.total_seconds >= ENROLL_SECONDS
                    and best_cluster.candidate_name
                    and best_cluster.candidate_name in self.speakers):
                sp = self.speakers[best_cluster.candidate_name]
                if not sp.is_enrolled:
                    sp.embedding    = best_cluster.embedding
                    sp.audio_sample = best_cluster.audio_sample
                    sp.confidence   = 0.8
                    log.info(f"Auto-enrolled: {sp.name}")
                    if self._loop:
                        asyncio.run_coroutine_threadsafe(
                            self._broadcast({
                                "type":    "auto_enrolled",
                                "name":    sp.name,
                                "message": f"✅ Auto-enrolled: {sp.name}"
                            }),
                            self._loop
                        )
                    self._clusters.remove(best_cluster)
        else:
            if len(self._clusters) < CLUSTER_MAX_CLUSTERS:
                sample  = (audio[:int(3.0 * SAMPLE_RATE)]
                           if len(audio) >= int(3.0 * SAMPLE_RATE) else audio)
                cluster = VoiceCluster(
                    cluster_id     = self._cluster_id,
                    embedding      = emb,
                    audio_sample   = sample,
                    total_seconds  = len(audio) / SAMPLE_RATE,
                    candidate_name = self._active_speaker
                )
                self._clusters.append(cluster)
                self._cluster_id += 1

    def _refresh_embedding(self, sp: SpeakerProfile):
        """
        Blend new audio into speaker's embedding (running average).
        Keeps identification robust as speech patterns vary across the meeting.
        """
        with self._enroll_lock:
            needed = int(ENROLL_SECONDS * SAMPLE_RATE)
            if len(sp.audio_buffer) < needed:
                return
            audio = np.array(list(sp.audio_buffer)[-needed:], dtype=np.float32)
            emb   = self._get_embedding(audio)
            if emb is None:
                return
            sp.embedding = F.normalize(0.7 * sp.embedding + 0.3 * emb, dim=-1)

    # ── Interruption detection ────────────────────────────────────────────────

    def _detect_interruption(self, audio: np.ndarray) -> tuple[bool, np.ndarray, Optional[np.ndarray]]:
        """
        Check enrollment recording for two distinct voices.
        If detected, extract and return the dominant (longer) voice.
        Only flags as interruption if second voice >= 10% of total duration.
        """
        window = int(INTERRUPT_WINDOW_SEC * SAMPLE_RATE)
        if len(audio) < window * 4:
            return False, audio, None

        segments = []
        step     = window // 2
        for i in range(0, len(audio) - window, step):
            seg = audio[i:i + window]
            if float(np.sqrt(np.mean(seg ** 2))) > VAD_THRESHOLD:
                emb = self._get_embedding(seg, silent=True)
                if emb is not None:
                    segments.append((i, emb, len(seg)))

        if len(segments) < 4:
            return False, audio, None

        ref_emb = segments[0][1]
        group_a = [segments[0]]
        group_b = []

        for seg in segments[1:]:
            if self._similarity(ref_emb, seg[1]) < INTERRUPT_THRESH:
                group_b.append(seg)
            else:
                group_a.append(seg)

        if not group_b:
            return False, audio, None

        dur_a = sum(s[2] for s in group_a)
        dur_b = sum(s[2] for s in group_b)
        if min(dur_a, dur_b) / (dur_a + dur_b) < 0.10:
            return False, audio, None

        dominant_group = group_a if dur_a >= dur_b else group_b
        dom_indices    = sorted([s[0] for s in dominant_group])
        dominant_audio = np.concatenate([
            audio[i:i + window]
            for i in dom_indices if i + window <= len(audio)
        ])

        log.info(f"Interruption detected: {dur_a/SAMPLE_RATE:.1f}s vs {dur_b/SAMPLE_RATE:.1f}s")
        return True, dominant_audio, None

    # ── Manual enrollment ─────────────────────────────────────────────────────

    def start_manual_enroll(self, name: str):
        """Start recording for manual enrollment (UI button pressed)."""
        self._enrolling_name = name
        self._enroll_buffer  = []
        log.info(f"Manual enrollment started: {name}")

    def stop_manual_enroll(self) -> dict:
        """
        Stop recording and process enrollment audio.
        Returns a result dict sent back to the C# UI containing
        success status, message, and serialized embedding for gallery storage.
        """
        name             = self._enrolling_name
        self._enrolling_name = None

        if not name or not self._enroll_buffer:
            return {
                "success": False, "name": name or "", "interrupted": False,
                "duration": 0.0, "embedding": None, "message": "No audio recorded."
            }

        audio    = np.concatenate(self._enroll_buffer).astype(np.float32)
        self._enroll_buffer = []
        duration = len(audio) / SAMPLE_RATE
        log.info(f"Enrollment audio: {duration:.1f}s for {name}")

        if duration < 1.5:
            return {
                "success": False, "name": name, "interrupted": False,
                "duration": round(duration, 1), "embedding": None,
                "message": "Too short — hold the button longer (at least 2 seconds)."
            }

        interrupted    = False
        dominant_audio = audio
        try:
            interrupted, detected, _ = self._detect_interruption(audio)
            if detected is not None and len(detected) >= SAMPLE_RATE:
                dominant_audio = detected
                log.info(f"Using dominant audio: {len(dominant_audio)/SAMPLE_RATE:.1f}s")
            else:
                interrupted = False
        except Exception as e:
            log.warning(f"Interruption detection failed: {e}")

        if name not in self.speakers:
            log.info(f"Creating profile on-the-fly for: {name}")
            self.speakers[name] = SpeakerProfile(name=name)

        sp  = self.speakers[name]
        emb = self._get_embedding(dominant_audio)
        if emb is None:
            emb = self._get_embedding(audio)
        if emb is None:
            return {
                "success": False, "name": name, "interrupted": False,
                "duration": round(duration, 1), "embedding": None,
                "message": "Could not process audio. Make sure there is clear speech."
            }

        # Average with existing embedding for robustness across multiple enrollments
        if sp.embedding is not None:
            sp.embedding = F.normalize(0.5 * sp.embedding + 0.5 * emb, dim=-1)
            log.info(f"Averaged embedding for {name}")
        else:
            sp.embedding = emb

        sample_src      = dominant_audio if len(dominant_audio) >= int(3.0 * SAMPLE_RATE) else audio
        sp.audio_sample = (sample_src[:int(3.0 * SAMPLE_RATE)]
                           if len(sample_src) >= int(3.0 * SAMPLE_RATE) else sample_src)
        sp.confidence         = 1.0
        sp.enroll_interrupted = interrupted
        sp.speech_seconds     = ENROLL_SECONDS

        try:
            emb_b64 = self._embedding_to_b64(sp.embedding)
        except Exception as e:
            log.error(f"Embedding serialization failed: {e}")
            emb_b64 = None

        log.info(f"✅ Enrolled {name} ({duration:.1f}s)")
        return {
            "success":     True,
            "name":        name,
            "interrupted": interrupted,
            "duration":    round(duration, 1),
            "embedding":   emb_b64,
            "message": (
                f"✅ Enrolled {name} ({duration:.1f}s)."
                + (" ⚠ Another voice detected — used dominant voice." if interrupted else "")
            )
        }

    def delete_enrollment(self, name: str):
        """Reset a speaker's voice profile to unenrolled state."""
        if name in self.speakers:
            sp = self.speakers[name]
            sp.embedding = sp.audio_sample = None
            sp.confidence = sp.speech_seconds = 0.0
            sp.enroll_interrupted = False
            sp.audio_buffer.clear()
            log.info(f"Enrollment deleted: {name}")

    def load_saved_voice(self, participant_name: str, saved_name: str, b64_embedding: str):
        """Apply a saved voice profile from the gallery to a participant."""
        if participant_name not in self.speakers:
            log.warning(f"Cannot load saved voice — {participant_name} not in participants")
            return
        try:
            emb = self._b64_to_embedding(b64_embedding)
            sp  = self.speakers[participant_name]
            sp.embedding = emb
            sp.confidence = sp.speech_seconds = ENROLL_SECONDS
            log.info(f"Loaded saved voice '{saved_name}' → {participant_name}")
        except Exception as e:
            log.warning(f"Failed to load saved voice: {e}")

    def get_audio_sample(self, name: str) -> Optional[str]:
        """Return a speaker's 3-second voice sample as base64 WAV (for UI Listen button)."""
        if name not in self.speakers or self.speakers[name].audio_sample is None:
            return None
        try:
            buf = io.BytesIO()
            wavfile.write(buf, SAMPLE_RATE,
                          (self.speakers[name].audio_sample * 32767).astype(np.int16))
            return base64.b64encode(buf.getvalue()).decode("utf-8")
        except Exception as e:
            log.warning(f"Audio sample error: {e}")
            return None

    # ── Audio processing ──────────────────────────────────────────────────────

    def _process_block(self, audio: np.ndarray) -> np.ndarray:
        """
        Real-time audio processing — called every ~32ms.

        1. Feed audio to manual enrollment buffer if recording
        2. Accumulate speech for passive clustering / identification
        3. Fire identification heartbeat every HEARTBEAT_SECONDS
        4. Apply per-speaker mute gain based on last identification
        5. Return processed audio (possibly silenced / faded)
        """
        rms = float(np.sqrt(np.mean(audio ** 2)))

        if self._enrolling_name:
            self._enroll_buffer.append(audio.copy())

        min_flush_blocks = int(1.0 * SAMPLE_RATE / BLOCK_SIZE)

        if rms > VAD_THRESHOLD:
            self._block_buffer.append(audio.copy())

            if len(self._block_buffer) >= self._block_buf_max:
                self._fire_cluster()
            elif (len(self._block_buffer) > min_flush_blocks
                  and time.time() - self._last_cluster_fire_time > HEARTBEAT_SECONDS):
                self._fire_cluster()
        else:
            if len(self._block_buffer) >= min_flush_blocks:
                self._fire_cluster()
            else:
                self._block_buffer = []

        # Reset stale identification after prolonged silence
        if time.time() - self._last_identification_time > IDENTIFICATION_TIMEOUT:
            for sp in self.speakers.values():
                sp.is_speaking = False
            self._last_identified_speaker = None

        # Apply per-speaker mute gain
        output    = audio.copy()
        active_sp = self.speakers.get(self._last_identified_speaker) \
                    if self._last_identified_speaker else None

        if active_sp and active_sp.is_enrolled:
            target         = 0.0 if active_sp.muted else 1.0
            step           = 1.0 / FADE_STEPS
            active_sp.gain = (
                min(target, active_sp.gain + step) if active_sp.gain < target
                else max(target, active_sp.gain - step)
            )
            output = output * active_sp.gain
        else:
            for sp in self.speakers.values():
                if sp.gain < 1.0:
                    sp.gain = min(1.0, sp.gain + 1.0 / FADE_STEPS)

        return output

    def _fire_cluster(self):
        """Flush the block buffer and run identification in a background thread."""
        if not self._block_buffer:
            return
        segment            = np.concatenate(self._block_buffer)
        self._block_buffer = []
        self._last_cluster_fire_time = time.time()
        threading.Thread(
            target=self._process_cluster, args=(segment,), daemon=True
        ).start()

    # ── Stream ────────────────────────────────────────────────────────────────

    def start_stream(self, input_idx: Optional[int] = None,
                           output_idx: Optional[int] = None):
        """Start the duplex audio stream (Zoom output → processing → headphones)."""
        self.stop_stream()

        if input_idx is None:
            input_idx,  in_name  = find_best_input_device()
            log.info(f"Input:  {in_name}")
        if output_idx is None:
            output_idx, out_name = find_best_output_device()
            log.info(f"Output: {out_name}")

        def callback(indata, outdata, frames, time_info, status):
            audio         = indata[:, 0].astype(np.float32)
            processed     = self._process_block(audio)
            outdata[:, 0] = processed
            if outdata.shape[1] > 1:
                outdata[:, 1] = processed

        self._stream = sd.Stream(
            samplerate = SAMPLE_RATE,
            blocksize  = BLOCK_SIZE,
            dtype      = "float32",
            channels   = (1, 2),
            device     = (input_idx, output_idx),
            callback   = callback,
            latency    = "low"
        )
        self._stream.start()
        self.running = True
        log.info(f"Stream started: In={input_idx}, Out={output_idx}")

    def stop_stream(self):
        """Stop and clean up the audio stream."""
        if self._stream:
            self._stream.stop()
            self._stream.close()
            self._stream = None
            self.running = False
            log.info("Stream stopped.")

    # ── Participant management ────────────────────────────────────────────────

    def update_participants(self, names: list):
        """Sync participant list with Zoom's current meeting."""
        current = set(self.speakers.keys())
        new_set = set(names)
        for name in new_set - current:
            self.speakers[name] = SpeakerProfile(name=name)
            log.info(f"Added: {name}")
        for name in current - new_set:
            del self.speakers[name]
            log.info(f"Removed: {name}")

    def set_muted(self, name: str, muted: bool):
        """Set local mute state for a participant."""
        if name in self.speakers:
            self.speakers[name].muted = muted
            log.info(f"{'Muted' if muted else 'Unmuted'}: {name}")

    def set_active_speaker(self, name: Optional[str]):
        """Receive Zoom's active speaker hint from C# UI Automation."""
        self._active_speaker = name
        if name:
            for cluster in self._clusters:
                if not cluster.candidate_name:
                    cluster.candidate_name = name

    def _speaker_status(self) -> list:
        """Serialize current speaker states for the UI."""
        return [
            {
                "name":           s.name,
                "muted":          s.muted,
                "enrolled":       s.is_enrolled,
                "enrollment_pct": s.enrollment_pct,
                "is_speaking":    s.is_speaking,
                "confidence":     round(s.confidence * 100),
                "interrupted":    s.enroll_interrupted,
                "has_sample":     s.audio_sample is not None,
            }
            for s in self.speakers.values()
        ]

    # ── WebSocket ─────────────────────────────────────────────────────────────

    async def _broadcast(self, msg: dict):
        """Send a message to all connected UI clients."""
        if self.clients:
            data = json.dumps(msg)
            await asyncio.gather(
                *[c.send(data) for c in self.clients],
                return_exceptions=True
            )

    async def _handle_client(self, ws):
        """Handle a WebSocket connection from the C# UI."""
        self.clients.add(ws)
        log.info(f"UI connected: {ws.remote_address}")
        try:
            await ws.send(json.dumps({"type": "devices",  "data": list_devices_for_ui()}))
            await ws.send(json.dumps({"type": "speakers", "data": self._speaker_status()}))

            if not self.running:
                self.start_stream()
                await ws.send(json.dumps({
                    "type": "status", "message": "▶ Audio filtering active."
                }))

            async for message in ws:
                try:
                    await self._handle_message(json.loads(message))
                except json.JSONDecodeError:
                    pass

        except websockets.ConnectionClosed:
            pass
        finally:
            self.clients.discard(ws)
            if not self.clients and self.running:
                self.stop_stream()
            log.info("UI disconnected.")

    async def _handle_message(self, msg: dict):
        """Route incoming messages from the C# UI."""
        t = msg.get("type")

        if t == "start_audio":
            self.start_stream(msg.get("input_device"), msg.get("output_device"))
            await self._broadcast({"type": "status", "message": "▶ Audio filtering active."})

        elif t == "stop_audio":
            self.stop_stream()
            await self._broadcast({"type": "status", "message": "⏹ Stopped."})

        elif t == "update_participants":
            self.update_participants(msg.get("names", []))
            await self._broadcast({"type": "speakers", "data": self._speaker_status()})

        elif t == "set_muted":
            self.set_muted(msg.get("name", ""), msg.get("muted", False))
            await self._broadcast({"type": "speakers", "data": self._speaker_status()})

        elif t == "set_active_speaker":
            self.set_active_speaker(msg.get("name"))

        elif t == "start_enroll":
            self.start_manual_enroll(msg.get("name", ""))
            await self._broadcast({
                "type":    "status",
                "message": f"🎤 Recording {msg.get('name')}... release when done."
            })

        elif t == "stop_enroll":
            result = self.stop_manual_enroll()
            await self._broadcast({"type": "enroll_result", "data": result})
            await self._broadcast({"type": "speakers",      "data": self._speaker_status()})
            await self._broadcast({"type": "status",        "message": result["message"]})

        elif t == "delete_enrollment":
            self.delete_enrollment(msg.get("name", ""))
            await self._broadcast({"type": "speakers", "data": self._speaker_status()})
            await self._broadcast({
                "type":    "status",
                "message": f"🗑 Enrollment deleted for {msg.get('name')}."
            })

        elif t == "load_saved_voice":
            name      = msg.get("name", "")
            saved     = msg.get("saved_name", name)
            embedding = msg.get("embedding", "")
            if embedding:
                self.load_saved_voice(name, saved, embedding)
                await self._broadcast({"type": "speakers", "data": self._speaker_status()})
                await self._broadcast({
                    "type": "status", "message": f"💾 Loaded saved voice for {name}."
                })

        elif t == "get_sample":
            name   = msg.get("name", "")
            sample = self.get_audio_sample(name)
            await self._broadcast({
                "type":  "audio_sample",
                "name":  name,
                "data":  sample,
                "found": sample is not None
            })

        elif t == "get_devices":
            await self._broadcast({"type": "devices", "data": list_devices_for_ui()})

        elif t == "get_status":
            await self._broadcast({"type": "speakers", "data": self._speaker_status()})

    async def _status_loop(self):
        """Push speaker status to UI every 500ms for live updates."""
        while True:
            await asyncio.sleep(0.5)
            if self.clients and self.speakers:
                await self._broadcast({
                    "type": "speakers",
                    "data": self._speaker_status()
                })

    async def run(self):
        """Start the WebSocket server and run forever."""
        self._loop = asyncio.get_event_loop()
        log.info(f"Starting WebSocket server on ws://localhost:{WS_PORT}")
        async with websockets.serve(self._handle_client, "localhost", WS_PORT):
            await asyncio.gather(self._status_loop(), asyncio.Future())


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    backend = AudioBackend()
    try:
        asyncio.run(backend.run())
    except KeyboardInterrupt:
        backend.stop_stream()
        log.info("Shutdown.")