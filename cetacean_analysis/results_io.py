"""Single source of truth for cetacean-analysis reproducibility.

This module centralises:
- the dataset/analysis configuration (pinned dataset revision, target species,
  sampling parameters) so every script uses identical settings;
- a memory-safe audio decoder (bounded to ANALYSIS_SECONDS, polyphase resample)
  that runs within a standard 8 GB machine;
- construction, writing and loading of ``output/results.json`` — the machine
  readable results file that the manuscript/figure generators read so that NO
  numeric result is hard-coded in a document generator.

All numbers reported in the manuscripts (body text, tables, slides) are derived
from ``results.json`` via the formatting helpers below, so the chain
data -> analysis -> results.json -> manuscript is reproducible end to end.
"""

import io
import json
import os

import numpy as np

# ---------------------------------------------------------------------------
# Configuration (shared by every script)
# ---------------------------------------------------------------------------
DATASET_NAME = "confit/wmms-parquet"
# Pin the exact dataset revision so the corpus cannot silently change.
DATASET_REVISION = "a90a38e0006991f0c6f6d4e05261949a1da7f14e"

TARGET_SPECIES = [
    "Sperm_Whale",
    "Humpback_Whale",
    "Killer_Whale",
    "Fin,_Finback_Whale",
    "Bottlenose_Dolphin",
    "Beluga,_White_Whale",
]

MAX_SAMPLES = 10          # recordings analysed per species (first N in corpus order)
TARGET_SR = 16000         # Hz, standardised sampling rate
ANALYSIS_SECONDS = 60     # analyse the first N seconds of each recording (bounded memory / deterministic)

# Inter-click-interval (ICI) detection window (seconds). Intervals outside this
# window are discarded as spurious. This is a fixed analysis parameter, NOT a
# measured per-species result.
ICI_MIN_S = 0.001
ICI_MAX_S = 2.0

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "output")
RESULTS_PATH = os.path.join(OUTPUT_DIR, "results.json")

SPECIES_JA = {
    "Sperm_Whale": "マッコウクジラ",
    "Humpback_Whale": "ザトウクジラ",
    "Killer_Whale": "シャチ",
    "Fin,_Finback_Whale": "ナガスクジラ",
    "Bottlenose_Dolphin": "バンドウイルカ",
    "Beluga,_White_Whale": "シロイルカ",
}

SPECIES_EN = {
    "Sperm_Whale": "Sperm whale",
    "Humpback_Whale": "Humpback whale",
    "Killer_Whale": "Killer whale",
    "Fin,_Finback_Whale": "Fin whale",
    "Bottlenose_Dolphin": "Bottlenose dolphin",
    "Beluga,_White_Whale": "Beluga whale",
}


# ---------------------------------------------------------------------------
# Memory-safe audio decoding
# ---------------------------------------------------------------------------
def decode_audio(raw_audio_bytes, target_sr=TARGET_SR, max_seconds=ANALYSIS_SECONDS):
    """Decode raw audio bytes to a mono float64 array at ``target_sr``.

    Reads at most ``max_seconds`` of audio at the native rate (so a ~20-min
    recording no longer materialises ~50M samples), then resamples with a
    polyphase filter (``resample_poly``) which is far more memory efficient
    than the FFT-based ``scipy.signal.resample``.
    """
    import math
    import soundfile as sf
    from scipy.signal import resample_poly

    buf = io.BytesIO(raw_audio_bytes)
    with sf.SoundFile(buf) as f:
        native_sr = f.samplerate
        n_frames = int(native_sr * max_seconds) if max_seconds else -1
        data = f.read(frames=n_frames, dtype="float64")
    if data.ndim > 1:
        data = np.mean(data, axis=1)
    if native_sr != target_sr:
        g = math.gcd(int(native_sr), int(target_sr))
        data = resample_poly(data, target_sr // g, int(native_sr) // g)
    return data, target_sr


# ---------------------------------------------------------------------------
# Results construction / IO
# ---------------------------------------------------------------------------
def _mean(x):
    return float(np.mean(x)) if x is not None and len(x) else None


def build_results(all_features, cdma_stats, dataset_meta, ici_features=None):
    """Assemble the results dict from computed features.

    ``all_features[species]`` is expected to contain:
      - ``spectrogram``: {spectral_centroids, spectral_bandwidths, dominant_freqs}
      - ``entropy``: list of Shannon-entropy values (bits)
      - ``bispectrum``: list of mean off-diagonal bicoherence values
      - ``temporal``: list of spectral-flatness values
      - ``ici`` (odontocetes only): list of ICI arrays (seconds)
    """
    species = {}
    for sp, f in all_features.items():
        spec = f.get("spectrogram") or {}
        ent = f.get("entropy") or []
        species[sp] = {
            "centroid_hz": _mean(spec.get("spectral_centroids")),
            "bandwidth_hz": _mean(spec.get("spectral_bandwidths")),
            "dominant_hz": _mean(spec.get("dominant_freqs")),
            "entropy_bits": _mean(ent),
            "bicoherence_mean": _mean(f.get("bispectrum")),
            "flatness_mean": _mean(f.get("temporal")),
            "n_samples": len((spec.get("spectral_centroids") or [])),
        }

    ici = {}
    ici_features = ici_features or {}
    for sp, icis in ici_features.items():
        medians = [float(np.median(a) * 1000) for a in icis if a is not None and len(a) > 0]
        counts = [int(len(a) + 1) for a in icis if a is not None and len(a) > 0]
        if medians:
            ici[sp] = {
                "median_ms_min": min(medians),
                "median_ms_max": max(medians),
                "n_clicks_min": min(counts),
                "n_clicks_max": max(counts),
            }

    cdma = {k: (float(v) if v is not None else None) for k, v in (cdma_stats or {}).items()}

    _species_entropy_means = [d["entropy_bits"] for d in species.values()]

    return {
        "dataset": dataset_meta,
        "species_order": list(all_features.keys()),
        "species": species,
        "entropy_range": {
            "min": min(m for m in _species_entropy_means if m is not None) if _species_entropy_means else None,
            "max": max(m for m in _species_entropy_means if m is not None) if _species_entropy_means else None,
        },
        "ici": ici,
        "cdma": cdma,
    }


def write_results(results, path=RESULTS_PATH):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(results, fh, ensure_ascii=False, indent=2)
    return path


def load_results(path=RESULTS_PATH):
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"{path} not found. Run 'python generate_papers.py' first to produce "
            "results.json (data -> analysis -> results.json -> manuscript)."
        )
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


# ---------------------------------------------------------------------------
# Extremes helpers (so prose claims like "the lowest centroid" trace to data)
# ---------------------------------------------------------------------------
def species_extreme(results, metric, which="max"):
    """Return (species_key, value) for the min/max of ``metric`` across species."""
    items = [(sp, d.get(metric)) for sp, d in results["species"].items() if d.get(metric) is not None]
    if not items:
        return None, None
    key = (max if which == "max" else min)(items, key=lambda kv: kv[1])
    return key[0], key[1]


# ---------------------------------------------------------------------------
# Number formatting (consistent rendering everywhere)
# ---------------------------------------------------------------------------
_SUP = str.maketrans("0123456789-", "⁰¹²³⁴⁵⁶⁷⁸⁹⁻")


def fmt_hz(x):
    """Integer Hz with thousands separator, e.g. 2309 -> '2,309'."""
    return f"{round(float(x)):,}"


def fmt_entropy(x):
    return f"{float(x):.2f}"


def fmt_flatness(x):
    return f"{float(x):.4f}"


def fmt_bicoh_plain(x):
    return f"{float(x):.6f}"


def _sci_parts(x):
    x = float(x)
    if x == 0:
        return "0", 0
    exp = int(np.floor(np.log10(abs(x))))
    mant = x / (10 ** exp)
    return f"{mant:.2f}", exp


def fmt_sci_unicode(x):
    """Scientific notation with unicode superscript, e.g. '7.51 × 10⁻⁴'."""
    mant, exp = _sci_parts(x)
    return f"{mant} × 10{str(exp).translate(_SUP)}"


def fmt_sci_ascii(x):
    """Scientific notation for ASCII contexts, e.g. '7.51e-04'."""
    return f"{float(x):.2e}"


def fmt_mean_sd(mean, sd, dec=2):
    return f"{float(mean):.{dec}f} ± {float(sd):.{dec}f}"
