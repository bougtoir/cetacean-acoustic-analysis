"""
鯨類音響コミュニケーション解析スクリプト
Cetacean Acoustic Communication Analysis

仮説検証:
1. エンコード・デコード構造の検出（情報エントロピー、Zipf則）
2. 個体/種識別パターンの抽出（スペクトル特徴、ICI/IPI分析）
3. 非線形音響現象の検出（バイスペクトル、高次統計量）
4. 時間構造・相互相関分析

データ: Watkins Marine Mammal Sound Database (HuggingFace: confit/wmms-parquet)
"""

import os
import warnings

import matplotlib
import matplotlib.pyplot as plt
import numpy as np
from scipy import signal as sp_signal
from scipy import stats

matplotlib.use("Agg")
warnings.filterwarnings("ignore")

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Japanese + English labels for figures
plt.rcParams["font.family"] = "DejaVu Sans"


def load_whale_data():
    """Load Watkins Marine Mammal Sound Database from HuggingFace.

    Returns the raw arrow table and label metadata to avoid torchcodec dependency.
    Audio is decoded manually via soundfile in get_species_samples().
    """
    from datasets import load_dataset

    print("Loading Watkins Marine Mammal Sound Database...")
    ds = load_dataset("confit/wmms-parquet", split="train")
    label_names = ds.features["label"].names
    # Extract arrow table and label column to bypass torchcodec
    table = ds.data
    labels = [v.as_py() for v in table.column("label")]
    n_species = len(set(labels))
    print(f"  Loaded {len(labels)} samples, {n_species} species")
    return table, labels, label_names


def _decode_audio(raw_audio_bytes, target_sr=16000):
    """Decode raw audio bytes using soundfile (avoids torch dependency)."""
    import io
    import soundfile as sf

    buf = io.BytesIO(raw_audio_bytes)
    data, sr = sf.read(buf, dtype="float64")
    # Convert stereo to mono if needed
    if data.ndim > 1:
        data = np.mean(data, axis=1)
    # Resample if needed
    if sr != target_sr:
        from scipy.signal import resample

        n_samples = int(len(data) * target_sr / sr)
        data = resample(data, n_samples)
        sr = target_sr
    return data, sr


def get_species_samples(table, labels, label_names, species_name, max_samples=10):
    """Get audio samples for a given species using raw arrow table access."""
    if species_name not in label_names:
        print(f"  Species '{species_name}' not found. Available: {label_names}")
        return []
    target_label = label_names.index(species_name)

    audio_col = table.column("audio")
    samples = []
    for i, lab in enumerate(labels):
        if lab == target_label and len(samples) < max_samples:
            raw_struct = audio_col[i].as_py()
            audio_bytes = raw_struct["bytes"]
            try:
                array, sr = _decode_audio(audio_bytes)
                samples.append(
                    {
                        "array": array,
                        "sr": sr,
                        "species": species_name,
                        "index": i,
                    }
                )
            except Exception as e:
                print(f"  Warning: could not decode sample {i}: {e}")
    print(f"  Found {len(samples)} samples for {species_name}")
    return samples


# =============================================================================
# Analysis 1: Spectrogram and Spectral Feature Analysis
# =============================================================================


def analyze_spectrograms(samples, species_name):
    """
    Generate spectrograms and extract spectral features for species identification.
    Tests hypothesis: individual/species-specific frequency patterns exist.
    """
    print(f"\n[Analysis 1] Spectrogram Analysis: {species_name}")

    if not samples:
        print("  No samples available.")
        return None

    n_samples = min(len(samples), 6)
    fig, axes = plt.subplots(2, 3, figsize=(18, 10))
    fig.suptitle(
        f"Spectrogram Analysis: {species_name}\n"
        f"(Hypothesis: Species-specific frequency patterns / "
        f"種固有の周波数パターンの検出)",
        fontsize=14,
    )

    spectral_centroids = []
    spectral_bandwidths = []
    dominant_freqs = []

    for idx in range(n_samples):
        sample = samples[idx]
        audio = sample["array"]
        sr = sample["sr"]

        # Compute spectrogram
        nperseg = min(1024, len(audio) // 4)
        if nperseg < 64:
            continue
        freqs, times, Sxx = sp_signal.spectrogram(
            audio, fs=sr, nperseg=nperseg, noverlap=nperseg // 2
        )

        # Plot spectrogram
        row, col = idx // 3, idx % 3
        ax = axes[row, col]
        im = ax.pcolormesh(
            times,
            freqs,
            10 * np.log10(Sxx + 1e-10),
            shading="gouraud",
            cmap="viridis",
        )
        ax.set_ylabel("Frequency (Hz)")
        ax.set_xlabel("Time (s)")
        ax.set_title(f"Sample {idx + 1}")
        fig.colorbar(im, ax=ax, label="Power (dB)")

        # Extract spectral features
        power_spectrum = np.mean(Sxx, axis=1)
        total_power = np.sum(power_spectrum)
        if total_power > 0:
            centroid = np.sum(freqs * power_spectrum) / total_power
            bandwidth = np.sqrt(
                np.sum(((freqs - centroid) ** 2) * power_spectrum) / total_power
            )
            dominant_freq = freqs[np.argmax(power_spectrum)]
            spectral_centroids.append(centroid)
            spectral_bandwidths.append(bandwidth)
            dominant_freqs.append(dominant_freq)

    plt.tight_layout()
    filepath = os.path.join(OUTPUT_DIR, f"spectrogram_{species_name}.png")
    plt.savefig(filepath, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  Saved: {filepath}")

    features = {
        "spectral_centroids": spectral_centroids,
        "spectral_bandwidths": spectral_bandwidths,
        "dominant_freqs": dominant_freqs,
    }

    if spectral_centroids:
        print(f"  Spectral centroid: {np.mean(spectral_centroids):.1f} +/- "
              f"{np.std(spectral_centroids):.1f} Hz")
        print(f"  Spectral bandwidth: {np.mean(spectral_bandwidths):.1f} +/- "
              f"{np.std(spectral_bandwidths):.1f} Hz")
        print(f"  Dominant frequency: {np.mean(dominant_freqs):.1f} +/- "
              f"{np.std(dominant_freqs):.1f} Hz")

    return features


# =============================================================================
# Analysis 2: Inter-Click Interval (ICI) Analysis for Sperm Whales
# =============================================================================


def detect_clicks(audio, sr, threshold_factor=3.0):
    """
    Detect click events in sperm whale audio using envelope detection.
    Returns click times in seconds.
    """
    # Compute analytic signal envelope
    analytic = sp_signal.hilbert(audio)
    envelope = np.abs(analytic)

    # Smooth the envelope
    window_size = max(int(sr * 0.001), 3)  # 1ms window
    if window_size % 2 == 0:
        window_size += 1
    smoothed = sp_signal.medfilt(envelope, kernel_size=window_size)

    # Threshold-based detection
    threshold = np.mean(smoothed) + threshold_factor * np.std(smoothed)
    above_threshold = smoothed > threshold

    # Find onset times
    diff = np.diff(above_threshold.astype(int))
    onsets = np.where(diff == 1)[0]

    click_times = onsets / sr
    return click_times


def analyze_ici(samples, species_name):
    """
    Analyze Inter-Click Intervals (ICI) for click-producing species.
    Tests hypothesis: ICI patterns encode individual/group identity (CDMA-like codes).
    """
    print(f"\n[Analysis 2] Inter-Click Interval Analysis: {species_name}")

    if not samples:
        print("  No samples available.")
        return None

    all_icis = []
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    fig.suptitle(
        f"Inter-Click Interval (ICI) Analysis: {species_name}\n"
        f"(Hypothesis: ICI patterns as identity codes / "
        f"ICI パターンによる個体識別符号の検出)",
        fontsize=14,
    )

    for idx, sample in enumerate(samples[:4]):
        audio = sample["array"]
        sr = sample["sr"]

        click_times = detect_clicks(audio, sr)

        if len(click_times) > 1:
            icis = np.diff(click_times)
            icis = icis[(icis > 0.001) & (icis < 2.0)]  # Filter reasonable range
            all_icis.append(icis)

            ax = axes[idx // 2, idx % 2]
            if len(icis) > 0:
                ax.hist(icis * 1000, bins=50, alpha=0.7, color="steelblue",
                        edgecolor="black")
                ax.axvline(
                    np.median(icis) * 1000,
                    color="red",
                    linestyle="--",
                    label=f"Median: {np.median(icis) * 1000:.1f} ms",
                )
                ax.set_xlabel("ICI (ms)")
                ax.set_ylabel("Count")
                ax.set_title(f"Sample {idx + 1} ({len(click_times)} clicks)")
                ax.legend()
            else:
                ax.text(0.5, 0.5, "No valid ICIs detected",
                        transform=ax.transAxes, ha="center")
                ax.set_title(f"Sample {idx + 1}")
        else:
            ax = axes[idx // 2, idx % 2]
            ax.text(
                0.5,
                0.5,
                f"Only {len(click_times)} click(s) detected",
                transform=ax.transAxes,
                ha="center",
            )
            ax.set_title(f"Sample {idx + 1}")

    plt.tight_layout()
    filepath = os.path.join(OUTPUT_DIR, f"ici_analysis_{species_name}.png")
    plt.savefig(filepath, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  Saved: {filepath}")

    # Summary statistics
    if all_icis:
        for i, icis in enumerate(all_icis):
            if len(icis) > 0:
                print(
                    f"  Sample {i + 1}: median ICI = {np.median(icis) * 1000:.1f} ms, "
                    f"std = {np.std(icis) * 1000:.1f} ms, n = {len(icis)}"
                )

    return all_icis


# =============================================================================
# Analysis 3: Bispectrum / Higher-Order Spectral Analysis
# =============================================================================


def compute_bispectrum(audio, sr, nfft=512):
    """
    Compute the bispectrum to detect nonlinear interactions (quadratic phase coupling).
    The bispectrum B(f1, f2) = E[X(f1) * X(f2) * conj(X(f1+f2))]
    is non-zero only when there is phase coupling between f1, f2, and f1+f2.
    This tests the beat frequency / parametric speaker hypothesis.
    """
    n_segments = max(1, len(audio) // nfft - 1)
    n_freq = nfft // 2

    bispectrum = np.zeros((n_freq, n_freq), dtype=complex)

    for seg in range(n_segments):
        start = seg * nfft
        segment = audio[start : start + nfft]
        if len(segment) < nfft:
            break

        # Apply window and compute FFT
        windowed = segment * np.hanning(nfft)
        X = np.fft.fft(windowed, nfft)

        # Compute bispectrum for this segment
        for i in range(n_freq):
            j_max = min(n_freq, nfft - i)
            for j in range(j_max):
                if i + j < nfft:
                    bispectrum[i, j] += X[i] * X[j] * np.conj(X[i + j])

    if n_segments > 0:
        bispectrum /= n_segments

    # Compute bicoherence (normalized bispectrum)
    bicoherence = np.abs(bispectrum) ** 2
    max_val = np.max(bicoherence)
    if max_val > 0:
        bicoherence = bicoherence / max_val

    freqs = np.fft.fftfreq(nfft, d=1.0 / sr)[:n_freq]
    return bicoherence, freqs


def analyze_bispectrum(samples, species_name):
    """
    Perform bispectral analysis to detect nonlinear frequency coupling.
    Tests hypothesis: beat frequencies / parametric effects carry meaning.
    """
    print(f"\n[Analysis 3] Bispectrum (Nonlinear Coupling) Analysis: {species_name}")

    if not samples:
        print("  No samples available.")
        return None

    n_samples = min(len(samples), 4)
    fig, axes = plt.subplots(2, 2, figsize=(14, 12))
    fig.suptitle(
        f"Bispectrum Analysis: {species_name}\n"
        f"(Hypothesis: Nonlinear frequency coupling / "
        f"非線形周波数結合 = うなり効果の検出)",
        fontsize=14,
    )

    coupling_strengths = []

    for idx in range(n_samples):
        sample = samples[idx]
        audio = sample["array"]
        sr = sample["sr"]

        # Truncate to manageable length for bispectrum
        max_len = sr * 5  # 5 seconds max
        audio_trunc = audio[: int(max_len)]

        nfft = min(512, len(audio_trunc) // 4)
        if nfft < 64:
            continue

        bicoherence, freqs = compute_bispectrum(audio_trunc, sr, nfft=nfft)

        ax = axes[idx // 2, idx % 2]
        max_freq_idx = min(len(freqs), nfft // 4)
        im = ax.pcolormesh(
            freqs[:max_freq_idx],
            freqs[:max_freq_idx],
            bicoherence[:max_freq_idx, :max_freq_idx],
            shading="gouraud",
            cmap="hot",
        )
        ax.set_xlabel("f1 (Hz)")
        ax.set_ylabel("f2 (Hz)")
        ax.set_title(f"Sample {idx + 1}")
        fig.colorbar(im, ax=ax, label="Bicoherence")

        # Quantify coupling strength (off-diagonal energy)
        n = min(max_freq_idx, bicoherence.shape[0])
        off_diag = bicoherence[:n, :n].copy()
        np.fill_diagonal(off_diag, 0)
        coupling_strength = np.mean(off_diag)
        coupling_strengths.append(coupling_strength)
        print(f"  Sample {idx + 1}: mean off-diagonal bicoherence = "
              f"{coupling_strength:.6f}")

    plt.tight_layout()
    filepath = os.path.join(OUTPUT_DIR, f"bispectrum_{species_name}.png")
    plt.savefig(filepath, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  Saved: {filepath}")

    return coupling_strengths


# =============================================================================
# Analysis 4: Information Entropy and Zipf-like Distribution
# =============================================================================


def analyze_information_content(samples, species_name):
    """
    Analyze information-theoretic properties of vocalizations.
    Tests hypothesis: vocalizations follow encoding-like statistical structure.
    """
    print(f"\n[Analysis 4] Information Entropy Analysis: {species_name}")

    if not samples:
        print("  No samples available.")
        return None

    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    fig.suptitle(
        f"Information-Theoretic Analysis: {species_name}\n"
        f"(Hypothesis: Encoding structure / エンコーディング構造の検出)",
        fontsize=14,
    )

    all_entropies = []
    all_power_distributions = []

    for idx, sample in enumerate(samples[:4]):
        audio = sample["array"]
        sr = sample["sr"]

        # Compute power spectrum
        nperseg = min(1024, len(audio) // 4)
        if nperseg < 64:
            continue
        freqs, psd = sp_signal.welch(audio, fs=sr, nperseg=nperseg)

        # Normalize PSD to probability distribution
        psd_norm = psd / (np.sum(psd) + 1e-10)
        psd_norm = psd_norm[psd_norm > 0]

        # Shannon entropy
        entropy = -np.sum(psd_norm * np.log2(psd_norm + 1e-10))
        all_entropies.append(entropy)

        # Rank-frequency distribution (Zipf-like analysis)
        sorted_psd = np.sort(psd_norm)[::-1]
        ranks = np.arange(1, len(sorted_psd) + 1)

        all_power_distributions.append(sorted_psd)

        # Plot rank-frequency on log-log scale
        ax = axes[idx // 2, idx % 2]
        ax.loglog(ranks, sorted_psd, "b-", alpha=0.7, label="Observed")

        # Fit power law (Zipf)
        log_ranks = np.log10(ranks[sorted_psd > 0])
        log_psd = np.log10(sorted_psd[sorted_psd > 0])
        if len(log_ranks) > 2:
            slope, intercept, r_value, p_value, _ = stats.linregress(
                log_ranks, log_psd
            )
            fitted = 10 ** (intercept + slope * log_ranks)
            ax.loglog(
                ranks[sorted_psd > 0],
                fitted,
                "r--",
                alpha=0.7,
                label=f"Zipf fit: alpha={-slope:.2f}, R^2={r_value**2:.3f}",
            )

        ax.set_xlabel("Rank")
        ax.set_ylabel("Normalized Power")
        ax.set_title(f"Sample {idx + 1} (H = {entropy:.2f} bits)")
        ax.legend(fontsize=8)

    plt.tight_layout()
    filepath = os.path.join(OUTPUT_DIR, f"entropy_{species_name}.png")
    plt.savefig(filepath, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  Saved: {filepath}")

    if all_entropies:
        print(f"  Mean Shannon entropy: {np.mean(all_entropies):.2f} +/- "
              f"{np.std(all_entropies):.2f} bits")

    return all_entropies


# =============================================================================
# Analysis 5: Cross-Correlation and Temporal Structure
# =============================================================================


def analyze_temporal_structure(samples, species_name):
    """
    Analyze temporal autocorrelation and cross-correlation structure.
    Tests hypothesis: temporal patterns carry structured information.
    """
    print(f"\n[Analysis 5] Temporal Structure Analysis: {species_name}")

    if not samples or len(samples) < 2:
        print("  Need at least 2 samples.")
        return None

    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    fig.suptitle(
        f"Temporal Structure Analysis: {species_name}\n"
        f"(Hypothesis: Structured temporal patterns / "
        f"構造化された時間パターンの検出)",
        fontsize=14,
    )

    # Autocorrelation analysis
    ax = axes[0, 0]
    for idx, sample in enumerate(samples[:4]):
        audio = sample["array"]
        sr = sample["sr"]

        # Compute autocorrelation
        max_lag = min(len(audio) // 2, int(sr * 0.5))  # up to 500ms
        autocorr = np.correlate(audio[:max_lag * 2], audio[:max_lag * 2], mode="full")
        autocorr = autocorr[len(autocorr) // 2 :]
        autocorr = autocorr / (autocorr[0] + 1e-10)

        lags_ms = np.arange(len(autocorr)) / sr * 1000
        ax.plot(lags_ms[:max_lag], autocorr[:max_lag], alpha=0.7,
                label=f"Sample {idx + 1}")

    ax.set_xlabel("Lag (ms)")
    ax.set_ylabel("Autocorrelation")
    ax.set_title("Autocorrelation Functions")
    ax.legend(fontsize=8)
    ax.set_xlim(0, 100)

    # Cross-correlation between samples
    ax = axes[0, 1]
    if len(samples) >= 2:
        for i in range(min(3, len(samples) - 1)):
            audio1 = samples[i]["array"]
            audio2 = samples[i + 1]["array"]
            sr = samples[i]["sr"]

            # Align lengths
            min_len = min(len(audio1), len(audio2))
            min_len = min(min_len, sr * 5)  # Max 5 seconds
            a1 = audio1[: int(min_len)]
            a2 = audio2[: int(min_len)]

            # Cross-correlation
            xcorr = np.correlate(a1, a2, mode="full")
            max_val = np.max(np.abs(xcorr))
            if max_val > 0:
                xcorr = xcorr / max_val

            center = len(xcorr) // 2
            lag_range = min(int(sr * 0.1), center)  # +/- 100ms
            lags = np.arange(-lag_range, lag_range) / sr * 1000
            xcorr_slice = xcorr[center - lag_range : center + lag_range]

            ax.plot(
                lags[: len(xcorr_slice)],
                xcorr_slice,
                alpha=0.7,
                label=f"S{i + 1} x S{i + 2}",
            )

    ax.set_xlabel("Lag (ms)")
    ax.set_ylabel("Cross-correlation")
    ax.set_title("Cross-correlation Between Samples")
    ax.legend(fontsize=8)

    # Spectral flatness (measure of encoding complexity)
    ax = axes[1, 0]
    flatness_values = []
    for idx, sample in enumerate(samples[:6]):
        audio = sample["array"]
        sr = sample["sr"]

        nperseg = min(1024, len(audio) // 4)
        if nperseg < 64:
            continue
        freqs, psd = sp_signal.welch(audio, fs=sr, nperseg=nperseg)
        psd_pos = psd[psd > 0]
        if len(psd_pos) > 0:
            geometric_mean = np.exp(np.mean(np.log(psd_pos)))
            arithmetic_mean = np.mean(psd_pos)
            flatness = geometric_mean / (arithmetic_mean + 1e-10)
            flatness_values.append(flatness)

    if flatness_values:
        ax.bar(range(len(flatness_values)), flatness_values, color="teal", alpha=0.7)
        ax.set_xlabel("Sample Index")
        ax.set_ylabel("Spectral Flatness")
        ax.set_title(
            "Spectral Flatness (1.0 = noise-like, 0.0 = tonal)\n"
            "Lower = more structured encoding"
        )
        ax.axhline(y=np.mean(flatness_values), color="red", linestyle="--",
                    label=f"Mean: {np.mean(flatness_values):.4f}")
        ax.legend()

    # Temporal modulation spectrum
    ax = axes[1, 1]
    for idx, sample in enumerate(samples[:4]):
        audio = sample["array"]
        sr = sample["sr"]

        # Compute envelope
        analytic = sp_signal.hilbert(audio[: min(len(audio), sr * 5)])
        envelope = np.abs(analytic)

        # Modulation spectrum (FFT of envelope)
        mod_spectrum = np.abs(np.fft.rfft(envelope))
        mod_freqs = np.fft.rfftfreq(len(envelope), d=1.0 / sr)

        # Show modulation rates up to 50 Hz
        mask = mod_freqs <= 50
        if np.max(mod_spectrum[mask]) > 0:
            ax.plot(
                mod_freqs[mask],
                mod_spectrum[mask] / np.max(mod_spectrum[mask]),
                alpha=0.7,
                label=f"Sample {idx + 1}",
            )

    ax.set_xlabel("Modulation Rate (Hz)")
    ax.set_ylabel("Normalized Amplitude")
    ax.set_title("Temporal Modulation Spectrum\n(Envelope periodicity)")
    ax.legend(fontsize=8)

    plt.tight_layout()
    filepath = os.path.join(OUTPUT_DIR, f"temporal_{species_name}.png")
    plt.savefig(filepath, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  Saved: {filepath}")

    return flatness_values


# =============================================================================
# Analysis 6: Cross-Species Comparison
# =============================================================================


def compare_species(all_features):
    """
    Compare spectral and temporal features across species.
    Tests hypothesis: each species has distinct encoding characteristics.
    """
    print("\n[Analysis 6] Cross-Species Feature Comparison")

    species_list = list(all_features.keys())
    if len(species_list) < 2:
        print("  Need at least 2 species for comparison.")
        return

    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    fig.suptitle(
        "Cross-Species Acoustic Feature Comparison\n"
        "(Hypothesis: Species-specific encoding signatures / "
        "種固有のエンコーディング特性)",
        fontsize=14,
    )

    # 1. Spectral centroid comparison
    ax = axes[0, 0]
    centroids = {}
    for sp in species_list:
        if "spectrogram" in all_features[sp] and all_features[sp]["spectrogram"]:
            c = all_features[sp]["spectrogram"]["spectral_centroids"]
            if c:
                centroids[sp] = c
    if centroids:
        positions = range(len(centroids))
        labels = [s.replace("_", "\n") for s in centroids.keys()]
        data = list(centroids.values())
        bp = ax.boxplot(data, positions=list(positions), labels=labels,
                        patch_artist=True)
        colors = plt.cm.Set3(np.linspace(0, 1, len(data)))
        for patch, color in zip(bp["boxes"], colors):
            patch.set_facecolor(color)
        ax.set_ylabel("Spectral Centroid (Hz)")
        ax.set_title("Spectral Centroid by Species")
        ax.tick_params(axis="x", rotation=45, labelsize=8)

    # 2. Entropy comparison
    ax = axes[0, 1]
    entropies = {}
    for sp in species_list:
        if "entropy" in all_features[sp] and all_features[sp]["entropy"]:
            entropies[sp] = all_features[sp]["entropy"]
    if entropies:
        positions = range(len(entropies))
        labels = [s.replace("_", "\n") for s in entropies.keys()]
        data = list(entropies.values())
        bp = ax.boxplot(data, positions=list(positions), labels=labels,
                        patch_artist=True)
        colors = plt.cm.Set3(np.linspace(0, 1, len(data)))
        for patch, color in zip(bp["boxes"], colors):
            patch.set_facecolor(color)
        ax.set_ylabel("Shannon Entropy (bits)")
        ax.set_title("Information Entropy by Species")
        ax.tick_params(axis="x", rotation=45, labelsize=8)

    # 3. Bispectral coupling strength
    ax = axes[1, 0]
    couplings = {}
    for sp in species_list:
        if "bispectrum" in all_features[sp] and all_features[sp]["bispectrum"]:
            couplings[sp] = all_features[sp]["bispectrum"]
    if couplings:
        positions = range(len(couplings))
        labels = [s.replace("_", "\n") for s in couplings.keys()]
        data = list(couplings.values())
        bp = ax.boxplot(data, positions=list(positions), labels=labels,
                        patch_artist=True)
        colors = plt.cm.Set3(np.linspace(0, 1, len(data)))
        for patch, color in zip(bp["boxes"], colors):
            patch.set_facecolor(color)
        ax.set_ylabel("Bicoherence (coupling strength)")
        ax.set_title("Nonlinear Coupling Strength by Species\n(Higher = more beat/parametric effects)")
        ax.tick_params(axis="x", rotation=45, labelsize=8)

    # 4. Spectral flatness comparison
    ax = axes[1, 1]
    flatness = {}
    for sp in species_list:
        if "temporal" in all_features[sp] and all_features[sp]["temporal"]:
            flatness[sp] = all_features[sp]["temporal"]
    if flatness:
        positions = range(len(flatness))
        labels = [s.replace("_", "\n") for s in flatness.keys()]
        data = list(flatness.values())
        bp = ax.boxplot(data, positions=list(positions), labels=labels,
                        patch_artist=True)
        colors = plt.cm.Set3(np.linspace(0, 1, len(data)))
        for patch, color in zip(bp["boxes"], colors):
            patch.set_facecolor(color)
        ax.set_ylabel("Spectral Flatness")
        ax.set_title("Spectral Flatness by Species\n(Lower = more structured/encoded)")
        ax.tick_params(axis="x", rotation=45, labelsize=8)

    plt.tight_layout()
    filepath = os.path.join(OUTPUT_DIR, "cross_species_comparison.png")
    plt.savefig(filepath, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  Saved: {filepath}")


# =============================================================================
# Analysis 7: CDMA-like Orthogonality Test
# =============================================================================


def analyze_cdma_orthogonality(all_species_samples):
    """
    Test whether different individuals/species have orthogonal spectral signatures.
    This is analogous to CDMA where each user has a code with low cross-correlation.
    """
    print("\n[Analysis 7] CDMA-like Orthogonality Analysis")

    # Compute spectral fingerprints for each sample
    fingerprints = []
    labels = []
    species_labels = []

    for species_name, samples in all_species_samples.items():
        for sample in samples[:5]:
            audio = sample["array"]
            sr = sample["sr"]

            nperseg = min(512, len(audio) // 4)
            if nperseg < 64:
                continue
            freqs, psd = sp_signal.welch(audio, fs=sr, nperseg=nperseg)

            # Normalize to unit vector (like a CDMA spreading code)
            norm = np.linalg.norm(psd)
            if norm > 0:
                fingerprint = psd / norm
                fingerprints.append(fingerprint)
                labels.append(f"{species_name}_{sample['index']}")
                species_labels.append(species_name)

    if len(fingerprints) < 2:
        print("  Not enough fingerprints for comparison.")
        return

    # Ensure all fingerprints have the same length
    min_len = min(len(fp) for fp in fingerprints)
    fingerprints = [fp[:min_len] for fp in fingerprints]
    fingerprints = np.array(fingerprints)

    # Compute cross-correlation matrix (like CDMA code correlation)
    n = len(fingerprints)
    correlation_matrix = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            correlation_matrix[i, j] = np.dot(fingerprints[i], fingerprints[j])

    # Plot
    fig, axes = plt.subplots(1, 2, figsize=(18, 8))
    fig.suptitle(
        "CDMA-like Orthogonality Analysis\n"
        "(Hypothesis: Species have orthogonal spectral codes / "
        "種ごとの直交スペクトル符号の検証)",
        fontsize=14,
    )

    ax = axes[0]
    im = ax.imshow(correlation_matrix, cmap="RdBu_r", vmin=-0.1, vmax=1.0)
    ax.set_title("Spectral Cross-Correlation Matrix")
    fig.colorbar(im, ax=ax)

    # Add species boundary lines
    unique_species = list(dict.fromkeys(species_labels))
    boundaries = []
    count = 0
    for sp in unique_species:
        sp_count = species_labels.count(sp)
        count += sp_count
        boundaries.append(count - 0.5)
    for b in boundaries[:-1]:
        ax.axhline(y=b, color="white", linewidth=2)
        ax.axvline(x=b, color="white", linewidth=2)

    # Compute within-species vs between-species correlation
    within = []
    between = []
    for i in range(n):
        for j in range(i + 1, n):
            if species_labels[i] == species_labels[j]:
                within.append(correlation_matrix[i, j])
            else:
                between.append(correlation_matrix[i, j])

    ax = axes[1]
    if within and between:
        ax.hist(within, bins=30, alpha=0.7, label=f"Within-species (n={len(within)})",
                color="steelblue", density=True)
        ax.hist(between, bins=30, alpha=0.7, label=f"Between-species (n={len(between)})",
                color="coral", density=True)
        ax.axvline(np.mean(within), color="blue", linestyle="--",
                   label=f"Within mean: {np.mean(within):.3f}")
        ax.axvline(np.mean(between), color="red", linestyle="--",
                   label=f"Between mean: {np.mean(between):.3f}")
        ax.set_xlabel("Spectral Correlation")
        ax.set_ylabel("Density")
        ax.set_title(
            "Within vs Between Species Correlation\n"
            "(Low between = good orthogonality = CDMA-like)"
        )
        ax.legend()

        # Statistical test
        if len(within) > 1 and len(between) > 1:
            t_stat, p_value = stats.mannwhitneyu(within, between, alternative="greater")
            print(f"  Within-species correlation: {np.mean(within):.4f} +/- "
                  f"{np.std(within):.4f}")
            print(f"  Between-species correlation: {np.mean(between):.4f} +/- "
                  f"{np.std(between):.4f}")
            print(f"  Mann-Whitney U test: U={t_stat:.1f}, p={p_value:.2e}")
            if p_value < 0.05:
                print("  -> Species show significantly different spectral codes "
                      "(supports CDMA-like hypothesis)")
            else:
                print("  -> No significant difference in spectral codes")

    plt.tight_layout()
    filepath = os.path.join(OUTPUT_DIR, "cdma_orthogonality.png")
    plt.savefig(filepath, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  Saved: {filepath}")


# =============================================================================
# Main Execution
# =============================================================================


def main():
    print("=" * 70)
    print("Cetacean Acoustic Communication Analysis")
    print("鯨類音響コミュニケーション解析")
    print("=" * 70)

    table, labels, label_names = load_whale_data()

    # Target species for analysis
    target_species = [
        "Sperm_Whale",
        "Humpback_Whale",
        "Killer_Whale",
        "Fin,_Finback_Whale",
        "Bottlenose_Dolphin",
        "Beluga,_White_Whale",
    ]

    all_features = {}
    all_species_samples = {}

    for species in target_species:
        print(f"\n{'='*50}")
        print(f"Processing: {species}")
        print(f"{'='*50}")

        samples = get_species_samples(
            table, labels, label_names, species, max_samples=10
        )
        if not samples:
            continue

        all_species_samples[species] = samples

        features = {}

        # Analysis 1: Spectrogram
        spec_features = analyze_spectrograms(samples, species)
        features["spectrogram"] = spec_features

        # Analysis 2: ICI (mainly for click-producing species)
        if species in ["Sperm_Whale", "Bottlenose_Dolphin", "Killer_Whale"]:
            ici_data = analyze_ici(samples, species)
            features["ici"] = ici_data

        # Analysis 3: Bispectrum
        coupling = analyze_bispectrum(samples, species)
        features["bispectrum"] = coupling

        # Analysis 4: Information entropy
        entropy_data = analyze_information_content(samples, species)
        features["entropy"] = entropy_data

        # Analysis 5: Temporal structure
        temporal_data = analyze_temporal_structure(samples, species)
        features["temporal"] = temporal_data

        all_features[species] = features

    # Analysis 6: Cross-species comparison
    if len(all_features) >= 2:
        compare_species(all_features)

    # Analysis 7: CDMA orthogonality
    if len(all_species_samples) >= 2:
        analyze_cdma_orthogonality(all_species_samples)

    # Generate summary report
    generate_summary_report(all_features)

    print(f"\n{'='*70}")
    print(f"Analysis complete. Results saved to: {OUTPUT_DIR}")
    print(f"{'='*70}")


def generate_summary_report(all_features):
    """Generate a text summary of all analyses."""
    report_path = os.path.join(OUTPUT_DIR, "analysis_summary.md")

    with open(report_path, "w", encoding="utf-8") as f:
        f.write("# Cetacean Acoustic Communication Analysis - Summary Report\n")
        f.write("# 鯨類音響コミュニケーション解析 - サマリーレポート\n\n")
        f.write("---\n\n")

        f.write("## Hypothesis Validation Results / 仮説検証結果\n\n")

        f.write("### 1. Encoding Structure (エンコーディング構造)\n\n")
        for sp, features in all_features.items():
            if "entropy" in features and features["entropy"]:
                mean_h = np.mean(features["entropy"])
                f.write(f"- **{sp}**: Shannon entropy = {mean_h:.2f} bits\n")
        f.write("\n")

        f.write("### 2. Nonlinear Coupling / Beat Frequency Effects "
                "(非線形結合 / うなり効果)\n\n")
        for sp, features in all_features.items():
            if "bispectrum" in features and features["bispectrum"]:
                mean_c = np.mean(features["bispectrum"])
                f.write(f"- **{sp}**: mean bicoherence = {mean_c:.6f}\n")
        f.write("\n")

        f.write("### 3. Spectral Features (スペクトル特徴)\n\n")
        for sp, features in all_features.items():
            if "spectrogram" in features and features["spectrogram"]:
                sc = features["spectrogram"]["spectral_centroids"]
                df = features["spectrogram"]["dominant_freqs"]
                if sc and df:
                    f.write(
                        f"- **{sp}**: centroid = {np.mean(sc):.0f} Hz, "
                        f"dominant = {np.mean(df):.0f} Hz\n"
                    )
        f.write("\n")

        f.write("### 4. Temporal Structure (時間構造)\n\n")
        for sp, features in all_features.items():
            if "temporal" in features and features["temporal"]:
                mean_f = np.mean(features["temporal"])
                f.write(f"- **{sp}**: spectral flatness = {mean_f:.4f}\n")
        f.write("\n")

        f.write("---\n\n")
        f.write("## Generated Visualizations / 生成された可視化\n\n")
        f.write("| File | Description |\n")
        f.write("|------|-------------|\n")

        output_files = sorted(os.listdir(OUTPUT_DIR))
        for fname in output_files:
            if fname.endswith(".png"):
                f.write(f"| `{fname}` | See analysis above |\n")

        f.write("\n---\n\n")
        f.write("## Interpretation / 解釈\n\n")
        f.write(
            "The analyses above provide quantitative measures that can be used to "
            "evaluate the hypotheses about cetacean communication encoding:\n\n"
        )
        f.write(
            "1. **Entropy analysis** shows whether vocalizations carry structured "
            "information (lower entropy = more structured = more encoded)\n"
        )
        f.write(
            "2. **Bispectral analysis** reveals nonlinear frequency coupling, "
            "which would support the beat frequency / parametric speaker hypothesis\n"
        )
        f.write(
            "3. **CDMA orthogonality** tests whether different species/individuals "
            "have separable spectral signatures, supporting the encoding hypothesis\n"
        )
        f.write(
            "4. **ICI analysis** for sperm whales shows temporal click patterns "
            "that could serve as identity codes\n"
        )

    print(f"\n  Summary report saved: {report_path}")


if __name__ == "__main__":
    main()
