# Cetacean Acoustic Communication Analysis

Quantitative analysis of encoding structures in cetacean acoustic communication,
testing a CDMA-like code-division hypothesis and a beat-frequency hypothesis on
the public **Watkins Marine Mammal Sound Database**.

All empirical numbers reported in the manuscripts, tables, figures, and slides are
computed from public data by the analysis code and stored in
`output/results.json` — **no empirical result is hard-coded** in the generators.

## Data

- Source: Watkins Marine Mammal Sound Database (Woods Hole Oceanographic Institution)
- HuggingFace dataset: `confit/wmms-parquet`
- Pinned revision: `a90a38e0006991f0c6f6d4e05261949a1da7f14e`
- 1,357 recordings across 32 species; six target species are analysed.

Audio is decoded directly from the parquet bytes with `soundfile` (avoiding the
optional `torchcodec` dependency) and resampled to 16 kHz with a polyphase filter
(`scipy.signal.resample_poly`).

### Bounded analysis (why 60 s)

Some Watkins recordings are ~20 minutes long (~50M samples at 44.1 kHz). Decoding
them in full together with FFT-based resampling and full-length spectrograms
exceeds ~8 GB of RAM and is non-deterministic across runs. To keep the pipeline
memory-safe and deterministic, only the **first 60 seconds** of each recording are
read and analysed. This bound is documented as a limitation in the manuscript.

Configuration lives in `results_io.py` (`DATASET_REVISION`, `TARGET_SPECIES`,
`MAX_SAMPLES=10`, `TARGET_SR=16000`, `ANALYSIS_SECONDS=60`).

## Reproduce (one command)

```bash
pip install -r requirements.txt   # or: make install
make all
```

`make all` runs, in order:

1. `generate_papers.py` — loads the public dataset, runs the analysis, writes
   `output/results.json` and all figures, and produces the JA/EN DOCX manuscripts.
2. `generate_jasa.py` — JASA-formatted manuscript + cover letter.
3. `generate_pptx.py` — editable English and Japanese figure/table presentations.

Individual targets: `make manuscripts`, `make jasa`, `make pptx`.

## Outputs

- `output/results.json` — single source of truth for every reported number.
- `output/*.png` — figures (JA and EN variants).
- `papers/鯨類音響コミュニケーション解析_日本語版.docx` — Japanese manuscript.
- `papers/Cetacean_Acoustic_Communication_Analysis_English.docx` — English manuscript.
- `papers/JASA_Manuscript_Cetacean_Encoding.docx`, `papers/JASA_Cover_Letter.docx`.
- `papers/*_図表集.pptx` / `papers/*_Figures.pptx` — editable slide decks.

## Reproduction chain

```
public dataset (pinned revision)
  -> bounded/polyphase audio decoding (results_io.decode_audio)
  -> feature/statistical analysis (generate_papers.py)
  -> output/results.json
  -> manuscript / table / figure / presentation generators
```

Every generator reads `output/results.json` via `results_io.load_results()`; if it
is missing, run `python generate_papers.py` (or `make manuscripts`) first.
