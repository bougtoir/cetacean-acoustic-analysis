.PHONY: all install analysis manuscripts jasa pptx clean

# One-command reproduction:
#   make install   # install pinned dependencies
#   make all       # data -> results.json -> figures -> DOCX + JASA + PPTX
#
# Reproduction chain (single source of truth is output/results.json):
#   public dataset (confit/wmms-parquet, pinned revision)
#     -> bounded/polyphase audio decoding
#     -> feature/statistical analysis
#     -> output/results.json
#     -> manuscript / table / figure / presentation generators
# No empirical result value is hard-coded in the generators; every number
# they emit is read from output/results.json.

all: manuscripts jasa pptx

install:
	pip install -r requirements.txt

# Runs the analysis on the public Watkins dataset, writes output/results.json,
# all figures, and the JA/EN DOCX manuscripts. This is the canonical runner
# that produces results.json.
analysis manuscripts: 
	python generate_papers.py

# JASA-formatted manuscript + cover letter (reads output/results.json).
jasa:
	python generate_jasa.py

# Editable English + Japanese figure/table presentations (reads output/results.json).
pptx:
	python generate_pptx.py

clean:
	rm -f output/results.json output/*.png papers/*.docx papers/*.pptx
