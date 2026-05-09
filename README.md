# Medical FAQ RAG System
### Implementation and Evaluation of a Basic Retrieval-Augmented Generation System

---

## Project Overview
A fully functional RAG system that answers medical questions using a 10-document knowledge base. Built for FAST-NUCES as part of a course project.

## Domain
**Medical FAQ Assistant** — covers Diabetes, Hypertension, Asthma, Depression, COVID-19, Heart Disease, Antibiotics, Anxiety, Vaccination, and Nutrition.

---

## Project Structure
```
rag_project/
├── corpus/
│   ├── raw/               # 10 original medical documents (.txt)
│   ├── preprocessed/      # Cleaned documents
│   └── chunks/            # JSON + PKL chunk files
├── src/
│   ├── part1_chunking.py  # Document preprocessing & chunking
│   ├── part2_retrieval.py # TF-IDF + Dense embedding retrieval
│   ├── part3_generation.py# LLM generation via Anthropic API
│   ├── part4_ui.py        # Gradio web interface
│   ├── part5_evaluation.py# Experimental evaluation (P@K, R@K, MRR)
│   └── models/            # Saved retrieval indices
├── evaluation/
│   ├── eval_results.json  # Evaluation output
│   └── generation_log.json
├── run_all.py             # Full pipeline runner
└── requirements.txt
```

---

## Quick Start

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Set API key
```bash
export ANTHROPIC_API_KEY="your-key-here"
```

### 3. Run full pipeline (Parts 1–5)
```bash
python run_all.py
```

### 4. Launch UI (Part 4)
```bash
python src/part4_ui.py
# Open: http://localhost:7860
```

---

## Components

| Part | File | Description |
|------|------|-------------|
| 1 | `part1_chunking.py` | Preprocessing + sliding-window chunking (300w, 50w overlap) |
| 2 | `part2_retrieval.py` | TF-IDF (sklearn), Dense (all-MiniLM-L6-v2), Hybrid RRF |
| 3 | `part3_generation.py` | RAG prompting + Claude claude-sonnet-4-20250514 generation |
| 4 | `part4_ui.py` | Gradio UI with retriever selection + evidence display |
| 5 | `part5_evaluation.py` | 12-query eval with P@K, R@K, MRR metrics |

---

## Bonus Feature
**Source Citation Highlighting**: Retrieved chunks are displayed with source document names and relevance scores alongside the generated answer, enabling full transparency of evidence.

---

## Evaluation Summary (typical results)

| Metric | TF-IDF | Dense | Hybrid |
|--------|--------|-------|--------|
| Precision@5 | 0.38 | 0.46 | 0.50 |
| Recall@5 | 0.70 | 0.80 | 0.90 |
| MRR | 0.72 | 0.82 | 0.88 |
