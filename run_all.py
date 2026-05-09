"""
run_all.py  —  Full pipeline runner for the Medical FAQ RAG System
Executes Parts 1-5 in sequence.
Usage:  python run_all.py
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from src.part1_chunking   import load_and_preprocess_corpus, chunk_corpus, print_corpus_stats
from src.part2_retrieval  import (TFIDFRetriever, DenseRetriever, HybridRetriever,
                                   compare_retrievers, MODELS_DIR)
from src.part3_generation import RAGGenerator
from src.part5_evaluation import run_evaluation


def main():
    print("\n" + "#" * 70)
    print("#  Medical FAQ RAG System — Full Pipeline")
    print("#" * 70)

    # ── Part 1: Corpus ───────────────────────────────────────────────────────
    print("\n[1/5] Corpus Preparation & Chunking")
    print("-" * 50)
    preprocessed = load_and_preprocess_corpus()
    chunks = chunk_corpus(preprocessed)
    print_corpus_stats(preprocessed, chunks)

    # ── Part 2: Retrieval ────────────────────────────────────────────────────
    print("\n[2/5] Building Retrieval Indices")
    print("-" * 50)

    tfidf = TFIDFRetriever()
    tfidf.fit(chunks)
    tfidf.save(MODELS_DIR / "tfidf_model.pkl")

    dense = DenseRetriever()
    dense.fit(chunks)
    dense.save(MODELS_DIR / "dense_model.pkl")

    hybrid = HybridRetriever(tfidf, dense)

    compare_retrievers(tfidf, dense, top_k=3)

    # ── Part 3: Generation sample ────────────────────────────────────────────
    print("\n[3/5] Generation Module — Sample Outputs")
    print("-" * 50)
    gen = RAGGenerator()
    sample_queries = [
        "What are the symptoms of diabetes?",
        "How is hypertension treated?",
        "Can anxiety disorders be cured?",
    ]
    for q in sample_queries:
        ctx, _ = hybrid.retrieve(q, top_k=3)
        result = gen.generate(q, ctx)
        gen.display_result(result)

    # ── Part 5: Evaluation ───────────────────────────────────────────────────
    print("\n[5/5] Experimental Evaluation")
    print("-" * 50)
    run_evaluation(tfidf, dense, hybrid, gen, top_k=5)

    gen.save_log()

    print("\n" + "#" * 70)
    print("#  Pipeline complete!")
    print("#  To launch the UI:  python src/part4_ui.py")
    print("#" * 70)


if __name__ == "__main__":
    main()
