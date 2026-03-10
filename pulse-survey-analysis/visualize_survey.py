#!/usr/bin/env python3
"""
Survey Results Visualization Script
Reads survey-responses/GOps Engagement Survey – Pulse Check August 2025.xlsx, sheet Form Responses 1,
outputs overall sentiment breakdown and a horizontal stacked bar chart per question.

Chart labels are derived from column headings via an LLM (OpenAI). Set OPENAI_API_KEY
for LLM summarization. Leave OPENAI_MODEL unset to auto-pick a model you have access to. Without the API,
labels fall back to truncated headings.
"""

import os
import re

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# Use a writable dir for matplotlib config (avoids issues when ~/.matplotlib is read-only)
_matplotlib_cache = os.path.join(SCRIPT_DIR, ".venv", "matplotlib-cache")
os.makedirs(_matplotlib_cache, exist_ok=True)
os.environ.setdefault("MPLCONFIGDIR", _matplotlib_cache)

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

EXCEL_FILE = os.path.join(SCRIPT_DIR, "survey-responses", "GOps Engagement Survey – Pulse Check August 2025.xlsx")
SHEET_NAME = "Form Responses 1"
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "output")

# LLM for summarizing question labels (set OPENAI_API_KEY; set OPENAI_MODEL to fix a model, or leave unset to auto-pick)
LLM_MODEL_ENV = os.environ.get("OPENAI_MODEL")
# Default only when we don't auto-pick (auto-pick happens inside the function when OPENAI_MODEL is unset)
LLM_MODEL_DEFAULT = "gpt-3.5-turbo"

# Preferred chat models in order (first one available for your account will be used when OPENAI_MODEL is unset)
PREFERRED_CHAT_MODELS = [
    "gpt-4o-mini",
    "gpt-4o",
    "gpt-4-turbo",
    "gpt-4",
    "gpt-3.5-turbo",
    "gpt-3.5-turbo-instruct",
]


def _pick_available_chat_model(client):
    """Return the first preferred chat model that the API key has access to, or a default."""
    try:
        models = client.models.list()
        available = {m.id for m in models.data}
    except Exception:
        return LLM_MODEL_DEFAULT
    for mid in PREFERRED_CHAT_MODELS:
        if mid in available:
            return mid
    for mid in sorted(available):
        if mid.startswith("gpt-"):
            return mid
    return LLM_MODEL_DEFAULT


def summarize_question_headings_with_llm(headings):
    """
    Use an LLM to turn full survey question column headings into short chart labels.
    Returns a list of short strings (one per heading), in the same order.
    Falls back to truncated heading if API is unavailable or returns wrong count.
    """
    if not headings:
        return []
    try:
        from openai import OpenAI  # type: ignore[reportMissingImports]
    except ImportError:
        print(f"OpenAI library not found, using fallback labels for LLM model")
        return [_truncate_label(h) for h in headings]
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        print("No API key found, using fallback labels for LLM model.")
        return [_truncate_label(h) for h in headings]

    client = OpenAI(api_key=api_key)
    if LLM_MODEL_ENV:
        model = LLM_MODEL_ENV
        print(f"Using OpenAI for chart labels (model: {model})")
    else:
        model = _pick_available_chat_model(client)
        print(f"Using OpenAI for chart labels (auto-selected model: {model})")

    numbered = "\n".join(f"{i+1}. {h.strip()}" for i, h in enumerate(headings))
    prompt = f"""You are helping label a survey chart. For each survey question below, write a single short label (about 4–8 words) suitable for a horizontal bar chart axis. Keep the meaning clear and concise. Do not use quotes or numbering in your answers.

Survey questions (one per line):
{numbered}

Reply with exactly one short label per question, in the same order (1 line per label). No other text."""

    try:
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
        )
        text = (response.choices[0].message.content or "").strip()
    except Exception as e:
        err_msg = str(e).lower()
        if "403" in err_msg or "permission" in err_msg or "model" in err_msg or "not have access" in err_msg:
            print(f"OpenAI model '{model}' is not available for your project. Set OPENAI_MODEL to a model you have access to, e.g.:")
            print("  export OPENAI_MODEL=gpt-3.5-turbo")
            print("Using truncated labels for this run.")
        return [_truncate_label(h) for h in headings]

    # Parse: one label per line, strip numbering and quotes
    lines = [re.sub(r"^\s*\d+[.)]\s*", "", line).strip().strip('"\'') for line in text.splitlines() if line.strip()]
    if len(lines) == len(headings):
        return [line[:50] or _truncate_label(h) for line, h in zip(lines, headings)]
    return [_truncate_label(h) for h in headings]


def _truncate_label(heading):
    """Fallback when LLM is not used or fails: first 45 chars of heading."""
    s = (heading or "").strip()
    return s[:45] + ("…" if len(s) > 45 else "")


def sentiment(value):
    """Map response text to Positive, Neutral, or Negative.
    Check Negative (disagree) before Positive (agree) so 'Disagree' is not matched by 'agree'.
    """
    if pd.isna(value) or not isinstance(value, str):
        return None
    s = value.strip().lower()
    if "strongly disagree" in s or "disagree" in s:
        return "Negative"
    if "neutral" in s:
        return "Neutral"
    if "strongly agree" in s or "agree" in s:
        return "Positive"
    return None


def extract_year_from_filename(filepath):
    """Extract a 4-digit year from the filename (e.g. '...August 2025.xlsx' -> '2025')."""
    match = re.search(r"\b(20\d{2})\b", filepath)
    return match.group(1) if match else None


def get_question_columns(df):
    """Return column names that are survey questions (exclude Timestamp, squad, optional)."""
    return [
        c for c in df.columns
        if c != "Timestamp"
        and "squad" not in c.lower()
        and "optional" not in c.lower()
        and "comments" not in c.lower()
        and "share anything" not in c.lower()
    ]


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    year = extract_year_from_filename(EXCEL_FILE)
    year_suffix = f"-{year}" if year else ""

    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    question_cols = get_question_columns(df)
    short_labels = summarize_question_headings_with_llm(question_cols)
    label_by_col = dict(zip(question_cols, short_labels))

    # Overall: pool all question responses
    all_responses = []
    for col in question_cols:
        all_responses.extend(df[col].dropna().astype(str).tolist())
    all_sentiments = [sentiment(v) for v in all_responses]
    all_sentiments = [s for s in all_sentiments if s is not None]
    n_total = len(all_sentiments)
    pct_pos = 100 * sum(1 for s in all_sentiments if s == "Positive") / n_total if n_total else 0
    pct_neu = 100 * sum(1 for s in all_sentiments if s == "Neutral") / n_total if n_total else 0
    pct_neg = 100 * sum(1 for s in all_sentiments if s == "Negative") / n_total if n_total else 0

    print(f"Overall sentiment breakdown {year_suffix}")
    print("🟩 {:.0f}% Positive (Agree / Strongly Agree)".format(pct_pos))
    print("🟨 {:.0f}% Neutral".format(pct_neu))
    print("🟥 {:.0f}% Negative (Disagree / Strongly Disagree)".format(pct_neg))

    # Per-question percentages for stacked bar chart
    labels_ordered = []
    pct_neg_list = []
    pct_neu_list = []
    pct_pos_list = []

    for col in question_cols:
        vals = df[col].dropna().astype(str)
        sents = [sentiment(v) for v in vals]
        sents = [s for s in sents if s is not None]
        n = len(sents)
        if n == 0:
            continue
        pn = 100 * sum(1 for s in sents if s == "Negative") / n
        pnu = 100 * sum(1 for s in sents if s == "Neutral") / n
        pp = 100 * sum(1 for s in sents if s == "Positive") / n
        labels_ordered.append(label_by_col.get(col, _truncate_label(col)))
        pct_neg_list.append(pn)
        pct_neu_list.append(pnu)
        pct_pos_list.append(pp)

    # Stacked horizontal bar chart (sample: red = Negative, orange = Neutral, green = Positive)
    fig, ax = plt.subplots(figsize=(16, 8))
    y_pos = np.arange(len(labels_ordered))
    left_neg = np.array(pct_neg_list)
    left_neu = left_neg + np.array(pct_neu_list)
    # bars: left-to-right order for stacking is negative, neutral, positive
    ax.barh(y_pos, pct_neg_list, color="#E24B4B", label="Negative (Disagree + Strongly Disagree)")
    ax.barh(y_pos, pct_neu_list, left=left_neg, color="#F5A623", label="Neutral")
    ax.barh(y_pos, pct_pos_list, left=left_neu, color="#7ED321", label="Positive (Agree / Strongly Agree)")
    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels_ordered, fontsize=10)
    ax.set_xlim(0, 100)
    ax.set_xlabel("Percentage of responses")
    ax.set_ylabel("")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=3, frameon=False)
    ax.set_title("Pulse Survey – Sentiment by question")
    plt.tight_layout()
    out_path = os.path.join(OUTPUT_DIR, f"responses-horizontal-graph{year_suffix}.png")
    plt.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close()
    print("\nChart saved to", out_path)


if __name__ == "__main__":
    main()
