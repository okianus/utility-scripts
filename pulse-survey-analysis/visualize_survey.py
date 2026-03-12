#!/usr/bin/env python3
"""
Survey Results Visualization Script
Reads survey-responses/GOps Engagement Survey – Pulse Check MONTH YEAR.xlsx, sheet Form Responses 1,
outputs overall sentiment breakdown, a horizontal stacked bar chart per question, and per-squad vertical bar charts (one PNG).

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

#EXCEL_FILE = os.path.join(SCRIPT_DIR, "survey-responses", "GOps Engagement Survey – Pulse Check August 2025.xlsx")
EXCEL_FILE = os.path.join(SCRIPT_DIR, "survey-responses", "GOps Engagement Survey – Pulse Check February 2026.xlsx")
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


# Month names as they appear in the survey filename (e.g. "...August 2025.xlsx")
_MONTH_NAMES = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)


def extract_month_year_from_filename(filepath):
    """Extract (month, year) from the filename only. Returns (str, str) or (None, str).
    E.g. '...February 2026.xlsx' -> ('February', '2026'); '...Survey 2025.xlsx' -> (None, '2025').
    """
    filename = os.path.basename(filepath)
    for month in _MONTH_NAMES:
        match = re.search(re.escape(month) + r"\s+(20\d{2})\b", filename, re.IGNORECASE)
        if match:
            return (month, match.group(1))
    match = re.search(r"\b(20\d{2})\b", filename)
    return (None, match.group(1)) if match else (None, None)


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


def get_squad_column(df):
    """Return the column name that contains squad (e.g. AI, Alerting, IRM, SLOs), or None."""
    for c in df.columns:
        if "squad" in c.lower():
            return c
    return None


def get_optional_comment_columns(df):
    """Return column names that are optional free-text comment fields."""
    return [
        c for c in df.columns
        if "optional" in c.lower()
        and ("explain" in c.lower() or "share" in c.lower() or "comments" in c.lower())
    ]


def _response_to_scale(value):
    """Map response text to numeric scale 1-5. Returns (scale, display_str). display_str is description only (e.g. 'Strongly Agree'), stripping leading 'N – ' if present."""
    if pd.isna(value):
        return None, ""
    s = str(value).strip()
    if not s:
        return None, ""
    # Use description only: strip leading "1 – ", "2 – ", etc.
    display_str = s.split(" – ", 1)[1].strip() if " – " in s else s
    lower = s.lower()
    if "strongly disagree" in lower:
        return 1, display_str
    if "disagree" in lower:
        return 2, display_str
    if "neutral" in lower:
        return 3, display_str
    if "agree" in lower and "strongly" in lower:
        return 5, display_str
    if "agree" in lower:
        return 4, display_str
    # Numeric fallback
    try:
        n = int(s.replace(".", "").strip())
        if 1 <= n <= 5:
            return n, display_str
    except ValueError:
        pass
    return None, display_str


def _get_optional_comment_to_question_map(df):
    """Return dict: optional_comment_col -> question_col (the question column immediately before it)."""
    all_cols = list(df.columns)
    question_cols = set(get_question_columns(df))
    optional_cols = get_optional_comment_columns(df)
    out = {}
    for opt_col in optional_cols:
        if opt_col not in all_cols:
            continue
        idx = all_cols.index(opt_col)
        for j in range(idx - 1, -1, -1):
            if all_cols[j] in question_cols:
                out[opt_col] = all_cols[j]
                break
        else:
            out[opt_col] = None
    return out


# Expected squad names for ordering in charts (others from data will be appended)
SQUAD_ORDER = ["AI", "Alerting", "IRM", "SLOs"]

# Question text (column header) -> attribute for aggregated-by-attribute chart. Must have equal questions per attribute.
QUESTION_TO_ATTRIBUTE = {
    "Over the past 6 months, I have had opportunities at work to learn and grow.": "Growth and autonomy",
    "I have received recognition or praise for doing good work.": "Recognition",
    "My fellow employees are committed to doing quality work.": "Meaningful work",
    "I have input into the goals or priorities for my team.": "Effective communication",
    "If I make a mistake on this team, it is not held against me.": "Team dynamics",
    "I know what is expected of me at work.": "Supportive leadership",
    "I understand what success looks like for Grafana, and I believe my team is focused on the right, highest-impact work to achieve it.": "Strategic clarity & leverage",
    "I have the opportunity to do what I do best every day.": "Meaningful work",
    "There is someone at work who encourages my development.": "Team dynamics",
    "I have the freedom to decide how to approach my work.": "Growth and autonomy",
    "My manager, or someone at work, seems to care about me as a person.": "Recognition",
    "I feel empowered to use AI and other tools to increase my impact and help my team move faster.": "Strategic clarity & leverage",
    "My manager gives me useful feedback to help me improve.": "Supportive leadership",
    "There are open channels for me to share ideas and concerns.": "Effective communication",
}


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    month, year = extract_month_year_from_filename(EXCEL_FILE)
    year_suffix = f"-{year}" if year else ""
    title_suffix = (
        f" – {month} {year}" if month and year else (f" – {year}" if year else "")
    )
    # Filename prefix: year-month (month lowercase) for year-month-description.png
    month_lower = month.lower() if month else ""
    filename_prefix = f"{year}-{month_lower}" if year and month_lower else (year or "output")

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

    print(f"\nOverall sentiment breakdown – {month} {year_suffix[1:]}")
    print("🟩 {:.0f}% Positive (Agree / Strongly Agree)".format(pct_pos))
    print("🟨 {:.0f}% Neutral".format(pct_neu))
    print("🟥 {:.0f}% Negative (Disagree / Strongly Disagree)".format(pct_neg))

    # Per-question percentages for stacked bar chart
    labels_ordered = []
    question_cols_with_data = []  # same order as labels_ordered
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
        question_cols_with_data.append(col)
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
    # Show positive % on each bar: inside green segment when wide enough, else at end of bar
    for i in range(len(labels_ordered)):
        pp = pct_pos_list[i]
        if pp >= 8:
            x_center = left_neu[i] + pp / 2
            ax.text(x_center, y_pos[i], f"{pp:.0f}%", ha="center", va="center", fontsize=8, color="white", fontweight="bold")
        else:
            ax.text(99, y_pos[i], f"{pp:.0f}%", ha="right", va="center", fontsize=8, color="#333")
    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels_ordered, fontsize=10)
    ax.set_xlim(0, 100)
    ax.set_xticks(np.arange(0, 101, 10))
    ax.set_xlabel("Percentage of responses")
    ax.set_ylabel("")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=3, frameon=False)
    ax.set_title("Pulse Survey – Sentiment by question" + title_suffix)
    plt.tight_layout()
    out_path = os.path.join(OUTPUT_DIR, f"{filename_prefix}-responses-by-question.png")
    plt.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close()
    print("Chart saved to", out_path)
    png_paths = [out_path]

    # Per-question, per-squad vertical bar charts in one file (3 rows × N columns)
    squad_col = get_squad_column(df)
    if squad_col is None:
        print("No squad column found; skipping per-squad charts.")
    else:
        # Build ordered list of squads (prefer SQUAD_ORDER, then any others from data)
        squads_in_data = df[squad_col].dropna().astype(str).str.strip().unique().tolist()
        squads_ordered = [s for s in SQUAD_ORDER if s in squads_in_data]
        for s in squads_in_data:
            if s not in squads_ordered:
                squads_ordered.append(s)

        n_q = len(labels_ordered)
        n_cols = max(1, (n_q + 2) // 3)  # 3 rows, as many columns as needed
        n_rows = 3
        fig2, axes = plt.subplots(n_rows, n_cols, figsize=(4 * n_cols, 4 * n_rows))
        if n_q == 1:
            axes = np.array([[axes]])
        elif axes.ndim == 1:
            axes = axes.reshape(1, -1)
        axes_flat = axes.flatten()

        for idx, col in enumerate(question_cols_with_data):
            if idx >= len(labels_ordered):
                break
            ax = axes_flat[idx]
            row = df[[col, squad_col]].dropna(subset=[col])
            row = row[row[squad_col].astype(str).str.strip().isin(squads_ordered)]
            neg_list, neu_list, pos_list = [], [], []
            for squad in squads_ordered:
                subset = row[row[squad_col].astype(str).str.strip() == squad]
                vals = subset[col].astype(str)
                sents = [s for s in [sentiment(v) for v in vals] if s is not None]
                n = len(sents)
                if n == 0:
                    neg_list.append(0)
                    neu_list.append(0)
                    pos_list.append(0)
                else:
                    neg_list.append(100 * sum(1 for s in sents if s == "Negative") / n)
                    neu_list.append(100 * sum(1 for s in sents if s == "Neutral") / n)
                    pos_list.append(100 * sum(1 for s in sents if s == "Positive") / n)
            x = np.arange(len(squads_ordered))
            w = 0.6
            ax.bar(x, neg_list, width=w, color="#E24B4B", label="Negative")
            ax.bar(x, neu_list, width=w, bottom=neg_list, color="#F5A623", label="Neutral")
            ax.bar(x, pos_list, width=w, bottom=np.array(neg_list) + np.array(neu_list), color="#7ED321", label="Positive")
            ax.set_xticks(x)
            ax.set_xticklabels(squads_ordered, rotation=0)
            ax.set_ylabel("Percentage")
            ax.set_ylim(0, 100)
            ax.set_title(labels_ordered[idx][:40] + ("…" if len(labels_ordered[idx]) > 40 else ""), fontsize=9)
            if idx == 0:
                ax.legend(loc="upper right", fontsize=7)

        for j in range(len(labels_ordered), len(axes_flat)):
            axes_flat[j].set_visible(False)
        plt.suptitle("Pulse Survey – Sentiment by question and squad" + title_suffix, y=1.02, fontsize=12)
        plt.tight_layout()
        out_path_squad = os.path.join(OUTPUT_DIR, f"{filename_prefix}-responses-by-squad.png")
        plt.savefig(out_path_squad, dpi=150, bbox_inches="tight")
        plt.close()
        print("Chart saved to", out_path_squad)
        png_paths.append(out_path_squad)

    # By-attribute chart: map questions to attributes, aggregate, require equal questions per attribute
    attr_to_questions = {}  # attribute -> list of column names (from question_cols_with_data)
    for col in question_cols_with_data:
        key = col.strip()
        attr = QUESTION_TO_ATTRIBUTE.get(key)
        if attr is None:
            # Try match by substring in case Excel trims or slightly differs
            for q, a in QUESTION_TO_ATTRIBUTE.items():
                if q.strip() in key or key in q.strip():
                    attr = a
                    break
        if attr is not None:
            attr_to_questions.setdefault(attr, []).append(col)
    mapped_count = sum(len(qs) for qs in attr_to_questions.values())
    if mapped_count != len(question_cols_with_data):
        unmapped = [c for c in question_cols_with_data if not any(c.strip() == q.strip() or c.strip() in q.strip() or q.strip() in c.strip() for q in QUESTION_TO_ATTRIBUTE)]
        raise SystemExit(
            f"Error: Every question must map to an attribute. {len(question_cols_with_data) - mapped_count} question(s) have no mapping. Unmapped columns (sample): {unmapped[:3]}"
        )
    counts = [len(qs) for qs in attr_to_questions.values()]
    if not counts or len(set(counts)) != 1:
        raise SystemExit(
            "Error: Each attribute must have the same number of questions. "
            f"Current per-attribute counts: {dict((a, len(qs)) for a, qs in attr_to_questions.items())}"
        )
    # Aggregate sentiment by attribute (pool all responses for questions in that attribute)
    attr_order = sorted(attr_to_questions.keys(), key=str.lower)
    attr_neg, attr_neu, attr_pos = [], [], []
    for attr in attr_order:
        cols = attr_to_questions[attr]
        all_sents = []
        for col in cols:
            vals = df[col].dropna().astype(str)
            all_sents.extend([s for s in [sentiment(v) for v in vals] if s is not None])
        n = len(all_sents)
        if n == 0:
            attr_neg.append(0)
            attr_neu.append(0)
            attr_pos.append(0)
        else:
            attr_neg.append(100 * sum(1 for s in all_sents if s == "Negative") / n)
            attr_neu.append(100 * sum(1 for s in all_sents if s == "Neutral") / n)
            attr_pos.append(100 * sum(1 for s in all_sents if s == "Positive") / n)
    fig3, ax3 = plt.subplots(figsize=(12, max(5, len(attr_order) * 0.6)))
    y_attr = np.arange(len(attr_order))
    left_neg = np.array(attr_neg)
    left_neu = left_neg + np.array(attr_neu)
    ax3.barh(y_attr, attr_neg, color="#E24B4B", label="Negative (Disagree + Strongly Disagree)")
    ax3.barh(y_attr, attr_neu, left=left_neg, color="#F5A623", label="Neutral")
    ax3.barh(y_attr, attr_pos, left=left_neu, color="#7ED321", label="Positive (Agree / Strongly Agree)")
    # Show positive % on each bar: inside green segment when wide enough, else at end of bar
    for i in range(len(attr_order)):
        pp = attr_pos[i]
        if pp >= 8:
            x_center = left_neu[i] + pp / 2
            ax3.text(x_center, y_attr[i], f"{pp:.0f}%", ha="center", va="center", fontsize=8, color="white", fontweight="bold")
        else:
            ax3.text(99, y_attr[i], f"{pp:.0f}%", ha="right", va="center", fontsize=8, color="#333")
    ax3.set_yticks(y_attr)
    ax3.set_yticklabels(attr_order, fontsize=10)
    ax3.set_xlim(0, 100)
    ax3.set_xticks(np.arange(0, 101, 10))
    ax3.set_xlabel("Percentage of responses")
    ax3.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=3, frameon=False)
    ax3.set_title("Pulse Survey – Sentiment by attribute (department aggregate)" + title_suffix)
    plt.tight_layout()
    out_path_attr = os.path.join(OUTPUT_DIR, f"{filename_prefix}-responses-by-attribute.png")
    plt.savefig(out_path_attr, dpi=150, bbox_inches="tight")
    plt.close()
    print("Chart saved to", out_path_attr)
    png_paths.append(out_path_attr)

    # Radar chart: same by-attribute data (positive % per attribute; higher = better)
    num_attrs = len(attr_order)
    angles = np.linspace(0, 2 * np.pi, num_attrs, endpoint=False).tolist()
    angles += angles[:1]
    values = attr_pos + attr_pos[:1]
    fig4, ax4 = plt.subplots(figsize=(8, 8), subplot_kw=dict(projection="polar"))
    ax4.plot(angles, values, "o-", color="#7ED321", linewidth=2)
    ax4.fill(angles, values, color="#7ED321", alpha=0.25)
    for angle, val in zip(angles[:-1], attr_pos):
        ax4.text(angle, val, f"{round(val)}%", ha="center", va="center", fontsize=9)
    ax4.set_xticks(angles[:-1])
    ax4.set_xticklabels(attr_order, fontsize=9)
    ax4.set_ylim(0, 100)
    ax4.set_title("Pulse Survey – Positive % by attribute (radar)" + title_suffix, pad=20)
    plt.tight_layout()
    out_path_radar = os.path.join(OUTPUT_DIR, f"{filename_prefix}-responses-by-attribute-radar.png")
    plt.savefig(out_path_radar, dpi=150, bbox_inches="tight")
    plt.close()
    print("Chart saved to", out_path_radar)
    png_paths.append(out_path_radar)

    # Step 7: Optional comments grouped by sentiment (from responder's rating 1-5), [Squad – question, response: X] prefix
    squad_col = get_squad_column(df)
    grouped_comments, additional_thoughts = _run_optional_comments_by_rating(df, squad_col, label_by_col)

    # LLM analysis: 5 things working well, 5 areas for improvement
    analysis_text = _run_llm_analysis(
        df=df,
        question_cols_with_data=question_cols_with_data,
        labels_ordered=labels_ordered,
        label_by_col=label_by_col,
        pct_neg_list=pct_neg_list,
        pct_neu_list=pct_neu_list,
        pct_pos_list=pct_pos_list,
        attr_order=attr_order,
        attr_neg=attr_neg,
        attr_neu=attr_neu,
        attr_pos=attr_pos,
    )

    # Step 9: PDF with PNGs, comments, and analysis (YEAR-MONTH-gops-survey-summary.pdf)
    pdf_path = os.path.join(OUTPUT_DIR, f"{filename_prefix}-gops-survey-summary.pdf")
    _build_summary_pdf(png_paths, grouped_comments, analysis_text, pdf_path, month=month, year=year, additional_thoughts=additional_thoughts)


def _run_llm_analysis(
    df,
    question_cols_with_data,
    labels_ordered,
    label_by_col,
    pct_neg_list,
    pct_neu_list,
    pct_pos_list,
    attr_order,
    attr_neg,
    attr_neu,
    attr_pos,
):
    """Call LLM to report 5 things working well and 5 areas for improvement."""
    try:
        from openai import OpenAI  # type: ignore[reportMissingImports]
    except ImportError:
        print("Skipping LLM analysis (openai not installed).")
        return
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        print("Skipping LLM analysis (OPENAI_API_KEY not set).")
        return

    client = OpenAI(api_key=api_key)
    model = LLM_MODEL_ENV or _pick_available_chat_model(client)

    # Build a concise data summary for the LLM
    lines = ["## Overall: department-wide sentiment by question (Negative % | Neutral % | Positive %)"]
    for i, col in enumerate(question_cols_with_data):
        if i < len(labels_ordered) and i < len(pct_neg_list):
            label = labels_ordered[i]
            lines.append(f"- {label}: {pct_neg_list[i]:.0f}% | {pct_neu_list[i]:.0f}% | {pct_pos_list[i]:.0f}%")
    lines.append("")
    lines.append("## By attribute (department aggregate):")
    for j, attr in enumerate(attr_order):
        if j < len(attr_neg):
            lines.append(f"- {attr}: Neg {attr_neg[j]:.0f}% | Neu {attr_neu[j]:.0f}% | Pos {attr_pos[j]:.0f}%")
    data_summary = "\n".join(lines)

    prompt = f"""You are analyzing engagement pulse survey results. Based on the data below, list exactly 5 things that are working well and 5 areas for improvement. Be specific and concise (1-2 sentences each). Use the exact headings "5 things working well" and "5 areas for improvement".

Survey data summary:
{data_summary}
"""

    try:
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
        )
        text = (response.choices[0].message.content or "").strip()
        print("\n--- LLM Analysis ---\n")
        print(text)
        print("\n--- End LLM Analysis ---")
        return text
    except Exception as e:
        print(f"Skipping LLM analysis (API error: {e}).")
        return None


def _run_optional_comments_by_rating(df, squad_col, label_by_col):
    """Collect optional comments; group by sentiment from the responder's rating for that question:
    Positive = 4 or 5, Neutral = 3, Negative = 1 or 2.
    The last column of the spreadsheet is treated as overall survey comments (additional thoughts).
    Returns (grouped, additional_thoughts): grouped = dict Positive/Neutral/Negative -> list of (squad, comment, short_question, response); additional_thoughts = list of (squad, comment)."""
    optional_cols = get_optional_comment_columns(df)
    if not optional_cols:
        print("No optional comment columns found.")
        return {"Positive": [], "Neutral": [], "Negative": []}, []

    all_cols = list(df.columns)
    overall_col = all_cols[-1] if (all_cols and all_cols[-1] in optional_cols) else None
    per_question_cols = [c for c in optional_cols if c != overall_col]

    opt_to_question = _get_optional_comment_to_question_map(df)
    squad_vals = df[squad_col].fillna("").astype(str).str.strip() if squad_col else [""] * len(df)
    grouped = {"Positive": [], "Neutral": [], "Negative": []}
    additional_thoughts = []

    for idx in range(len(df)):
        squad = squad_vals.iloc[idx] if squad_col else "Unknown"
        for opt_col in per_question_cols:
            val = df[opt_col].iloc[idx]
            if pd.isna(val) or not str(val).strip():
                continue
            comment = str(val).strip()
            question_col = opt_to_question.get(opt_col)
            if question_col is None:
                short_question = "Unknown question"
                response_display = ""
                scale = None
            else:
                short_question = label_by_col.get(question_col, _truncate_label(question_col))
                response_val = df[question_col].iloc[idx]
                scale, response_display = _response_to_scale(response_val)
            # 1,2 -> Negative; 3 -> Neutral; 4,5 -> Positive; unparseable -> Neutral
            if scale is not None:
                if scale in (1, 2):
                    sent = "Negative"
                elif scale == 3:
                    sent = "Neutral"
                else:
                    sent = "Positive"
            else:
                sent = "Neutral"
            grouped[sent].append((squad, comment, short_question, response_display))

        if overall_col is not None:
            val = df[overall_col].iloc[idx]
            if pd.notna(val) and str(val).strip():
                additional_thoughts.append((squad, str(val).strip()))

    for sent in grouped:
        grouped[sent] = sorted(grouped[sent], key=lambda x: (x[0].lower(), x[2], x[1]))
    additional_thoughts.sort(key=lambda x: (x[0].lower(), x[1]))

    print("\n--- Optional comments (grouped by responder rating) ---")
    for sent in ("Positive", "Neutral", "Negative"):
        print(f"\n{sent}:")
        for squad, comment, short_question, response_display in grouped[sent]:
            prefix = f"[{squad} – {short_question}, response: {response_display}]" if (short_question or response_display) else f"[{squad}]"
            print(f"  {prefix} {comment}")
    if additional_thoughts:
        print("\nAdditional thoughts:")
        for squad, comment in additional_thoughts:
            print(f"  [{squad}] {comment}")
    print("\n--- End optional comments ---")
    return grouped, additional_thoughts


def _build_summary_pdf(png_paths, grouped_comments, analysis_text, output_path, month=None, year=None, additional_thoughts=None):
    """Write YEAR-MONTH-gops-survey-summary.pdf with PNGs, comments, and analysis. additional_thoughts = list of (squad, comment) for the overall survey comments subsection."""
    try:
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.units import inch
        from reportlab.pdfgen import canvas
        from reportlab.lib.utils import ImageReader
    except ImportError:
        print("Skipping PDF (reportlab not installed).")
        return
    c = canvas.Canvas(output_path, pagesize=letter)
    w, h = letter
    margin = inch * 0.75
    margin_bottom = margin
    y = h - margin

    def draw_text_block(canvas_obj, text, x, y_ref, max_width, font_size=9):
        """Draw wrapped text; start new page if below margin_bottom. Returns new y. Restore font after page break."""
        canvas_obj.setFont("Helvetica", font_size)
        y_pos = y_ref
        line_height = max(10, font_size + 2)
        for para in text.replace("\r", "").split("\n"):
            lines = []
            for word in para.split():
                lines.append(word)
                if canvas_obj.stringWidth(" ".join(lines), "Helvetica", font_size) > max_width:
                    if len(lines) > 1:
                        lines.pop()
                        if y_pos - line_height < margin_bottom:
                            canvas_obj.showPage()
                            canvas_obj.setFont("Helvetica", font_size)
                            y_pos = h - margin
                        y_pos -= line_height
                        canvas_obj.drawString(x, y_pos, " ".join(lines))
                        lines = [word]
                    else:
                        if y_pos - line_height < margin_bottom:
                            canvas_obj.showPage()
                            canvas_obj.setFont("Helvetica", font_size)
                            y_pos = h - margin
                        y_pos -= line_height
                        canvas_obj.drawString(x, y_pos, lines[0])
                        lines = []
            if lines:
                if y_pos - line_height < margin_bottom:
                    canvas_obj.showPage()
                    canvas_obj.setFont("Helvetica", font_size)
                    y_pos = h - margin
                y_pos -= line_height
                canvas_obj.drawString(x, y_pos, " ".join(lines))
            y_pos -= 4
        return y_pos

    def draw_analysis_with_bold(canvas_obj, text, x, y_ref, max_width):
        """Draw analysis text, rendering **...** as bold. One block per numbered item, line break before each new #."""
        # Normalize newlines to spaces so we can split by " N. " pattern
        normalized = re.sub(r"\s+", " ", text).strip()
        # Split on " 1. ", " 2. ", etc. (space + digits + period + space) to get one chunk per numbered item; keep the delimiter so we can attach it
        parts = re.split(r"\s+(\d+\.)\s+", normalized)
        # parts[0] = optional header (no leading number), then alternating: "1.", "text for 1", "2.", "text for 2", ...
        blocks = []
        if parts and parts[0].strip():
            blocks.append(parts[0].strip())
        for i in range(1, len(parts) - 1, 2):
            if i + 1 < len(parts):
                blocks.append(parts[i] + " " + parts[i + 1].strip())
        # Split any block that contains "5 areas for improvement" so that phrase starts its own block (enables 2 newlines before it)
        expanded = []
        for block in blocks:
            bl = block.lower()
            if "5 areas for improvement" in bl:
                i = bl.find("5 areas for improvement")
                before = block[:i].strip()
                after = block[i:].strip()
                if before:
                    expanded.append(before)
                expanded.append(after)
            else:
                expanded.append(block)
        blocks = expanded
        y_pos = y_ref
        font_size = 9
        line_height = font_size + 2
        space_width = canvas_obj.stringWidth(" ", "Helvetica", font_size)
        for idx, block in enumerate(blocks):
            # Line break before each new item (except the first)
            if idx > 0:
                y_pos -= 10
            # One extra newline before the "5 things working well" heading; two before "5 areas for improvement"
            block_starts = block.strip().lower()
            if block_starts.startswith("5 things working well"):
                y_pos -= 20  # two extra lines before this heading
            if block_starts.startswith("5 areas for improvement"):
                y_pos -= 20
            # Section headings: draw entire block in bold (they may lack ** in source after block split)
            is_section_heading = block_starts.startswith("5 things working well") or block_starts.startswith("5 areas for improvement")
            if is_section_heading:
                clean = block.replace("**", "").strip()
                word_stream = [(w, True) for w in clean.split() if w]
            else:
                # Split by ** for bold; strip any literal ** so they don't appear in the PDF
                segments = re.split(r"\*\*(.+?)\*\*", block)
                word_stream = []
                for i, part in enumerate(segments):
                    part = part.replace("**", "").strip()
                    if not part:
                        continue
                    is_bold = i % 2 == 1
                    for word in part.split():
                        if word != "**":
                            word_stream.append((word, is_bold))
            # Build lines: each line is a list of (word, is_bold); wrap when line width would exceed max_width
            lines_to_draw = []
            current_line = []
            current_width = 0
            for word, is_bold in word_stream:
                w = canvas_obj.stringWidth(word, "Helvetica-Bold" if is_bold else "Helvetica", font_size)
                if current_line and current_width + space_width + w > max_width:
                    lines_to_draw.append(current_line)
                    current_line = [(word, is_bold)]
                    current_width = w
                else:
                    if current_line:
                        current_width += space_width
                    current_line.append((word, is_bold))
                    current_width += w
            if current_line:
                lines_to_draw.append(current_line)
            # Draw each line (mixed bold/normal) at same y, then move down
            for line_tuples in lines_to_draw:
                if y_pos - line_height < margin_bottom:
                    canvas_obj.showPage()
                    canvas_obj.setFont("Helvetica", font_size)
                    y_pos = h - margin
                x_pos = x
                for word, is_bold in line_tuples:
                    canvas_obj.setFont("Helvetica-Bold", font_size) if is_bold else canvas_obj.setFont("Helvetica", font_size)
                    canvas_obj.drawString(x_pos, y_pos, word)
                    x_pos += canvas_obj.stringWidth(word, "Helvetica-Bold" if is_bold else "Helvetica", font_size) + space_width
                y_pos -= line_height
        return y_pos

    # Title: GOps Department Pulse Survey Summary – MONTH YEAR
    title = "GOps Department Pulse Survey Summary" + (f" – {month} {year}" if month and year else (f" – {year}" if year else ""))
    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin, y, title)
    y -= 24

    # Images (no heading)
    for path in png_paths:
        if not os.path.isfile(path):
            continue
        y -= 6
        if y < margin + 2 * inch:
            c.showPage()
            y = h - margin
        try:
            img = ImageReader(path)
            iw, ih = img.getSize()
            scale = min((w - 2 * margin) / iw, (y - margin - 36) / ih, 1.0)
            nw, nh = iw * scale, ih * scale
            c.drawImage(path, margin, y - nh, width=nw, height=nh)
            y -= nh + 12
        except Exception as e:
            c.setFont("Helvetica", 9)
            c.drawString(margin, y, f"[Image error: {path}]")
            y -= 14

    # Comments section
    y -= 12
    if y < margin + inch:
        c.showPage()
        y = h - margin
    c.setFont("Helvetica-Bold", 11)
    c.drawString(margin, y, "Optional comments (grouped by responder rating)")
    y -= 18
    for sent in ("Positive", "Neutral", "Negative"):
        items = grouped_comments.get(sent, [])
        if not items:
            continue
        y -= 12  # newline before each sentiment subheading
        if y < margin + 1.5 * inch:
            c.showPage()
            y = h - margin
        c.setFont("Helvetica-Bold", 9)
        c.drawString(margin, y, f"{sent}:")
        y -= 12
        c.setFont("Helvetica", 8)
        for num, item in enumerate(items, 1):
            squad, comment, short_question, response_display = (item if len(item) == 4 else (item[0], item[1], "", ""))
            prefix = f"[{squad} – {short_question}, response: {response_display}]" if (short_question or response_display) else f"[{squad}]"
            line = f"{num}. {prefix} {comment}"
            y = draw_text_block(c, line, margin, y, w - 2 * margin, font_size=8)
            y -= 4
        y -= 6
    y -= 12

    # Additional thoughts (overall survey comments from last column)
    if additional_thoughts:
        y -= 12
        if y < margin + 1.5 * inch:
            c.showPage()
            y = h - margin
        c.setFont("Helvetica-Bold", 9)
        c.drawString(margin, y, "Additional thoughts:")
        y -= 12
        c.setFont("Helvetica", 8)
        for num, (squad, comment) in enumerate(additional_thoughts, 1):
            line = f"{num}. [{squad}] {comment}"
            y = draw_text_block(c, line, margin, y, w - 2 * margin, font_size=8)
            y -= 4
        y -= 12

    # Analysis section (render **...** as bold, no raw markdown)
    if analysis_text:
        y -= 12  # newline before Analysis heading
        if y < margin + inch:
            c.showPage()
            y = h - margin
        c.setFont("Helvetica-Bold", 11)
        c.drawString(margin, y, "Analysis")
        y -= 18
        y -= 12  # extra line break after Analysis heading
        y = draw_analysis_with_bold(c, analysis_text, margin, y, w - 2 * margin)

    c.save()
    print("PDF saved to", output_path)


if __name__ == "__main__":
    main()
