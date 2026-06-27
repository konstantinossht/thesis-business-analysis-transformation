"""
Blind Review Analysis Script
Πτυχιακή: Μετασχηματισμός της Επιχειρηματικής Ανάλυσης μέσω AI

Σκοπός:
- Ανάλυση δύο blind reviews από senior Business Analysts.
- Υπολογισμός total scores, quality %, Manual vs AI comparison.
- Υπολογισμός inter-rater reliability:
  exact agreement, within-one agreement, weighted Cohen's κ,
  mean absolute difference.
- Εξαγωγή συγκεντρωτικών CSV/Excel και γραφημάτων.

Αρχεία εισόδου:
- Αξιολόγηση_παραδοτέων 1.xlsx
- Αξιολόγηση_παραδοτέων2.xlsx

Αναμενόμενη δομή Excel:
Η πρώτη πραγματική γραμμή περιέχει headers:
Κριτήριο αξιολόγησης | Περιγραφή | Scope A | Scope B | FRS A | FRS B | Σχόλια αξιολογητή

"""

from __future__ import annotations

import os
import re
import warnings
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

try:
    import seaborn as sns
except ImportError:
    sns = None

from sklearn.metrics import cohen_kappa_score

warnings.filterwarnings("ignore")


BASE_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()

REVIEWER_FILES = {
    "Reviewer_1": BASE_DIR / "Αξιολόγηση_παραδοτέων 1.xlsx",
    "Reviewer_2": BASE_DIR / "Αξιολόγηση_παραδοτέων2.xlsx",
}

OUTPUT_DIR = BASE_DIR / "blind_review_results"
OUTPUT_DIR.mkdir(exist_ok=True)

DOCUMENT_MAPPING = {
    "Scope A": {"Method": "Manual", "Document_Type": "Scope"},
    "Scope B": {"Method": "AI-generated", "Document_Type": "Scope"},
    "FRS A": {"Method": "Manual", "Document_Type": "FRS"},
    "FRS B": {"Method": "AI-generated", "Document_Type": "FRS"},
}

MAX_SCORE_PER_DOCUMENT = 50  # 10 κριτήρια x 5 μέγιστος βαθμός

def normalize_text(value) -> str:
    """Καθαρίζει κείμενο για πιο εύκολη αναγνώριση στηλών."""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def extract_criterion_short_name(raw_criterion: str) -> str:
    """
    Από 'Πληρότητα / Completeness' κρατάει το αγγλικό μέρος αν υπάρχει,
    αλλιώς επιστρέφει καθαρισμένη ελληνική ονομασία.
    """
    raw_criterion = normalize_text(raw_criterion)
    if "/" in raw_criterion:
        return normalize_text(raw_criterion.split("/")[-1])
    return raw_criterion


def read_review_file(file_path: Path, reviewer_name: str) -> pd.DataFrame:
    """
    Διαβάζει ένα Excel αξιολόγησης και επιστρέφει standardized wide dataframe:
    Criterion | Description | Scope A | Scope B | FRS A | FRS B | Comments
    """
    if not file_path.exists():
        raise FileNotFoundError(f"Δεν βρέθηκε το αρχείο: {file_path}")

    # Διαβάζουμε χωρίς header ώστε να εντοπίσουμε τη γραμμή των πραγματικών headers.
    raw = pd.read_excel(file_path, header=None)

    # Εντοπισμός γραμμής που περιέχει 'Κριτήριο αξιολόγησης'.
    header_row_idx = None
    for idx, row in raw.iterrows():
        row_text = " ".join(normalize_text(x) for x in row.tolist())
        if "Κριτήριο αξιολόγησης" in row_text or "criterion" in row_text.lower():
            header_row_idx = idx
            break

    if header_row_idx is None:
        # Fallback: συνήθως είναι η πρώτη γραμμή.
        header_row_idx = 0

    headers = raw.iloc[header_row_idx].tolist()
    data = raw.iloc[header_row_idx + 1:].copy()
    data.columns = headers

    # Αφαίρεση εντελώς κενών γραμμών.
    data = data.dropna(how="all").reset_index(drop=True)

    # Αυτόματη αναγνώριση στηλών.
    cols = list(data.columns)
    col_text = {c: normalize_text(c) for c in cols}

    criterion_col = None
    description_col = None
    comments_col = None
    document_cols = {}

    for c in cols:
        txt = col_text[c]
        txt_low = txt.lower()

        if criterion_col is None and ("κριτήριο" in txt_low or "criterion" in txt_low):
            criterion_col = c
        elif description_col is None and ("περιγραφή" in txt_low or "description" in txt_low):
            description_col = c
        elif comments_col is None and ("σχόλια" in txt_low or "comment" in txt_low):
            comments_col = c

        for doc in DOCUMENT_MAPPING.keys():
            if txt == doc:
                document_cols[doc] = c

    if criterion_col is None:
        criterion_col = cols[0]
    if description_col is None and len(cols) > 1:
        description_col = cols[1]
    if comments_col is None and len(cols) > 6:
        comments_col = cols[6]

    if len(document_cols) < 4:
        positional_docs = ["Scope A", "Scope B", "FRS A", "FRS B"]
        for doc, pos in zip(positional_docs, [2, 3, 4, 5]):
            if pos < len(cols):
                document_cols[doc] = cols[pos]

    standardized = pd.DataFrame()
    standardized["Criterion_Raw"] = data[criterion_col].map(normalize_text)
    standardized["Criterion"] = standardized["Criterion_Raw"].map(extract_criterion_short_name)
    standardized["Description"] = data[description_col].map(normalize_text) if description_col is not None else ""

    for doc in ["Scope A", "Scope B", "FRS A", "FRS B"]:
        standardized[doc] = pd.to_numeric(data[document_cols[doc]], errors="coerce")

    standardized["Comments"] = data[comments_col].map(normalize_text) if comments_col is not None else ""
    standardized["Reviewer"] = reviewer_name

    # Κρατάμε μόνο γραμμές με κριτήριο και τουλάχιστον ένα score.
    score_cols = ["Scope A", "Scope B", "FRS A", "FRS B"]
    standardized = standardized[
        (standardized["Criterion_Raw"] != "") &
        (standardized[score_cols].notna().any(axis=1))
    ].reset_index(drop=True)

    return standardized


def wide_to_long(df_wide: pd.DataFrame) -> pd.DataFrame:
    """Μετατρέπει wide format σε tidy long format."""
    id_vars = ["Reviewer", "Criterion_Raw", "Criterion", "Description", "Comments"]
    value_vars = ["Scope A", "Scope B", "FRS A", "FRS B"]

    long_df = df_wide.melt(
        id_vars=id_vars,
        value_vars=value_vars,
        var_name="Document",
        value_name="Score"
    )

    long_df["Score"] = pd.to_numeric(long_df["Score"], errors="coerce")
    long_df = long_df.dropna(subset=["Score"]).copy()
    long_df["Score"] = long_df["Score"].astype(int)

    long_df["Method"] = long_df["Document"].map(lambda d: DOCUMENT_MAPPING[d]["Method"])
    long_df["Document_Type"] = long_df["Document"].map(lambda d: DOCUMENT_MAPPING[d]["Document_Type"])

    return long_df[
        ["Reviewer", "Document", "Method", "Document_Type",
         "Criterion", "Criterion_Raw", "Description", "Score", "Comments"]
    ]


def significance_free_interpretation(diff: float) -> str:
    """Πρακτική ερμηνεία διαφορών AI minus Manual."""
    if abs(diff) < 0.5:
        return "Πρακτικά ισοδύναμη ποιότητα"
    if diff > 0:
        return "Υπέρ AI-generated"
    return "Υπέρ manually-written"


def calculate_inter_rater_reliability(long_df: pd.DataFrame) -> pd.DataFrame:
    """
    Υπολογίζει exact agreement, within-one agreement, weighted Cohen's kappa
    και mean absolute difference για όλες τις αντιστοιχισμένες βαθμολογίες.
    """
    pivot = long_df.pivot_table(
        index=["Document", "Criterion"],
        columns="Reviewer",
        values="Score",
        aggfunc="first"
    ).reset_index()

    reviewer_cols = [c for c in pivot.columns if str(c).startswith("Reviewer")]
    if len(reviewer_cols) != 2:
        raise ValueError(f"Αναμένονται ακριβώς 2 reviewers, βρέθηκαν: {reviewer_cols}")

    r1, r2 = reviewer_cols
    paired = pivot.dropna(subset=[r1, r2]).copy()

    y1 = paired[r1].astype(int).values
    y2 = paired[r2].astype(int).values

    exact_agreement = np.mean(y1 == y2) * 100
    within_one_agreement = np.mean(np.abs(y1 - y2) <= 1) * 100
    mean_abs_diff = np.mean(np.abs(y1 - y2))

    kappa_quadratic = cohen_kappa_score(y1, y2, weights="quadratic")
    kappa_linear = cohen_kappa_score(y1, y2, weights="linear")
    kappa_unweighted = cohen_kappa_score(y1, y2)

    out = pd.DataFrame([{
        "Exact_Agreement_Percentage": round(exact_agreement, 2),
        "Within_One_Agreement_Percentage": round(within_one_agreement, 2),
        "Weighted_Cohen_Kappa_Quadratic": round(kappa_quadratic, 3),
        "Weighted_Cohen_Kappa_Linear": round(kappa_linear, 3),
        "Unweighted_Cohen_Kappa": round(kappa_unweighted, 3),
        "Mean_Absolute_Difference": round(mean_abs_diff, 3),
        "Number_of_Rated_Items": len(paired)
    }])

    return out

def run_analysis() -> Dict[str, pd.DataFrame]:
    """Εκτελεί όλη την ανάλυση και επιστρέφει τα βασικά outputs."""
    all_long = []

    for reviewer, file_path in REVIEWER_FILES.items():
        wide = read_review_file(file_path, reviewer)
        long = wide_to_long(wide)
        all_long.append(long)

    long_df = pd.concat(all_long, ignore_index=True)


    totals = (
        long_df
        .groupby(["Reviewer", "Document", "Method", "Document_Type"], as_index=False)
        .agg(Total_Score=("Score", "sum"), Mean_Score=("Score", "mean"))
    )

    document_summary = (
        totals
        .pivot_table(
            index=["Document", "Method", "Document_Type"],
            columns="Reviewer",
            values="Total_Score",
            aggfunc="first"
        )
        .reset_index()
    )

    reviewer_cols = [c for c in document_summary.columns if str(c).startswith("Reviewer")]
    document_summary["Mean_Total"] = document_summary[reviewer_cols].mean(axis=1)
    document_summary["Quality_Percentage"] = document_summary["Mean_Total"] / MAX_SCORE_PER_DOCUMENT * 100
    document_summary["Mean_Score"] = document_summary["Mean_Total"] / 10

    # Καθαρό naming στηλών.
    for col in reviewer_cols:
        document_summary.rename(columns={col: f"{col}_Total"}, inplace=True)

    document_summary = document_summary.sort_values(["Document_Type", "Method"], ascending=[False, False])

 
    criterion_summary = (
        long_df
        .pivot_table(
            index=["Document", "Method", "Document_Type", "Criterion"],
            columns="Reviewer",
            values="Score",
            aggfunc="first"
        )
        .reset_index()
    )

    reviewer_cols_crit = [c for c in criterion_summary.columns if str(c).startswith("Reviewer")]
    criterion_summary["Mean_Score"] = criterion_summary[reviewer_cols_crit].mean(axis=1)

    if len(reviewer_cols_crit) == 2:
        criterion_summary["Difference_R2_R1"] = (
            criterion_summary[reviewer_cols_crit[1]] - criterion_summary[reviewer_cols_crit[0]]
        )


    method_pivot = (
        document_summary
        .pivot_table(
            index="Document_Type",
            columns="Method",
            values="Mean_Total",
            aggfunc="first"
        )
        .reset_index()
    )

    method_pivot["AI_minus_Manual"] = method_pivot["AI-generated"] - method_pivot["Manual"]
    method_pivot["Interpretation"] = method_pivot["AI_minus_Manual"].map(significance_free_interpretation)
    method_comparison = method_pivot.rename(columns={
        "Manual": "Manual_Mean_Total",
        "AI-generated": "AI_Mean_Total"
    })


    criterion_method = (
        criterion_summary
        .groupby(["Document_Type", "Method", "Criterion"], as_index=False)
        .agg(Mean_Score=("Mean_Score", "mean"))
    )

    criterion_method_pivot = (
        criterion_method
        .pivot_table(
            index=["Document_Type", "Criterion"],
            columns="Method",
            values="Mean_Score",
            aggfunc="first"
        )
        .reset_index()
    )

    criterion_method_pivot["AI_minus_Manual"] = (
        criterion_method_pivot["AI-generated"] - criterion_method_pivot["Manual"]
    )

    # E. Inter-rater reliability
    inter_rater = calculate_inter_rater_reliability(long_df)

    # Save outputs
    outputs = {
        "long_data": long_df,
        "document_summary": document_summary,
        "criterion_summary": criterion_summary,
        "method_comparison": method_comparison,
        "criterion_method_comparison": criterion_method_pivot,
        "inter_rater_reliability": inter_rater,
    }

    long_df.to_csv(OUTPUT_DIR / "long_data.csv", index=False, encoding="utf-8-sig")
    document_summary.to_csv(OUTPUT_DIR / "document_summary.csv", index=False, encoding="utf-8-sig")
    criterion_summary.to_csv(OUTPUT_DIR / "criterion_summary.csv", index=False, encoding="utf-8-sig")
    method_comparison.to_csv(OUTPUT_DIR / "method_comparison.csv", index=False, encoding="utf-8-sig")
    criterion_method_pivot.to_csv(OUTPUT_DIR / "criterion_method_comparison.csv", index=False, encoding="utf-8-sig")
    inter_rater.to_csv(OUTPUT_DIR / "inter_rater_reliability.csv", index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(OUTPUT_DIR / "blind_review_analysis_results.xlsx", engine="openpyxl") as writer:
        for name, df in outputs.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)

    return outputs


# 4. ΓΡΑΦΗΜΑΤΑ

def create_charts(outputs: Dict[str, pd.DataFrame]) -> None:
    """Δημιουργεί βασικά γραφήματα για ένταξη στην πτυχιακή."""
    document_summary = outputs["document_summary"].copy()
    criterion_summary = outputs["criterion_summary"].copy()
    criterion_method_comparison = outputs["criterion_method_comparison"].copy()


    plt.figure(figsize=(9, 5))
    plt.bar(document_summary["Document"], document_summary["Mean_Total"])
    plt.ylim(0, 50)
    plt.title("Μέση Συνολική Βαθμολογία ανά Παραδοτέο")
    plt.ylabel("Mean Total Score / 50")
    plt.xlabel("Παραδοτέο")
    for i, val in enumerate(document_summary["Mean_Total"]):
        plt.text(i, val + 0.8, f"{val:.1f}", ha="center", fontweight="bold")
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / "01_total_score_by_document.png", dpi=150)
    plt.close()


    pivot = document_summary.pivot_table(
        index="Document_Type",
        columns="Method",
        values="Mean_Total",
        aggfunc="first"
    )

    pivot = pivot[["Manual", "AI-generated"]]
    ax = pivot.plot(kind="bar", figsize=(8, 5))
    ax.set_ylim(0, 50)
    ax.set_title("Σύγκριση Manual vs AI-generated ανά τύπο παραδοτέου")
    ax.set_ylabel("Mean Total Score / 50")
    ax.set_xlabel("Τύπος Παραδοτέου")
    ax.tick_params(axis="x", rotation=0)
    for container in ax.containers:
        ax.bar_label(container, fmt="%.1f", padding=3, fontweight="bold")
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / "02_manual_vs_ai_by_type.png", dpi=150)
    plt.close()

    heat = criterion_summary.pivot_table(
        index="Criterion",
        columns="Document",
        values="Mean_Score",
        aggfunc="first"
    )

    plt.figure(figsize=(9, 7))
    if sns is not None:
        sns.heatmap(heat, annot=True, fmt=".1f", vmin=1, vmax=5, cmap="YlGnBu")
    else:
        plt.imshow(heat.values, vmin=1, vmax=5)
        plt.colorbar(label="Mean Score")
        plt.xticks(range(len(heat.columns)), heat.columns, rotation=45)
        plt.yticks(range(len(heat.index)), heat.index)
        for i in range(heat.shape[0]):
            for j in range(heat.shape[1]):
                plt.text(j, i, f"{heat.values[i, j]:.1f}", ha="center", va="center")
    plt.title("Μέσος Βαθμός ανά Κριτήριο και Παραδοτέο")
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / "03_heatmap_criterion_document.png", dpi=150)
    plt.close()


    for doc_type in criterion_method_comparison["Document_Type"].unique():
        temp = criterion_method_comparison[
            criterion_method_comparison["Document_Type"] == doc_type
        ].copy()

        temp = temp.sort_values("AI_minus_Manual")

        plt.figure(figsize=(9, 6))
        plt.barh(temp["Criterion"], temp["AI_minus_Manual"])
        plt.axvline(0, linestyle="--", linewidth=1)
        plt.title(f"Διαφορά AI-generated minus Manual ανά Κριτήριο — {doc_type}")
        plt.xlabel("AI minus Manual Mean Score")
        plt.tight_layout()
        safe_doc_type = re.sub(r"[^A-Za-z0-9_-]+", "_", doc_type)
        plt.savefig(OUTPUT_DIR / f"04_ai_minus_manual_by_criterion_{safe_doc_type}.png", dpi=150)
        plt.close()


# 5. ΣΥΝΤΟΜΗ ΕΡΜΗΝΕΙΑ ΓΙΑ ΤΗΝ ΠΤΥΧΙΑΚΗ

def write_methodological_note(outputs: Dict[str, pd.DataFrame]) -> None:
    """Γράφει σύντομο txt με ερμηνεία των δεικτών agreement."""
    rel = outputs["inter_rater_reliability"].iloc[0]

    note = f"""
Ερμηνεία δεικτών συμφωνίας αξιολογητών

Το exact agreement εκφράζει το ποσοστό των περιπτώσεων όπου οι δύο αξιολογητές έδωσαν ακριβώς την ίδια βαθμολογία στο ίδιο κριτήριο και παραδοτέο. Στην παρούσα ανάλυση το exact agreement είναι {rel['Exact_Agreement_Percentage']:.2f}%.

Το within-one agreement εκφράζει το ποσοστό των περιπτώσεων όπου οι δύο αξιολογητές διαφέρουν το πολύ κατά μία μονάδα στην ordinal κλίμακα 1–5. Είναι χρήσιμο σε αξιολογήσεις Likert επειδή μικρές αποκλίσεις, όπως 4 έναντι 5, δεν έχουν την ίδια πρακτική σημασία με μεγάλες αποκλίσεις, όπως 1 έναντι 5. Στην παρούσα ανάλυση το within-one agreement είναι {rel['Within_One_Agreement_Percentage']:.2f}%.

Το weighted Cohen's kappa μετρά τη συμφωνία μεταξύ δύο αξιολογητών λαμβάνοντας υπόψη τη συμφωνία που μπορεί να προκύψει τυχαία. Επειδή η κλίμακα αξιολόγησης είναι ordinal 1–5, το weighted kappa είναι καταλληλότερο από το απλό Cohen's kappa, καθώς διαφοροποιεί το βάρος των μικρών και μεγάλων αποκλίσεων. Στην παρούσα ανάλυση, το quadratic weighted Cohen's κ είναι {rel['Weighted_Cohen_Kappa_Quadratic']:.3f}, ενώ το linear weighted Cohen's κ είναι {rel['Weighted_Cohen_Kappa_Linear']:.3f}.

Το mean absolute difference δείχνει τη μέση απόλυτη απόκλιση στη βαθμολογία των δύο reviewers. Στην παρούσα ανάλυση είναι {rel['Mean_Absolute_Difference']:.3f}, σε σύνολο {int(rel['Number_of_Rated_Items'])} αντιστοιχισμένων αξιολογήσεων.

Οι δείκτες αυτοί πρέπει να ερμηνευθούν υποστηρικτικά και όχι ως βάση ισχυρής γενίκευσης, καθώς η μελέτη περιλαμβάνει δύο senior reviewers και περιορισμένο αριθμό αξιολογούμενων παραδοτέων. Ωστόσο, ενισχύουν τη διαφάνεια της blind review διαδικασίας και επιτρέπουν πιο τεκμηριωμένη σύγκριση μεταξύ manually-written και AI-generated παραδοτέων.
""".strip()

    (OUTPUT_DIR / "methodological_note_inter_rater_reliability.txt").write_text(note, encoding="utf-8")


# 6. MAIN

if __name__ == "__main__":
    outputs = run_analysis()
    create_charts(outputs)
    write_methodological_note(outputs)

    print("\nΟλοκληρώθηκε η ανάλυση blind reviews.")
    print(f"Αποτελέσματα αποθηκεύτηκαν στον φάκελο: {OUTPUT_DIR}")

    print("\nDocument summary:")
    print(outputs["document_summary"].to_string(index=False))

    print("\nMethod comparison:")
    print(outputs["method_comparison"].to_string(index=False))

    print("\nInter-rater reliability:")
    print(outputs["inter_rater_reliability"].to_string(index=False))
