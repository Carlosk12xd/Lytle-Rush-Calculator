import io
import zipfile
from typing import Optional, List, Dict, Tuple

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Cover + Density Data Cleaner", page_icon="🌿", layout="wide")


# -----------------------------
# Professional UI styling
# -----------------------------
st.markdown(
    """
    <style>
        .stApp {
            background: linear-gradient(180deg, #f7fbf7 0%, #ffffff 38%);
        }

        .block-container {
            padding-top: 2rem;
            padding-bottom: 3rem;
            max-width: 1280px;
        }

        .hero-card {
            padding: 2rem 2.25rem;
            border-radius: 24px;
            background: linear-gradient(135deg, #163d2a 0%, #256b46 55%, #7aa95c 100%);
            color: white;
            box-shadow: 0 18px 42px rgba(22, 61, 42, 0.22);
            margin-bottom: 1.4rem;
        }

        .hero-card h1 {
            font-size: 2.35rem;
            line-height: 1.15;
            margin: 0 0 0.55rem 0;
            font-weight: 800;
            letter-spacing: -0.03em;
        }

        .hero-card p {
            font-size: 1.05rem;
            line-height: 1.6;
            max-width: 850px;
            margin: 0;
            opacity: 0.96;
        }

        .section-card {
            border: 1px solid #dfe8df;
            border-radius: 20px;
            background: rgba(255, 255, 255, 0.94);
            padding: 1.2rem 1.35rem;
            box-shadow: 0 10px 28px rgba(20, 50, 32, 0.07);
            margin: 0.85rem 0 1.1rem 0;
        }

        .mode-card {
            border-left: 6px solid #2f7d4f;
            border-radius: 18px;
            background: #f3faf4;
            padding: 1.05rem 1.2rem;
            box-shadow: 0 8px 22px rgba(20, 50, 32, 0.06);
            margin: 1rem 0 1.2rem 0;
        }

        .mode-card h3 {
            margin: 0 0 0.35rem 0;
            color: #163d2a;
            font-size: 1.18rem;
            font-weight: 750;
        }

        .mode-card p {
            margin: 0;
            color: #2d4738;
            line-height: 1.55;
        }

        div[data-testid="stFileUploader"] {
            border: 1px dashed #8eb99b;
            border-radius: 20px;
            padding: 1rem;
            background: #fbfdfb;
        }

        div[data-testid="stDownloadButton"] > button {
            width: 100%;
            min-height: 3.4rem;
            border-radius: 16px;
            font-size: 1.05rem;
            font-weight: 750;
            border: 0;
            background: linear-gradient(135deg, #1f6f43 0%, #2f8f5b 100%);
            color: white;
            box-shadow: 0 12px 24px rgba(31, 111, 67, 0.22);
        }

        div[data-testid="stDownloadButton"] > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 16px 30px rgba(31, 111, 67, 0.28);
            border: 0;
        }

        div[data-testid="stDataFrame"] {
            border-radius: 14px;
            overflow: hidden;
        }

        .small-muted {
            color: #647067;
            font-size: 0.92rem;
            margin-top: -0.3rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


PLOT_DESCRIPTIONS = {
    "BN": ("Burned", "Excluded"),
    "BS": ("Burned", "Present"),
    "UN": ("Unburned", "Excluded"),
    "US": ("Unburned", "Present"),
}

PLOT_ORDER = ["BN", "BS", "UN", "US"]


# -----------------------------
# Shared helper functions
# -----------------------------

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize messy Excel/CSV headers into predictable uppercase text."""
    out = df.copy()
    out.columns = (
        out.columns.astype(str)
        .str.replace("\ufeff", "", regex=False)
        .str.strip()
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
    )
    return out


def find_first_matching_column(columns: List[str], required_terms: List[str]) -> Optional[str]:
    for col in columns:
        if all(term in col for term in required_terms):
            return col
    return None


def choose_year_value(df: pd.DataFrame) -> str:
    year_col = find_first_matching_column(df.columns.tolist(), ["YEAR"])
    if year_col:
        non_null = pd.to_numeric(df[year_col], errors="coerce").dropna()
        if not non_null.empty:
            return int(non_null.iloc[0])
        first_text = df[year_col].dropna().astype(str).str.strip()
        if not first_text.empty:
            return first_text.iloc[0]

    date_col = find_first_matching_column(df.columns.tolist(), ["DATE"])
    if date_col:
        first_text = df[date_col].dropna().astype(str).str.strip()
        if not first_text.empty:
            return first_text.iloc[0]

    return ""


def clean_text_series(series: pd.Series, fill_value: str = "0") -> pd.Series:
    return (
        series
        .fillna(fill_value)
        .astype(str)
        .str.strip()
        .str.upper()
        .replace("", fill_value)
        .replace("NAN", fill_value)
    )


@st.cache_data(show_spinner=False)
def read_uploaded_file(file_bytes: bytes, filename: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    lower = filename.lower()
    buffer = io.BytesIO(file_bytes)
    if lower.endswith(".csv"):
        return pd.read_csv(buffer)
    if lower.endswith(".xlsx") or lower.endswith(".xls"):
        return pd.read_excel(buffer, sheet_name=sheet_name)
    raise ValueError("Unsupported file type. Upload a CSV or Excel file.")


@st.cache_data(show_spinner=False)
def get_excel_sheet_names(file_bytes: bytes) -> List[str]:
    buffer = io.BytesIO(file_bytes)
    xl = pd.ExcelFile(buffer)
    return xl.sheet_names


# -----------------------------
# Cover data processing
# This implements the same cover logic as the uploaded Python program:
# 1. Clean BLOCK and PLOT.
# 2. Use one hit per row, with foliar priority:
#    1ST FOLIAR > 2ND FOLIAR > 3RD FOLIAR > BASAL > empty label.
# 3. Group by BLOCK + PLOT + HIT.
# 4. Force every BLOCK + PLOT combo to appear.
# 5. Output clean Cover Proportion and Cover Percent sheets.
# -----------------------------

def process_cover_dataframe(
    df: pd.DataFrame,
    empty_hit_label: str = "NO_HIT",
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, Optional[str]]]:
    df = normalize_columns(df)

    # Drop blank rows, and drop empty columns only after normalizing headers.
    # If a foliar layer is completely empty, dropping it is okay because the code handles missing layers.
    df = df.dropna(axis=1, how="all").dropna(how="all").copy()

    columns = df.columns.tolist()

    block_col = find_first_matching_column(columns, ["BLOCK"])
    plot_col = find_first_matching_column(columns, ["PLOT"])
    basal_col = find_first_matching_column(columns, ["BASAL"])
    first_foliar_col = find_first_matching_column(columns, ["1ST", "FOLIAR"])
    second_foliar_col = find_first_matching_column(columns, ["2ND", "FOLIAR"])
    third_foliar_col = find_first_matching_column(columns, ["3RD", "FOLIAR"])
    transect_col = find_first_matching_column(columns, ["TRANSECT"])

    if block_col is None or plot_col is None:
        raise ValueError(
            "Could not find the required BLOCK and PLOT columns. "
            f"Columns found: {columns}"
        )
    if basal_col is None:
        raise ValueError(
            "Could not find a BASAL column. "
            f"Columns found: {columns}"
        )

    matched = {
        "DATA_TYPE": "Cover Data",
        "BLOCK": block_col,
        "PLOT": plot_col,
        "TRANSECT": transect_col,
        "BASAL": basal_col,
        "1ST FOLIAR": first_foliar_col,
        "2ND FOLIAR": second_foliar_col,
        "3RD FOLIAR": third_foliar_col,
        "EMPTY_HIT_LABEL": empty_hit_label,
    }

    work = df.copy()
    work[block_col] = pd.to_numeric(work[block_col], errors="coerce").astype("Int64")
    work[plot_col] = work[plot_col].astype(str).str.strip().str.upper()

    # Keep real observation rows only.
    work = work[work[block_col].notna() & work[plot_col].isin(PLOT_ORDER)].copy()

    cover_columns = [c for c in [basal_col, third_foliar_col, second_foliar_col, first_foliar_col] if c is not None]
    for col in cover_columns:
        work[col] = (
            work[col]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
            .replace("NAN", "")
        )

    def choose_hit(row: pd.Series) -> str:
        if first_foliar_col is not None and row[first_foliar_col] != "":
            return row[first_foliar_col]
        elif second_foliar_col is not None and row[second_foliar_col] != "":
            return row[second_foliar_col]
        elif third_foliar_col is not None and row[third_foliar_col] != "":
            return row[third_foliar_col]
        elif row[basal_col] != "":
            return row[basal_col]
        else:
            return empty_hit_label

    work["HIT"] = work.apply(choose_hit, axis=1)

    counts_df = (
        work.groupby([block_col, plot_col, "HIT"])
        .size()
        .reset_index(name="COUNT")
    )

    cover_df = counts_df.pivot_table(
        index=[block_col, plot_col],
        columns="HIT",
        values="COUNT",
        fill_value=0,
        aggfunc="sum",
    ).reset_index()

    cover_df = cover_df.rename(columns={block_col: "BLOCK", plot_col: "PLOT"})

    # Make sure every BLOCK + PLOT combination exists, following the original code pattern.
    all_blocks = sorted(work[block_col].dropna().unique())
    full_index = pd.MultiIndex.from_product([all_blocks, PLOT_ORDER], names=["BLOCK", "PLOT"])

    cover_df = (
        cover_df.set_index(["BLOCK", "PLOT"])
        .reindex(full_index, fill_value=0)
        .reset_index()
    )

    cover_df["FIRE"] = cover_df["PLOT"].map(lambda x: PLOT_DESCRIPTIONS.get(x, ("", ""))[0])
    cover_df["RODENTS"] = cover_df["PLOT"].map(lambda x: PLOT_DESCRIPTIONS.get(x, ("", ""))[1])
    cover_df["YEAR"] = choose_year_value(work)

    fixed_cols = ["YEAR", "BLOCK", "PLOT", "FIRE", "RODENTS"]
    species_cols = sorted([col for col in cover_df.columns if col not in fixed_cols])
    cover_df = cover_df[fixed_cols + species_cols]

    # Match the original program output: keep TOTAL_HITS and ROW_SUM in the exported sheets.
    cover_df["TOTAL_HITS"] = cover_df[species_cols].sum(axis=1)

    # Normalize by the total number of chosen hits in each BLOCK + PLOT.
    # This keeps each row at 1.0 / 100% for real plot-treatment combinations.
    proportion_df = cover_df.copy()
    for col in species_cols:
        proportion_df[col] = proportion_df[col] / proportion_df["TOTAL_HITS"]
    proportion_df[species_cols] = proportion_df[species_cols].fillna(0)
    proportion_df["ROW_SUM"] = proportion_df[species_cols].sum(axis=1)

    percent_df = proportion_df.copy()
    for col in species_cols:
        percent_df[col] = percent_df[col] * 100
    percent_df["ROW_SUM"] = percent_df[species_cols].sum(axis=1)

    point_counts_df = (
        work.groupby([block_col, plot_col])
        .size()
        .reset_index(name="TOTAL_POINTS")
        .rename(columns={block_col: "BLOCK", plot_col: "PLOT"})
    )

    return proportion_df, percent_df, point_counts_df, matched


# -----------------------------
# Density data processing
# This implements the uploaded Rush Valley density calculator structure:
# 1. Read DATE, BLOCK, PLOT, SPECIES, and columns 1 through 12.
# 2. Sum columns 1-12 for each input species row.
# 3. Add species totals by BLOCK + PLOT.
# 4. Output one wide sheet named "Species Totals Wide" with:
#    DATE, BLOCK, PLOT, FIRE, RODENTS, species columns...
# 5. Asterisks are ignored and never multiplied. Example: 2* -> 2, * -> 0.
# -----------------------------

def parse_density_cell(value) -> float:
    """Parse density count cells. Asterisks are ignored, not multiplied."""
    if pd.isna(value):
        return 0.0

    text = str(value).strip()
    if text == "" or text.lower() == "nan":
        return 0.0

    # Ignore stars completely. No *5 multiplier.
    # Examples: "2*" -> 2, "5*" -> 5, "*" -> 0.
    text = text.replace("*", "").replace(",", "").strip()

    if text == "":
        return 0.0

    try:
        return float(text)
    except ValueError:
        return 0.0


def find_density_count_columns(columns: List[str]) -> List[str]:
    count_cols = []
    for col in columns:
        clean_col = str(col).strip()
        if clean_col.isdigit() and 1 <= int(clean_col) <= 12:
            count_cols.append(col)
    return sorted(count_cols, key=lambda x: int(str(x).strip()))


def choose_date_value_for_density(df: pd.DataFrame) -> str:
    """Match the old density output column name DATE. Use DATE first, then YEAR."""
    date_col = find_first_matching_column(df.columns.tolist(), ["DATE"])
    if date_col:
        first_text = df[date_col].dropna().astype(str).str.strip()
        first_text = first_text[first_text.ne("") & first_text.str.lower().ne("nan")]
        if not first_text.empty:
            value = first_text.iloc[0]
            numeric_value = pd.to_numeric(pd.Series([value]), errors="coerce").dropna()
            if not numeric_value.empty and float(numeric_value.iloc[0]).is_integer():
                return int(numeric_value.iloc[0])
            return value

    year_col = find_first_matching_column(df.columns.tolist(), ["YEAR"])
    if year_col:
        first_text = df[year_col].dropna().astype(str).str.strip()
        first_text = first_text[first_text.ne("") & first_text.str.lower().ne("nan")]
        if not first_text.empty:
            value = first_text.iloc[0]
            numeric_value = pd.to_numeric(pd.Series([value]), errors="coerce").dropna()
            if not numeric_value.empty and float(numeric_value.iloc[0]).is_integer():
                return int(numeric_value.iloc[0])
            return value

    return ""


def normalize_block_for_density(value):
    """Keep numeric block labels as integers when possible, otherwise keep cleaned text."""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text == "" or text.lower() == "nan":
        return ""
    numeric_value = pd.to_numeric(pd.Series([text]), errors="coerce").dropna()
    if not numeric_value.empty and float(numeric_value.iloc[0]).is_integer():
        return int(numeric_value.iloc[0])
    return text.upper()


def density_species_columns_to_int_if_possible(df: pd.DataFrame, species_cols: List[str]) -> pd.DataFrame:
    out = df.copy()
    for col in species_cols:
        numeric = pd.to_numeric(out[col], errors="coerce")
        if numeric.notna().all() and (numeric % 1 == 0).all():
            out[col] = numeric.astype(int)
        else:
            out[col] = numeric
    return out


def process_density_dataframe(
    df: pd.DataFrame,
    source_name: str,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, Optional[str]]]:
    df = normalize_columns(df)
    df = df.dropna(axis=1, how="all").dropna(how="all").copy()

    columns = df.columns.tolist()

    block_col = find_first_matching_column(columns, ["BLOCK"])
    plot_col = find_first_matching_column(columns, ["PLOT"])
    transect_col = find_first_matching_column(columns, ["TRANSECT"])
    species_col = find_first_matching_column(columns, ["SPECIES"])
    if species_col is None:
        species_col = find_first_matching_column(columns, ["PLANT"])

    count_cols = find_density_count_columns(columns)

    if block_col is None or plot_col is None:
        raise ValueError(
            "Could not find the required BLOCK and PLOT columns for density data. "
            f"Columns found: {columns}"
        )
    if species_col is None:
        raise ValueError(
            "Could not find a SPECIES or PLANT SPECIES column for density data. "
            f"Columns found: {columns}"
        )
    if len(count_cols) == 0:
        raise ValueError(
            "Could not find count columns named 1 through 12. "
            f"Columns found: {columns}"
        )

    matched = {
        "DATA_TYPE": "Density Data",
        "SOURCE_FILE": source_name,
        "BLOCK": block_col,
        "PLOT": plot_col,
        "TRANSECT": transect_col,
        "SPECIES": species_col,
        "COUNT_COLUMNS": ", ".join(count_cols),
        "ASTERISK_RULE": "Asterisks are ignored. 2* is treated as 2. No *5 multiplier is applied.",
    }

    work = df.copy()
    work[block_col] = work[block_col].apply(normalize_block_for_density)
    work[plot_col] = (
        work[plot_col]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .replace("NAN", "")
    )

    # Match the original density program: strip species labels, but do not uppercase them.
    # This keeps labels like "Unknown" exactly like the provided example workbook.
    work[species_col] = (
        work[species_col]
        .fillna("")
        .astype(str)
        .str.strip()
        .replace("nan", "")
        .replace("NAN", "")
    )

    # Keep rows that have a block, plot, and species label.
    work = work[
        work[block_col].ne("")
        & work[plot_col].isin(PLOT_ORDER)
        & work[species_col].ne("")
    ].copy()

    for col in count_cols:
        work[col] = work[col].apply(parse_density_cell)

    work["ROW_TOTAL"] = work[count_cols].sum(axis=1)

    # Sum species totals by BLOCK + PLOT + SPECIES, like the original dictionary approach.
    totals_df = (
        work.groupby([block_col, plot_col, species_col], dropna=False)["ROW_TOTAL"]
        .sum()
        .reset_index(name="TOTAL_COUNT")
    )

    all_species = sorted(totals_df[species_col].dropna().astype(str).unique().tolist())

    # Match the original output order: BN blocks 1-5, then BS blocks 1-5, then UN, then US.
    observed_blocks = [b for b in work[block_col].dropna().unique().tolist() if b != ""]
    numeric_blocks = sorted([b for b in observed_blocks if isinstance(b, int)])
    other_blocks = sorted([b for b in observed_blocks if not isinstance(b, int)], key=lambda x: str(x))

    default_blocks = [1, 2, 3, 4, 5]
    blocks = default_blocks + [b for b in numeric_blocks if b not in default_blocks] + other_blocks

    total_lookup = {}
    for _, row in totals_df.iterrows():
        key = (row[block_col], row[plot_col], row[species_col])
        total_lookup[key] = row["TOTAL_COUNT"]

    date_value = choose_date_value_for_density(df)

    rows_for_output = []
    for plot_value in PLOT_ORDER:
        for block_value in blocks:
            row_dict = {
                "DATE": date_value,
                "BLOCK": block_value,
                "PLOT": plot_value,
                "FIRE": PLOT_DESCRIPTIONS.get(plot_value, ("", ""))[0],
                "RODENTS": PLOT_DESCRIPTIONS.get(plot_value, ("", ""))[1],
            }
            for species in all_species:
                row_dict[species] = total_lookup.get((block_value, plot_value, species), 0)
            rows_for_output.append(row_dict)

    fixed_cols = ["DATE", "BLOCK", "PLOT", "FIRE", "RODENTS"]
    density_totals_wide = pd.DataFrame(rows_for_output)
    if all_species:
        density_totals_wide = density_totals_wide[fixed_cols + all_species]
        density_totals_wide = density_species_columns_to_int_if_possible(density_totals_wide, all_species)
    else:
        density_totals_wide = density_totals_wide[fixed_cols]

    # This is only used as a preview/check in the app; it is not exported, so the output workbook
    # stays in the same one-sheet format as the provided example.
    input_row_totals_df = work[[block_col, plot_col, species_col, "ROW_TOTAL"]].rename(
        columns={
            block_col: "BLOCK",
            plot_col: "PLOT",
            species_col: "SPECIES",
            "ROW_TOTAL": "ROW_TOTAL",
        }
    )

    warnings = []
    if len(count_cols) != 12:
        warnings.append(
            f"Found {len(count_cols)} count columns instead of 12. The app used: {', '.join(count_cols)}."
        )
    if work.empty:
        warnings.append("No usable density rows were found after cleaning the uploaded file.")

    warnings_df = pd.DataFrame({"WARNING": warnings}) if warnings else pd.DataFrame({"WARNING": ["No structural warnings detected."]})

    matched["warnings_df"] = warnings_df
    return density_totals_wide, input_row_totals_df, warnings_df, matched


# -----------------------------
# Workbook writers
# -----------------------------

def matched_columns_dataframe(matched_columns: Dict[str, Optional[str]]) -> pd.DataFrame:
    rows = []
    for key, value in matched_columns.items():
        if key == "warnings_df":
            continue
        rows.append({"EXPECTED_FIELD": key, "MATCHED_SOURCE_COLUMN": value or ""})
    return pd.DataFrame(rows)


def workbook_bytes_cover(
    proportion_df: pd.DataFrame,
    percent_df: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        proportion_df.to_excel(writer, sheet_name="Cover Proportion", index=False)
        percent_df.to_excel(writer, sheet_name="Cover Percent", index=False)
    output.seek(0)
    return output.getvalue()


def workbook_bytes_density(
    density_totals_wide_df: pd.DataFrame,
) -> bytes:
    """Write density output in the same one-sheet format as the provided example workbook."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        density_totals_wide_df.to_excel(writer, sheet_name="Species Totals Wide", index=False)
    output.seek(0)
    return output.getvalue()


# -----------------------------
# Streamlit UI
# -----------------------------

st.markdown(
    """
    <div class="hero-card">
        <h1>🌿 Lytle + Rush Data Cleaner</h1>
        <p>
            Upload field data in CSV or Excel format, choose the processing type, preview the cleaned output,
            and download a ready-to-use Excel workbook.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

left_intro, right_intro = st.columns([1.25, 1])

with left_intro:
    processing_mode = st.radio(
        "Choose the type of data to clean",
        options=["Cover Data", "Density Data"],
        horizontal=True,
    )

with right_intro:
    st.markdown(
        """
        <div class="section-card">
            <strong>Output formats</strong><br>
            <span class="small-muted">
                Cover data creates proportion and percent sheets. Density data creates a species totals wide sheet.
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )

empty_hit_label = "NO_HIT"

if processing_mode == "Cover Data":
    st.markdown(
        """
        <div class="mode-card">
            <h3>Cover Data</h3>
            <p>
                Use this option for point-intercept cover files with columns such as
                <strong>BASAL</strong>, <strong>1ST FOLIAR</strong>, <strong>2ND FOLIAR</strong>, and
                <strong>3RD FOLIAR</strong>. The app selects one cover hit per row using foliar layers first,
                then basal, groups the results by <strong>BLOCK + PLOT</strong>, and exports clean cover
                proportion and percent spreadsheets.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    empty_hit_label = st.selectbox(
        "Label for completely empty cover rows",
        options=["NO_HIT", "0"],
        index=0,
        help="NO_HIT matches the uploaded example workbook. Choose 0 if you want blank cover rows labeled as 0 instead.",
    )
else:
    st.markdown(
        """
        <div class="mode-card">
            <h3>Density Data</h3>
            <p>
                Use this option for density files with <strong>BLOCK</strong>, <strong>PLOT</strong>,
                <strong>SPECIES</strong>, and count columns <strong>1 through 12</strong>. The app sums the
                count columns for each species, combines totals by <strong>BLOCK + PLOT</strong>, and exports
                a wide species totals spreadsheet. Blank count cells are treated as 0, and asterisks are ignored.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

uploaded_files = st.file_uploader(
    "Upload CSV or Excel files",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=True,
    help="You can upload one file or several files at once.",
)

if uploaded_files:
    all_outputs = []

    st.markdown("### Processing results")

    for idx, uploaded_file in enumerate(uploaded_files, start=1):
        file_bytes = uploaded_file.getvalue()
        st.divider()
        st.markdown(f"#### File {idx}: `{uploaded_file.name}`")

        selected_sheet = None
        if uploaded_file.name.lower().endswith((".xlsx", ".xls")):
            try:
                sheet_names = get_excel_sheet_names(file_bytes)
                selected_sheet = st.selectbox(
                    f"Choose a sheet for {uploaded_file.name}",
                    options=sheet_names,
                    index=0,
                    key=f"sheet_{idx}",
                )
            except Exception as exc:
                st.error(f"Could not read workbook sheets: {exc}")
                continue

        try:
            raw_df = read_uploaded_file(file_bytes, uploaded_file.name, selected_sheet)

            with st.expander("Raw file preview", expanded=False):
                st.dataframe(raw_df.head(10), use_container_width=True)

            if processing_mode == "Cover Data":
                proportion_df, percent_df, point_counts_df, matched = process_cover_dataframe(
                    raw_df,
                    empty_hit_label=empty_hit_label,
                )
                xlsx_bytes = workbook_bytes_cover(proportion_df, percent_df)
                preview_title = "Cover percent preview"
                check_title = "Point counts check"
                preview_df = percent_df
                check_df = point_counts_df
                suffix = "cover_cleaned"

                bad_counts = point_counts_df[point_counts_df["TOTAL_POINTS"] != 36]
                if not bad_counts.empty:
                    st.warning(
                        "Some BLOCK + PLOT combinations do not have exactly 36 rows in the uploaded file. "
                        "The cover percentages were normalized by the total chosen hits available for each BLOCK + PLOT."
                    )

            else:
                totals_df, check_df, warnings_df, matched = process_density_dataframe(raw_df, uploaded_file.name)
                xlsx_bytes = workbook_bytes_density(totals_df)
                preview_title = "Species totals wide preview"
                check_title = "Input row totals check"
                preview_df = totals_df
                suffix = "density_cleaned"

                if not warnings_df.empty:
                    for msg in warnings_df["WARNING"].tolist():
                        if msg != "No structural warnings detected.":
                            st.warning(msg)
                        else:
                            st.success(msg)

            preview_col, check_col = st.columns([1.2, 1])
            with preview_col:
                st.markdown(f"**{preview_title}**")
                st.dataframe(preview_df.head(10), use_container_width=True)
            with check_col:
                st.markdown(f"**{check_title}**")
                st.dataframe(check_df, use_container_width=True)

            output_name = uploaded_file.name.rsplit(".", 1)[0] + f"_{suffix}.xlsx"
            all_outputs.append((output_name, xlsx_bytes))

            st.download_button(
                label=f"⬇️ Download cleaned Excel workbook: {output_name}",
                data=xlsx_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{idx}",
            )
        except Exception as exc:
            st.error(f"Could not process {uploaded_file.name}: {exc}")

    if all_outputs:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for output_name, xlsx_bytes in all_outputs:
                zf.writestr(output_name, xlsx_bytes)
        zip_buffer.seek(0)

        st.divider()
        zip_name = "cleaned_cover_workbooks.zip" if processing_mode == "Cover Data" else "cleaned_density_workbooks.zip"
        st.download_button(
            label="⬇️ Download all cleaned workbooks as ZIP",
            data=zip_buffer.getvalue(),
            file_name=zip_name,
            mime="application/zip",
        )
else:
    st.markdown(
        """
        <div class="section-card">
            <strong>Ready when you are.</strong><br>
            <span class="small-muted">Upload at least one CSV or Excel file to generate a cleaned workbook.</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
