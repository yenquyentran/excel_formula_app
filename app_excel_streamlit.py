import re
from pathlib import Path
from typing import Dict

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


# ================= CONFIG =================
EXCEL_FILE = Path("Excel Formula.xlsx")
SHEET_NAME = "Sheet1"

SUPPORTED_LANGS = ["English", "German", "French"]

LOGICAL_CONSTANTS = {
    "TRUE": {"English": "TRUE", "German": "WAHR", "French": "VRAI"},
    "FALSE": {"English": "FALSE", "German": "FALSCH", "French": "FAUX"},
}


# ================= LOAD MAPPING =================
@st.cache_data
def load_mapping() -> Dict[str, Dict[str, str]]:
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Không tìm thấy file: {EXCEL_FILE}")

    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

    df = df[["ENG", "GER", "FRE"]].copy()

    for col in ["ENG", "GER", "FRE"]:
        df[col] = df[col].astype(str).str.strip()

    df = df[
        (df["ENG"] != "")
        & (df["GER"] != "")
        & (df["FRE"] != "")
        & (df["ENG"].str.lower() != "nan")
        & (df["GER"].str.lower() != "nan")
        & (df["FRE"].str.lower() != "nan")
    ].drop_duplicates(subset=["ENG"])

    mapping = {}

    for _, row in df.iterrows():
        en = row["ENG"].upper()
        mapping[en] = {
            "English": en,
            "German": row["GER"].upper(),
            "French": row["FRE"].upper(),
        }

    return mapping


def build_reverse_lookup(mapping):
    reverse = {"English": {}, "German": {}, "French": {}}

    for en, names in mapping.items():
        for lang in SUPPORTED_LANGS:
            reverse[lang][names[lang].upper()] = en

    return reverse


# ================= CORE LOGIC =================
def split_string(formula):
    pattern = re.compile(r'"(?:[^"]|"")*"')
    last = 0
    for m in pattern.finditer(formula):
        if m.start() > last:
            yield False, formula[last:m.start()]
        yield True, formula[m.start():m.end()]
        last = m.end()
    if last < len(formula):
        yield False, formula[last:]


def replace_functions(formula, src, tgt, mapping, reverse):
    funcs = sorted(reverse[src].keys(), key=len, reverse=True)

    pattern = re.compile(
        r"\b(" + "|".join(map(re.escape, funcs)) + r")(?=\s*\()",
        flags=re.IGNORECASE,
    )

    out = []
    for is_str, part in split_string(formula):
        if is_str:
            out.append(part)
        else:
            out.append(
                pattern.sub(
                    lambda m: mapping[reverse[src][m.group(1).upper()]][tgt],
                    part,
                )
            )
    return "".join(out)


def replace_logic(formula, src, tgt):
    mapping = {v[src]: k for k, v in LOGICAL_CONSTANTS.items()}

    pattern = re.compile(
        r"\b(" + "|".join(map(re.escape, mapping.keys())) + r")\b",
        flags=re.IGNORECASE,
    )

    out = []
    for is_str, part in split_string(formula):
        if is_str:
            out.append(part)
        else:
            out.append(
                pattern.sub(
                    lambda m: LOGICAL_CONSTANTS[mapping[m.group(1).upper()]][tgt],
                    part,
                )
            )
    return "".join(out)


def replace_separator(formula, sep):
    res = []
    in_str = False
    depth = 0

    i = 0
    while i < len(formula):
        ch = formula[i]

        if ch == '"':
            res.append(ch)
            if in_str and i + 1 < len(formula) and formula[i + 1] == '"':
                res.append('"')
                i += 2
                continue
            in_str = not in_str
            i += 1
            continue

        if not in_str:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth = max(0, depth - 1)
            elif ch in {",", ";"} and depth > 0:
                res.append(sep)
                i += 1
                continue

        res.append(ch)
        i += 1

    return "".join(res)


def format_formula(formula, sep):
    formula = re.sub(r"\s*[,;]\s*", sep + " ", formula)
    formula = re.sub(r"\s+\)", ")", formula)
    formula = re.sub(r"\s+\(", "(", formula)
    return formula.strip()


def pretty(formula):
    out = []
    depth = 0
    in_str = False

    for ch in formula:
        if ch == '"':
            in_str = not in_str
            out.append(ch)
            continue

        if in_str:
            out.append(ch)
            continue

        if ch == "(":
            depth += 1
            out.append("(\n" + "    " * depth)
        elif ch == ")":
            depth -= 1
            out.append("\n" + "    " * depth + ")")
        elif ch in {",", ";"}:
            out.append(ch + "\n" + "    " * depth)
        else:
            out.append(ch)

    return "".join(out)


def translate(formula, src, tgt, sep, pretty_mode, mapping, reverse):
    formula = replace_functions(formula, src, tgt, mapping, reverse)
    formula = replace_logic(formula, src, tgt)
    formula = replace_separator(formula, sep)
    formula = format_formula(formula, sep)
    return pretty(formula) if pretty_mode else formula


# ================= COPY BUTTON =================
def copy_button(text):
    safe = text.replace("\\", "\\\\").replace("`", "\\`")

    components.html(
        f"""
        <button onclick="navigator.clipboard.writeText(`{safe}`)"
        style="background:#ff4b4b;color:white;border:none;padding:8px 12px;border-radius:6px;cursor:pointer;">
        Copy
        </button>
        """,
        height=40,
    )


# ================= APP =================
st.set_page_config(page_title="Excel Formula Translator", layout="wide")

try:
    mapping = load_mapping()
    reverse = build_reverse_lookup(mapping)
except Exception as e:
    st.error(f"Lỗi đọc file Excel: {e}")
    st.stop()

st.title("🔁 Excel Formula Translator")

with st.sidebar:
    src = st.selectbox("Source", SUPPORTED_LANGS)
    tgt = st.selectbox("Target", SUPPORTED_LANGS, index=2)
    sep = st.radio("Separator", [",", ";"], horizontal=True, index=1)
    pretty_mode = st.toggle("Pretty format", True)

    ##st.metric("Functions loaded", len(mapping))

col1, col2 = st.columns(2)

with col1:
    formula = st.text_area("Input", height=250)
    run = st.button("Translate", type="primary")

with col2:
    if run and formula:
        result = translate(formula, src, tgt, sep, pretty_mode, mapping, reverse)
        st.code(result)
        st.text_area("Copy", result, height=250)
        copy_button(result)
    elif run:
        st.warning("Nhập công thức trước")
