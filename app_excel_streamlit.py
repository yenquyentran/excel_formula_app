import json
import re
from pathlib import Path
from typing import Dict

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

EXCEL_FILE = Path("Excel Formula.xlsx")
SHEET_NAME = "Sheet1"
SUPPORTED_LANGS = ["English", "German", "French"]

LOGICAL_CONSTANTS = {
    "TRUE": {"English": "TRUE", "German": "WAHR", "French": "VRAI"},
    "FALSE": {"English": "FALSE", "German": "FALSCH", "French": "FAUX"},
}


@st.cache_data
def load_mapping() -> Dict[str, Dict[str, str]]:
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Không tìm thấy file: {EXCEL_FILE}")

    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, engine="openpyxl")
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
    logic_map = {v[src].upper(): k for k, v in LOGICAL_CONSTANTS.items()}
    pattern = re.compile(
        r"\b(" + "|".join(map(re.escape, logic_map.keys())) + r")\b",
        flags=re.IGNORECASE,
    )

    out = []
    for is_str, part in split_string(formula):
        if is_str:
            out.append(part)
        else:
            out.append(
                pattern.sub(
                    lambda m: LOGICAL_CONSTANTS[logic_map[m.group(1).upper()]][tgt],
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


def compact_formula(formula, sep):
    formula = re.sub(r"\s*[,;]\s*", sep + " ", formula)
    formula = re.sub(r"\s+\)", ")", formula)
    formula = re.sub(r"\s+\(", "(", formula)
    formula = re.sub(r"[ \t]+", " ", formula)
    return formula.strip()


def pretty_formula(formula):
    out = []
    depth = 0
    in_str = False
    i = 0

    while i < len(formula):
        ch = formula[i]

        if ch == '"':
            out.append(ch)
            if in_str and i + 1 < len(formula) and formula[i + 1] == '"':
                out.append('"')
                i += 2
                continue
            in_str = not in_str
            i += 1
            continue

        if in_str:
            out.append(ch)
            i += 1
            continue

        if ch == "(":
            depth += 1
            out.append("(\n" + "    " * depth)
        elif ch == ")":
            depth = max(0, depth - 1)
            out.append("\n" + "    " * depth + ")")
        elif ch in {",", ";"}:
            out.append(ch + "\n" + "    " * depth)
        else:
            out.append(ch)

        i += 1

    formatted = "".join(out)
    formatted = re.sub(r"\n[ \t]+\n", "\n", formatted)
    formatted = re.sub(r"\n{3,}", "\n\n", formatted)
    return formatted.strip()


def translate(formula, src, tgt, sep, pretty_mode, mapping, reverse):
    formula = replace_functions(formula, src, tgt, mapping, reverse)
    formula = replace_logic(formula, src, tgt)
    formula = replace_separator(formula, sep)
    formula = compact_formula(formula, sep)
    return pretty_formula(formula) if pretty_mode else formula


def copy_button(text: str):
    payload = json.dumps(text)

    html = """
    <div style="margin-top: 8px;">
        <button id="copy-btn"
            style="
                background:#ff4b4b;
                color:white;
                border:none;
                padding:10px 14px;
                border-radius:8px;
                cursor:pointer;
                font-size:14px;
                font-weight:600;
            ">
            Copy
        </button>

        <span id="copy-msg"
            style="
                display:none;
                margin-left:10px;
                color:#16a34a;
                font-weight:600;
                font-size:14px;
            ">
            Formula copied!
        </span>
    </div>

    <script>
    const textToCopy = PAYLOAD_TEXT;
    const btn = document.getElementById("copy-btn");
    const msg = document.getElementById("copy-msg");

    btn.addEventListener("click", async function () {
        try {
            await navigator.clipboard.writeText(textToCopy);
            msg.style.display = "inline";
            setTimeout(function () {
                msg.style.display = "none";
            }, 1500);
        } catch (err) {
            msg.style.display = "inline";
            msg.style.color = "#dc2626";
            msg.innerText = "Copy failed";
            setTimeout(function () {
                msg.style.display = "none";
                msg.style.color = "#16a34a";
                msg.innerText = "Formula copied!";
            }, 1500);
        }
    });
    </script>
    """.replace("PAYLOAD_TEXT", payload)

    components.html(html, height=70)


st.set_page_config(page_title="Excel Formula Translator", page_icon="🔁", layout="wide")

try:
    mapping = load_mapping()
    reverse = build_reverse_lookup(mapping)
except Exception as e:
    st.error(f"Lỗi đọc file Excel: {e}")
    st.stop()

st.title("🔁 Excel Formula Translator")

with st.sidebar:
    src = st.selectbox("Source language", SUPPORTED_LANGS)
    tgt = st.selectbox("Target language", SUPPORTED_LANGS, index=2)
    sep = st.radio("Separator", [",", ";"], horizontal=True, index=1)
    pretty_mode = st.toggle("Pretty format", value=True)

    ##st.divider()
    ##st.metric("Functions loaded", len(mapping))

col1, col2 = st.columns(2)

with col1:
    formula = st.text_area(
        "Input formula",
        height=320,
        placeholder='Ví dụ: =IF(SUM(A1,B1)>10,VLOOKUP(C1,Sheet2!A:B,2,FALSE),"No")',
    )
    run = st.button("Translate", type="primary", use_container_width=True)

with col2:
    if run:
        if not formula.strip():
            st.warning("Nhập công thức trước.")
        else:
            result = translate(formula, src, tgt, sep, pretty_mode, mapping, reverse)
            st.text_area("Translated formula", result, height=320)
            copy_button(result)
    else:
        st.info("Nhập công thức rồi bấm Translate.")
