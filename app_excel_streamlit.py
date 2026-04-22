import json
import re
from pathlib import Path
from typing import Dict

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import base64

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
    if not funcs:
        return formula

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
    if not logic_map:
        return formula

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


def translate_single_function_name(formula, src, tgt, mapping, reverse):
    token = formula.strip().upper()
    if token in reverse[src]:
        en_name = reverse[src][token]
        return mapping[en_name][tgt]
    return None


def translate(formula, src, tgt, sep, pretty_mode, mapping, reverse):
    raw = formula.strip()

    # 🔥 Detect dấu '='
    has_equal = raw.startswith("=")
    if has_equal:
        raw = raw[1:]  # bỏ '=' đi

    # 🔹 case chỉ nhập tên hàm (IF, SUM,...)
    single_func = translate_single_function_name(raw, src, tgt, mapping, reverse)
    if single_func is not None:
        return "=" + single_func if has_equal else single_func

    # 🔹 xử lý bình thường
    formula = replace_functions(raw, src, tgt, mapping, reverse)
    formula = replace_logic(formula, src, tgt)
    formula = replace_separator(formula, sep)
    formula = compact_formula(formula, sep)

    result = pretty_formula(formula) if pretty_mode else formula

    # 🔥 add lại '=' nếu có
    return "=" + result if has_equal else result


def copy_button(text: str):
    payload = json.dumps(text)

    html = """
    <div style="
        width:100%;
        display:flex;
        justify-content:center;
        align-items:center;
        gap:12px;
        margin-top:2px;
    ">
        <button id="copy-btn"
            style="
                min-width:140px;
                height:40px;
                padding:0 24px;
                border-radius:999px;
                border:none;
                background:#3d9e9d;
                color:white;
                cursor:pointer;
                font-size:16px;
                font-weight:500;
                display:inline-flex;
                align-items:center;
                justify-content:center;
                box-shadow:0 6px 16px rgba(61, 158, 157, 0.35);
                transition:all 0.2s ease;
            "
            onmouseover="this.style.background='#348b8a'; this.style.boxShadow='0 8px 20px rgba(61, 158, 157, 0.45)'"
            onmouseout="this.style.background='#3d9e9d'; this.style.boxShadow='0 6px 16px rgba(61, 158, 157, 0.35)'"
            >
            Copy
        </button>

        <span id="copy-msg"
            style="
                display:none;
                color:#16a34a;
                font-weight:600;
                font-size:15px;
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
            msg.innerText = "Copy failed";
            msg.style.color = "#dc2626";
            msg.style.display = "inline";

            setTimeout(function () {
                msg.style.display = "none";
                msg.innerText = "Formula copied!";
                msg.style.color = "#16a34a";
            }, 1500);
        }
    });
    </script>
    """.replace("PAYLOAD_TEXT", payload)

    components.html(html, height=60)


st.set_page_config(page_title="Excel Formula Translator", page_icon="logo_xanh.png", layout="wide")

st.markdown(
    """
    <style>
    /* Label */
    div[data-testid="stTextArea"] label {
        color: #111827 !important;
    }

    div[data-testid="stTextArea"] label p {
        font-size: 0.875rem !important;
    }

    /* Text màu đen cho cả 2 box */
    div[data-testid="stTextArea"] textarea,
    div[data-testid="stTextArea"] textarea:disabled {
        color: #111827 !important;
        -webkit-text-fill-color: #111827 !important;
        opacity: 1 !important;
        font-size: 15px !important;
    }

    /* INPUT FORMULA: style wrapper ngoài */
    div[data-testid="stTextArea"]:first-of-type > div {
        background: #ffffff !important;
        border: 1px solid #d1d5db !important;
        border-radius: 14px !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05) !important;
        overflow: hidden !important;
    }

    /* INPUT FORMULA: textarea bên trong không còn border riêng */
    div[data-testid="stTextArea"]:first-of-type textarea {
        background: transparent !important;
        border: none !important;
        outline: none !important;
        box-shadow: none !important;
        padding: 12px !important;
    }

    div[data-testid="stTextArea"]:first-of-type textarea:focus {
        border: none !important;
        outline: none !important;
        box-shadow: none !important;
    }

    /* Focus cho wrapper ngoài */
    div[data-testid="stTextArea"]:first-of-type > div:focus-within {
        border: 1px solid #3d9e9d !important;
        box-shadow: 0 0 0 2px rgba(61,158,157,0.15) !important;
    }

    div[data-testid="stTextArea"]:first-of-type textarea::placeholder {
        color: #9ca3af !important;
        opacity: 1 !important;
        font-style: italic;
    }

    /* OUTPUT FORMULA: giữ nền xám */
    div[data-testid="stTextArea"]:nth-of-type(2) > div {
        background: #f0f2f6 !important;
        border: none !important;
        border-radius: 14px !important;
        box-shadow: none !important;
        overflow: hidden !important;
    }

    div[data-testid="stTextArea"]:nth-of-type(2) textarea,
    div[data-testid="stTextArea"]:nth-of-type(2) textarea:disabled {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        padding: 12px !important;
    }

    /* Translate button */
    div[data-testid="stButton"] {
        display: flex;
        justify-content: center;
    }

    div[data-testid="stButton"] > button {
        min-width: 140px !important;
        width: 140px !important;
        height: 40px !important;
        padding: 0 24px !important;
        border-radius: 999px !important;
        font-size: 16px !important;
        font-weight: 500 !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        background: #3d9e9d !important;
        color: white !important;
        border: none !important;
        box-shadow: 0 6px 16px rgba(61, 158, 157, 0.35) !important;
        transition: all 0.2s ease !important;
    }

    div[data-testid="stButton"] > button:hover {
        background: #348b8a !important;
        box-shadow: 0 8px 20px rgba(61, 158, 157, 0.45) !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

try:
    mapping = load_mapping()
    reverse = build_reverse_lookup(mapping)
except Exception as e:
    st.error(f"Lỗi đọc file Excel: {e}")
    st.stop()

if "translated_result" not in st.session_state:
    st.session_state.translated_result = ""

with open("logo_xanh.png", "rb") as f:
    logo_base64 = base64.b64encode(f.read()).decode()

st.markdown(
    f"""
    <div style="text-align:center; margin-top:10px;">
        <img src="data:image/png;base64,{logo_base64}"
             style="width:80px; border-radius:16px;" />
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <h1 style="
        margin: 10px 0 20px 0;
        color: #3d9e9d;
        text-align: center;
    ">
        Excel Formula Translator
    </h1>
    """,
    unsafe_allow_html=True,
)
with st.sidebar:
    st.markdown("### Language")
    src = st.selectbox("Source language", SUPPORTED_LANGS)
    tgt = st.selectbox("Target language", SUPPORTED_LANGS, index=2)
    
    st.markdown("<hr style='margin:20px 0; opacity:0.3;'>", unsafe_allow_html=True)
    
    st.markdown("### Format")
    st.markdown("Argument separator")

    sep = st.radio(
        "",
        [",", ";"],
        horizontal=False,
        index=1,
        label_visibility="collapsed"
    )
    

    st.markdown("""
<style>
div[role="radiogroup"] > label {
    background: #f3f4f6;
    padding: 6px 12px;
    border-radius: 999px;
    margin-right: 6px;
    cursor: pointer;
}

div[role="radiogroup"] > label[data-checked="true"] {
    background: #3d9e9d;
    color: white;
}
</style>
""", unsafe_allow_html=True)
    
    pretty_mode = st.toggle("Pretty format", value=True)

HEIGHT = 320
col1, col2 = st.columns(2, gap="small")

with col1:
    formula = st.text_area(
        "Input formula",
        height=HEIGHT,
        placeholder='Example: =IF(SUM(A1,B1)>10,VLOOKUP(C1,Sheet2!A:B,2,FALSE),"No") \n or: =mid() \n or: today()',
        key="input_formula_box",
    )

with col2:
    st.text_area(
        "Translated formula",
        st.session_state.translated_result,
        height=HEIGHT,
        disabled=True,
    )

btn_col1, btn_col2 = st.columns(2, gap="small")

with btn_col1:
    left_pad, center_btn, right_pad = st.columns([2, 1, 2])
    with center_btn:
        run = st.button("Translate", type="primary")

with btn_col2:
    if st.session_state.translated_result:
        copy_button(st.session_state.translated_result)

if run:
    if not formula.strip():
        st.warning("Nhập công thức trước.")
    else:
        st.session_state.translated_result = translate(
            formula, src, tgt, sep, pretty_mode, mapping, reverse
        )
        st.rerun()

st.markdown(
    """
    <hr style="margin-top:40px; margin-bottom:10px;">

    <div style="
        text-align:center;
        color:#6b7280;
        font-size:14px;
    ">
        This product is developed by 
        <b style="color:#3d9e9d;">Quyendatahub</b>. 
        Visit 
        <a href="https://quyendatahub.com" target="_blank" style="color:#3d9e9d; text-decoration:none;">
            quyendatahub.com
        </a> 
        for more data analytics tools and courses.
    </div>
    """,
    unsafe_allow_html=True,
)
