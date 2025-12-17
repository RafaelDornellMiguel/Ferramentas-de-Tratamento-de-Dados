import streamlit as st
import pandas as pd
import re
import html
import math
import json
from bs4 import BeautifulSoup
from io import BytesIO
from typing import Any, Union

# ===================== 🎛️ THEME (LIGHT/DARK) =====================
st.set_page_config(page_title="Ferramentas de Dados", page_icon="🧹", layout="wide")

if "theme_mode" not in st.session_state:
    st.session_state.theme_mode = "dark"

def inject_theme(theme: str) -> None:
    if theme == "dark":
        st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        :root{
          --accent:#F57C00; --accent-dark:#E65100;
          --bg:#0B0F14; --panel:#111827; --panel2:#0F172A;
          --text:#E5E7EB; --muted:#9CA3AF; --border:#243042; --success:#22C55E;
        }
        html,body,[class^="stApp"]{background:var(--bg)!important;color:var(--text)!important;font-family:'Inter',sans-serif;}
        h1{font-weight:900;text-align:center;color:var(--text)!important;}
        /* Tabs (evita “apagado”) */
        button[role="tab"]{color:var(--muted)!important;font-weight:800!important;border-bottom:2px solid transparent!important;
          padding:10px 14px!important;border-radius:10px!important;background:transparent!important;}
        button[role="tab"][aria-selected="true"]{color:var(--text)!important;background:rgba(245,124,0,.12)!important;border-bottom:2px solid var(--accent)!important;}
        [data-testid="stTabs"] div[role="tablist"]{border-bottom:1px solid var(--border)!important;}
        /* Uploader */
        section[data-testid="stFileUploader"]{border:2px dashed var(--accent)!important;background:rgba(245,124,0,.06)!important;border-radius:12px!important;padding:1rem!important;}
        section[data-testid="stFileUploader"] *{color:var(--text)!important;}
        /* Buttons */
        .stButton>button{background:var(--accent)!important;color:#fff!important;font-weight:900!important;border:none!important;border-radius:10px!important;padding:.65rem 1.25rem!important;}
        .stButton>button:hover{background:var(--accent-dark)!important;}
        .stDownloadButton>button{background:var(--success)!important;color:#052e14!important;font-weight:900!important;border:none!important;border-radius:10px!important;}
        /* DataFrame */
        [data-testid="stDataFrame"], .stDataFrame{background:var(--panel)!important;border:1px solid var(--border)!important;border-radius:12px!important;}
        /* Inputs */
        input, textarea, [data-baseweb="select"]{background:var(--panel)!important;color:var(--text)!important;border:1px solid var(--border)!important;border-radius:10px!important;}
        </style>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        :root{
          --accent:#F57C00; --accent-dark:#E65100;
          --bg:#FFFFFF; --panel:#FFFFFF;
          --text:#212121; --muted:#6B7280; --border:#E5E7EB; --success:#22C55E;
        }
        html,body,[class^="stApp"]{background:var(--bg)!important;color:var(--text)!important;font-family:'Inter',sans-serif;}
        h1{font-weight:900;text-align:center;color:var(--accent)!important;}
        button[role="tab"]{color:var(--muted)!important;font-weight:800!important;border-bottom:2px solid transparent!important;
          padding:10px 14px!important;border-radius:10px!important;background:transparent!important;}
        button[role="tab"][aria-selected="true"]{color:var(--text)!important;background:rgba(245,124,0,.10)!important;border-bottom:2px solid var(--accent)!important;}
        [data-testid="stTabs"] div[role="tablist"]{border-bottom:1px solid var(--border)!important;}
        section[data-testid="stFileUploader"]{border:2px dashed var(--accent)!important;background:#FFF8F3!important;border-radius:12px!important;padding:1rem!important;}
        .stButton>button{background:var(--accent)!important;color:#fff!important;font-weight:900!important;border:none!important;border-radius:10px!important;padding:.65rem 1.25rem!important;}
        .stButton>button:hover{background:var(--accent-dark)!important;}
        .stDownloadButton>button{background:var(--success)!important;color:#053014!important;font-weight:900!important;border:none!important;border-radius:10px!important;}
        [data-testid="stDataFrame"], .stDataFrame{border:1px solid var(--border)!important;border-radius:12px!important;}
        </style>
        """, unsafe_allow_html=True)

# Toggle
top = st.columns([6, 1])
with top[1]:
    is_dark = st.toggle("🌙", value=(st.session_state.theme_mode == "dark"))
st.session_state.theme_mode = "dark" if is_dark else "light"
inject_theme(st.session_state.theme_mode)

# ===================== 🧾 TITLE =====================
st.markdown("<h1>🧰 Ferramentas de Tratamento de Dados</h1>", unsafe_allow_html=True)

# ===================== 🧩 CONST / REGEX =====================
RE_HTML_TAG   = re.compile(r"<[^>]+>")
RE_ENTITIES   = re.compile(r"&[a-zA-Z#0-9]+;")
RE_STYLE_TAG  = re.compile(r"<style[\s\S]*?>[\s\S]*?</style>", flags=re.I)
RE_CSS_BLOCK  = re.compile(r"\{[^{}]*:[^{};]+;[^{}]*\}")
RE_CTRL       = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]")

# ===================== 🧠 EXCEL CLEAN PIPELINE =====================
def is_nan(x: Any) -> bool:
    return x is None or (isinstance(x, float) and math.isnan(x))

def to_str(x: Any) -> str:
    return "" if is_nan(x) else str(x)

def deep_unescape(s: str, rounds: int = 5) -> str:
    prev = s
    for _ in range(rounds):
        cur = html.unescape(prev)
        if cur == prev:
            break
        prev = cur
    return prev

def sanitize_weird_chars(s: str) -> str:
    if not s:
        return s
    s = s.replace("\u00A0", " ").replace("\u200B", "").replace("\uFEFF", "")
    return RE_CTRL.sub(" ", s)

def looks_like_html(s: str) -> bool:
    return ("<" in s and ">" in s) or ("&" in s and ";" in s)

def clean_html_css(raw: Any) -> str:
    s = sanitize_weird_chars(to_str(raw)).strip()
    if not s:
        return s

    s = deep_unescape(s, rounds=5)
    s = RE_STYLE_TAG.sub(" ", s)
    s = RE_CSS_BLOCK.sub(" ", s)

    if looks_like_html(s):
        text = BeautifulSoup(s, "html.parser").get_text(separator=" ", strip=True)
    else:
        text = s

    text = RE_HTML_TAG.sub(" ", text)
    text = RE_ENTITIES.sub(" ", text)

    text = re.sub(r"([.,!?;])([A-Za-zÀ-ÿ])", r"\1 \2", text)
    return re.sub(r"\s+", " ", text).strip()

def smart_spacing(text: str) -> str:
    if not text:
        return text
    text = re.sub(r":\s*", ": ", text)
    text = re.sub(r"\?\s*", "? ", text)
    text = re.sub(r"(?<=[a-záéíóúãõç])(?=[A-ZÁÉÍÓÚÂÊÔÃÕÇ0-9])", " ", text)
    text = re.sub(r"(?<=\S)(?=[A-ZÁÉÍÓÚÂÊÔÃÕÇ][a-zà-ÿ\s]{0,20}:)", " ", text)
    text = re.sub(r"([A-Za-zÀ-ÿ])(\d)", r"\1 \2", text)
    text = re.sub(r"(\d)([A-Za-zÀ-ÿ])", r"\1 \2", text)
    return re.sub(r"\s+", " ", text).strip()

def clean_and_polish_cell(raw: Any) -> Any:
    if is_nan(raw):
        return raw
    return smart_spacing(clean_html_css(raw))

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    return df.copy().applymap(clean_and_polish_cell)

def count_dirty_cells(df: pd.DataFrame) -> int:
    def dirty(v: Any) -> bool:
        s = to_str(v)
        if not s:
            return False
        low = s.lower()
        return bool(RE_ENTITIES.search(s) or RE_HTML_TAG.search(s) or "<br" in low or "<style" in low)
    return int(df.applymap(dirty).values.sum())

# ===================== 📦 JSON -> EXCEL =====================
def json_load_bytes(file_bytes: bytes) -> Union[dict, list]:
    text = file_bytes.decode("utf-8-sig", errors="replace")
    return json.loads(text)

def json_to_dataframe(data: Union[dict, list]) -> pd.DataFrame:
    if isinstance(data, list):
        return pd.json_normalize(data)
    if isinstance(data, dict):
        return pd.json_normalize(data)
    raise ValueError("JSON precisa ser um objeto (dict) ou lista (list).")

def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf.getvalue()

# ===================== 🧭 UI =====================
tab_excel, tab_json = st.tabs(["🧹 Limpar (Excel)", "📦 Converter (JSON)"])

with tab_excel:
    st.markdown("### 🧹 Limpeza + tratamento de espaços (Excel)")
    up_xlsx = st.file_uploader("📤 Envie o arquivo Excel (.xlsx):", type=["xlsx"], key="xlsx_uploader")
    if up_xlsx:
        try:
            df = pd.read_excel(up_xlsx)
            st.success("✅ Arquivo carregado com sucesso!")
            st.info(f"🔎 Células com possível HTML/entidades: {count_dirty_cells(df)}")

            t1, t2 = st.tabs(["📄 Original", "🧽 Resultado"])
            with t1:
                st.dataframe(df.head(20), use_container_width=True)

            with t2:
                if st.button("🚀 Limpar e formatar", key="btn_clean_excel"):
                    df_clean = clean_dataframe(df)
                    st.subheader("🧾 Prévia após tratamento")
                    st.dataframe(df_clean.head(20), use_container_width=True)

                    st.download_button(
                        "⬇️ Baixar Excel limpo (.xlsx)",
                        data=df_to_xlsx_bytes(df_clean),
                        file_name="planilha_tratada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_excel_clean",
                    )
                    st.download_button(
                        "⬇️ Baixar CSV limpo (.csv)",
                        data=df_clean.to_csv(index=False).encode("utf-8-sig"),
                        file_name="planilha_tratada.csv",
                        mime="text/csv",
                        key="dl_csv_clean",
                    )
        except Exception as e:
            st.error(f"❌ Erro ao processar Excel: {e}")
    else:
        st.info("☝️ Envie um `.xlsx` para começar.")

with tab_json:
    st.markdown("### 📦 Converter JSON em Excel")
    st.write("Envie um `.json` e baixe como `.xlsx` ou `.csv`.")
    up_json = st.file_uploader("📤 Envie o arquivo JSON (.json):", type=["json"], key="json_uploader")
    if up_json:
        try:
            data = json_load_bytes(up_json.getvalue())
            df_json = json_to_dataframe(data)

            st.success("✅ JSON convertido em tabela!")
            st.dataframe(df_json.head(50), use_container_width=True)

            st.download_button(
                "⬇️ Baixar Excel (.xlsx)",
                data=df_to_xlsx_bytes(df_json),
                file_name="json_convertido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_json_xlsx",
            )
            st.download_button(
                "⬇️ Baixar CSV (.csv)",
                data=df_json.to_csv(index=False).encode("utf-8-sig"),
                file_name="json_convertido.csv",
                mime="text/csv",
                key="dl_json_csv",
            )
        except Exception as e:
            st.error(f"❌ Erro ao processar JSON: {e}")
    else:
        st.info("☝️ Envie um `.json` para converter.")
