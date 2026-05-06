"""
E-learning PDF naar Thinkific Quiz Converter

Streamlit web-app die e-learning PDF's omzet naar
quizvragen in Thinkific importformaat (.xlsx).
"""

import json
import os
from pathlib import Path

import streamlit as st

from pdf_to_thinkific import (
    AI_PROVIDERS,
    extract_text_from_bytes,
    build_prompt,
    generate_questions,
    parse_questions_json,
    export_xlsx_to_bytes,
)

# ---------------------------------------------------------------------------
# API key storage (local only, never on Streamlit Cloud)
# ---------------------------------------------------------------------------
KEY_FILE = Path.home() / ".elo_pdf_to_xlsx_keys.json"
IS_LOCAL = not os.environ.get("STREAMLIT_SHARING_MODE") and not os.environ.get("STREAMLIT_SERVER_HEADLESS")


def _load_saved_keys() -> dict:
    if KEY_FILE.exists():
        try:
            return json.loads(KEY_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def _save_key(provider: str, key: str):
    keys = _load_saved_keys()
    keys[provider] = key
    KEY_FILE.write_text(json.dumps(keys, indent=2), encoding="utf-8")


def _delete_key(provider: str):
    keys = _load_saved_keys()
    keys.pop(provider, None)
    KEY_FILE.write_text(json.dumps(keys, indent=2), encoding="utf-8")

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="ELO PDF naar Thinkific Quiz",
    page_icon="📝",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Custom styling
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    /* Bigger step headers */
    .step-header {
        background: linear-gradient(135deg, #1e3a5f, #2d5f8a);
        color: white;
        padding: 12px 20px;
        border-radius: 8px;
        margin: 20px 0 10px 0;
        font-size: 1.1em;
    }
    /* Question cards */
    .question-card {
        background: #f8f9fa;
        border-left: 4px solid #2d5f8a;
        padding: 15px;
        border-radius: 0 8px 8px 0;
        margin: 8px 0;
    }
    .correct-answer { color: #28a745; font-weight: bold; }
    .wrong-answer { color: #6c757d; }
    /* Info boxes */
    .info-box {
        background: #e8f4f8;
        border: 1px solid #bee5eb;
        border-radius: 8px;
        padding: 15px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Session state defaults
# ---------------------------------------------------------------------------
for key, default in {
    "extracted_text": None,
    "questions": None,
    "selected": None,
    "pdf_name": None,
    "step": 1,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ---------------------------------------------------------------------------
# Title
# ---------------------------------------------------------------------------
st.title("📝 ELO PDF naar Thinkific Quiz Converter")
st.caption("Zet e-learning PDF's om naar quizvragen voor Thinkific")

st.info(
    "**Auteursrecht:** vragen worden in eigen formulering gegenereerd, niet letterlijk uit de "
    "bron-PDF gekopieerd. Controleer toch altijd zelf de output voordat je deze publiceert."
)

# ---------------------------------------------------------------------------
# Sidebar: how it works
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Hoe werkt het?")
    st.markdown("""
**Stap 1** — Upload een PDF van de leeromgeving

**Stap 2** — Kies hoe je vragen wilt maken:
- *Handmatig*: kopieer de tekst naar een AI naar keuze
- *Automatisch*: laat de app het doen via een AI-dienst

**Stap 3** — Controleer de vragen en pas aan

**Stap 4** — Download het Thinkific-importbestand (.xlsx)
    """)

    st.divider()
    st.markdown("**Tip:** Je kunt altijd terug naar een vorige stap.")
    st.divider()
    st.caption("ELO PDF to XLSX v1.1")


# ===================================================================
# STAP 1: PDF uploaden
# ===================================================================
st.markdown('<div class="step-header">Stap 1: PDF uploaden</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Kies een PDF bestand van de leeromgeving",
    type=["pdf"],
    help="Upload een PDF van een e-learning module.",
)

if uploaded_file:
    if st.session_state.pdf_name != uploaded_file.name:
        with st.spinner("Tekst uit PDF halen..."):
            pdf_bytes = uploaded_file.read()
            st.session_state.extracted_text = extract_text_from_bytes(pdf_bytes)
            st.session_state.pdf_name = uploaded_file.name
            st.session_state.questions = None
            st.session_state.selected = None

    text = st.session_state.extracted_text
    st.success(f"**{uploaded_file.name}** geladen — {len(text):,} tekens geextraheerd")

    with st.expander("Bekijk de geextraheerde tekst"):
        st.text_area("Lesstof tekst", text, height=300, disabled=True, label_visibility="collapsed")

else:
    st.info("Upload een PDF om te beginnen.")
    st.stop()


# ===================================================================
# STAP 2: Vragen genereren
# ===================================================================
st.markdown('<div class="step-header">Stap 2: Vragen maken</div>', unsafe_allow_html=True)

st.markdown("Kies hoe je de quizvragen wilt aanmaken:")

tab_manual, tab_ai = st.tabs([
    "📋 Handmatig (zonder AI in deze app)",
    "🤖 Automatisch met AI",
])

# --- Tab: Handmatig ---
with tab_manual:
    st.markdown("""
    **Hoe werkt dit?**
    1. Kopieer de lesstof-tekst of de kant-en-klare AI-prompt hieronder
    2. Plak het in een AI-tool naar keuze (bijv. ChatGPT, Claude.ai, Gemini)
    3. Kopieer het JSON-antwoord van de AI
    4. Plak het hieronder of laad een JSON-bestand
    """)

    col1, col2 = st.columns(2)
    with col1:
        prompt_text = build_prompt(text)
        st.download_button(
            "⬇️ Download AI-prompt als tekstbestand",
            data=prompt_text,
            file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}_prompt.txt",
            mime="text/plain",
            help="Download een tekstbestand met de volledige prompt. Open het en kopieer de inhoud naar een AI.",
            use_container_width=True,
        )
    with col2:
        st.download_button(
            "⬇️ Download lesstof als tekstbestand",
            data=text,
            file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}_tekst.txt",
            mime="text/plain",
            help="Download alleen de schone lesstof-tekst.",
            use_container_width=True,
        )

    st.divider()
    st.markdown("**Plak hier het JSON-antwoord van de AI:**")

    json_input = st.text_area(
        "JSON invoer",
        height=200,
        placeholder='[\n  {"type": "SA", "question": "...", "explanation": "...", "choices": [...]}\n]',
        label_visibility="collapsed",
    )

    json_file = st.file_uploader(
        "Of laad een JSON-bestand",
        type=["json", "txt"],
        key="json_upload",
    )

    if st.button("Vragen laden", key="load_manual", use_container_width=True, type="primary"):
        raw = None
        if json_file:
            raw = json_file.read().decode("utf-8")
        elif json_input.strip():
            raw = json_input.strip()

        if not raw:
            st.warning("Plak JSON-tekst of upload een bestand.")
        else:
            try:
                questions = parse_questions_json(raw)
                st.session_state.questions = questions
                st.session_state.selected = [True] * len(questions)
                st.success(f"{len(questions)} vragen geladen!")
                st.rerun()
            except (json.JSONDecodeError, ValueError) as e:
                st.error(f"Kon de JSON niet lezen: {e}")
                with st.expander("Verwacht formaat"):
                    st.code("""[
  {
    "type": "SA",
    "question": "Vraagtekst?",
    "explanation": "Uitleg.",
    "choices": [
      {"text": "Correct antwoord", "correct": true},
      {"text": "Fout antwoord", "correct": false}
    ]
  }
]""", language="json")


# --- Tab: Automatisch met AI ---
with tab_ai:
    st.markdown("Laat de app automatisch vragen genereren via een AI-dienst.")

    provider = st.selectbox(
        "Kies een AI-dienst",
        options=list(AI_PROVIDERS.keys()),
        format_func=lambda k: AI_PROVIDERS[k]["name"],
    )
    info = AI_PROVIDERS[provider]

    model = st.selectbox(
        "Model",
        options=info["models"],
        index=0,
        help="Het standaard model werkt voor de meeste situaties prima.",
    )

    api_key = None
    if info["needs_api_key"]:
        saved_keys = _load_saved_keys()
        saved_key = saved_keys.get(provider, "")

        st.markdown(f"""
        **API key nodig** — [Maak er hier een aan]({info['key_url']})

        Een API key is een soort wachtwoord waarmee de app verbinding maakt met de AI-dienst.
        """)

        if saved_key:
            st.success(f"Opgeslagen API key gevonden voor {info['name']}.")
            use_saved = st.checkbox("Opgeslagen key gebruiken", value=True, key="use_saved")
            if use_saved:
                api_key = saved_key
            else:
                api_key = st.text_input(
                    f"{info['name']} API Key",
                    type="password",
                    placeholder=f"Plak hier je {info['env_key']} key...",
                    help=f"Ga naar {info['key_url']} om een key aan te maken.",
                )
            if st.button("Opgeslagen key verwijderen", key="delete_key"):
                _delete_key(provider)
                st.rerun()
        else:
            api_key = st.text_input(
                f"{info['name']} API Key",
                type="password",
                placeholder=f"Plak hier je {info['env_key']} key...",
                help=f"Ga naar {info['key_url']} om een key aan te maken.",
            )
            if api_key:
                save_key = st.checkbox(
                    "API key onthouden op deze computer",
                    value=False,
                    key="save_key",
                    help="De key wordt lokaal opgeslagen zodat je hem niet elke keer opnieuw hoeft in te vullen.",
                )
                if save_key:
                    _save_key(provider, api_key)
                    st.success("Key opgeslagen!")
    else:
        st.info("""
        **Ollama draait lokaal op je computer** — er is geen API key nodig.
        Zorg dat Ollama geinstalleerd en gestart is (ollama.com).
        """)

    if st.button("🚀 Vragen genereren", key="generate_ai", use_container_width=True, type="primary"):
        if info["needs_api_key"] and not api_key:
            st.warning("Vul eerst een API key in.")
        else:
            with st.spinner(f"Vragen genereren met {info['name']}... Dit kan 30-60 seconden duren."):
                try:
                    questions = generate_questions(text, provider, model, api_key or None)
                    st.session_state.questions = questions
                    st.session_state.selected = [True] * len(questions)
                    st.success(f"{len(questions)} vragen gegenereerd!")
                    st.rerun()
                except Exception as e:
                    error_msg = str(e)
                    st.error(f"Er ging iets mis: {error_msg}")
                    if "api_key" in error_msg.lower() or "auth" in error_msg.lower():
                        st.info("Controleer of je API key correct is en nog geldig.")
                    elif provider == "ollama":
                        st.info("Controleer of Ollama draait (ollama serve) en het model beschikbaar is.")


# ===================================================================
# STAP 3: Vragen controleren
# ===================================================================
if st.session_state.questions is None:
    st.stop()

questions = st.session_state.questions
selected = st.session_state.selected

st.markdown('<div class="step-header">Stap 3: Vragen controleren</div>', unsafe_allow_html=True)

n_selected = sum(selected)
st.markdown(f"**{n_selected} van {len(questions)}** vragen geselecteerd voor export")

col_all, col_none = st.columns(2)
with col_all:
    if st.button("Alles selecteren", use_container_width=True):
        st.session_state.selected = [True] * len(questions)
        st.rerun()
with col_none:
    if st.button("Alles deselecteren", use_container_width=True):
        st.session_state.selected = [False] * len(questions)
        st.rerun()

for i, q in enumerate(questions):
    qtype = q.get("type", "SA")
    type_label = "Eén antwoord" if qtype == "SA" else "Meerdere antwoorden"

    with st.container():
        col_check, col_content = st.columns([0.05, 0.95])

        with col_check:
            new_val = st.checkbox(
                f"q{i}",
                value=selected[i],
                key=f"sel_{i}",
                label_visibility="collapsed",
            )
            if new_val != selected[i]:
                st.session_state.selected[i] = new_val

        with col_content:
            with st.expander(f"**Vraag {i+1}** [{type_label}] — {q['question'][:80]}{'...' if len(q['question']) > 80 else ''}"):
                new_question = st.text_area(
                    "Vraagtekst",
                    value=q["question"],
                    key=f"qt_{i}",
                    height=80,
                )
                if new_question != q["question"]:
                    st.session_state.questions[i]["question"] = new_question

                new_explanation = st.text_area(
                    "Uitleg",
                    value=q.get("explanation", ""),
                    key=f"ex_{i}",
                    height=60,
                )
                if new_explanation != q.get("explanation", ""):
                    st.session_state.questions[i]["explanation"] = new_explanation

                st.markdown("**Antwoorden:**")
                for j, choice in enumerate(q.get("choices", [])):
                    if choice.get("correct"):
                        st.markdown(f"- ✅ **{choice['text']}**")
                    else:
                        st.markdown(f"- ❌ {choice['text']}")


# ===================================================================
# STAP 4: Exporteren
# ===================================================================
st.markdown('<div class="step-header">Stap 4: Exporteren naar Thinkific</div>', unsafe_allow_html=True)

final_questions = [q for q, s in zip(questions, selected) if s]

if not final_questions:
    st.warning("Selecteer minstens 1 vraag om te exporteren.")
    st.stop()

n_sa = sum(1 for q in final_questions if q.get("type") == "SA")
n_ma = sum(1 for q in final_questions if q.get("type") == "MA")

col1, col2, col3 = st.columns(3)
col1.metric("Totaal vragen", len(final_questions))
col2.metric("Eén antwoord (SA)", n_sa)
col3.metric("Meerdere antwoorden (MA)", n_ma)

xlsx_bytes = export_xlsx_to_bytes(final_questions)
base_name = st.session_state.pdf_name.rsplit(".", 1)[0] if st.session_state.pdf_name else "vragen"

st.download_button(
    "⬇️ Download Thinkific XLSX",
    data=xlsx_bytes,
    file_name=f"{base_name}_thinkific.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
    type="primary",
)

st.success("Klaar! Importeer het .xlsx bestand in Thinkific via je quiz-instellingen.")
