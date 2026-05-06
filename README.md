# ELO PDF naar Thinkific Quiz Converter

Zet e-learning PDF's om naar quizvragen die je direct kunt importeren in Thinkific.

## Wat doet deze tool?

Deze tool haalt de lesstof uit een e-learning PDF en helpt je om daar meerkeuzevragen van te maken in het juiste formaat voor Thinkific. Je kunt vragen laten genereren door een AI-dienst, of de lesstof handmatig kopieren naar een AI naar keuze.

### Mogelijkheden

- PDF uploaden en automatisch tekst extraheren
- Vragen genereren via **Claude**, **ChatGPT**, **Google Gemini** of **Ollama** (lokaal)
- Of: lesstof/prompt kopieren, in een externe AI plakken, en het resultaat terugplakken
- Vragen per stuk controleren en aanpassen
- Exporteren naar Thinkific XLSX-importbestand

### Auteursrecht

De prompt instrueert de AI nadrukkelijk om vragen **in eigen formulering** te maken
en NOOIT letterlijke zinsdelen uit de bron-PDF over te nemen. De gegenereerde vragen
toetsen dezelfde kennis, maar via eigen woordkeuze, scenario's en invalshoek. Cijfers,
normen en vakbegrippen blijven uiteraard hetzelfde — die zijn feitelijk.

Controleer altijd de output zelf voordat je publiceert.

## Snel starten (online)

Open de gehoste versie:

> **[ELO PDF to XLSX op Streamlit Cloud](https://elopdftoxslx.streamlit.app)**

Geen installatie nodig. Werkt op elke computer met een browser.

## Lokaal draaien

Zie **[INSTALLATIE.md](INSTALLATIE.md)** voor een uitgebreide stap-voor-stap handleiding voor Windows en Mac (geen technische kennis vereist).

Bij lokaal gebruik kun je API keys opslaan zodat je ze niet elke keer opnieuw hoeft in te vullen.

### Snelle versie

```bash
git clone https://github.com/ezero-m/ELOpdftoXSLX.git
cd ELOpdftoXSLX
pip install -r requirements.txt
```

### Starten

**Web-interface (aanbevolen):**

```bash
python -m streamlit run app.py
```

Open daarna http://localhost:8501 in je browser.

**Command-line (voor gevorderden):**

```bash
# Alleen tekst uit PDF halen
python pdf_to_thinkific.py "Module.pdf" --text-only

# AI prompt genereren om extern te gebruiken
python pdf_to_thinkific.py "Module.pdf" --prompt-only

# Vragen genereren met een AI-dienst
python pdf_to_thinkific.py "Module.pdf" --ai claude
python pdf_to_thinkific.py "Module.pdf" --ai openai
python pdf_to_thinkific.py "Module.pdf" --ai gemini
python pdf_to_thinkific.py "Module.pdf" --ai ollama

# JSON met vragen omzetten naar XLSX
python pdf_to_thinkific.py --from-json vragen.json
```

## Hoe werkt het?

### Stap 1 - PDF uploaden

Upload een PDF van een e-learning module. De tool extraheert automatisch de lesstof en filtert navigatie-elementen en andere ruis.

### Stap 2 - Vragen maken

Kies een van twee routes:

**Route A: Handmatig (zonder AI in de app)**
1. Download de lesstof-tekst of een kant-en-klare AI-prompt
2. Plak het in een AI naar keuze (ChatGPT, Claude.ai, Gemini, etc.)
3. Kopieer het JSON-antwoord terug in de app

**Route B: Automatisch met AI**
1. Kies een AI-dienst (Claude, OpenAI, Gemini of Ollama)
2. Vul je API key in (niet nodig bij Ollama)
3. Klik op "Vragen genereren"

### Stap 3 - Controleren

Bekijk elke vraag, pas teksten aan en selecteer welke vragen je wilt meenemen.

### Stap 4 - Exporteren

Download het XLSX-bestand en importeer het in Thinkific via de quiz-instellingen.

## AI-diensten

| Dienst | API key nodig | Aanmaken |
|--------|:---:|----------|
| Claude (Anthropic) | Ja | [console.anthropic.com](https://console.anthropic.com/settings/keys) |
| OpenAI (ChatGPT) | Ja | [platform.openai.com](https://platform.openai.com/api-keys) |
| Google Gemini | Ja | [aistudio.google.com](https://aistudio.google.com/apikey) |
| Ollama (lokaal) | Nee | [ollama.com](https://ollama.com) |

API keys worden niet opgeslagen en alleen gebruikt tijdens je sessie.

## Thinkific importformaat

Het gegenereerde XLSX-bestand bevat:
- **QuestionType**: `SA` (een correct antwoord) of `MA` (meerdere correcte antwoorden)
- **QuestionText**: de vraagtekst
- **Explanation**: toelichting bij het correcte antwoord
- **Choice1 t/m Choice10**: antwoordopties, correct antwoord gemarkeerd met `*`
