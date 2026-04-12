# Installatiehandleiding

Stap-voor-stap uitleg om de app lokaal te installeren op je computer.
Geen technische kennis vereist.

---

## Optie 1: Online gebruiken (geen installatie)

De makkelijkste manier. Open deze link in je browser:

**https://elopdftoxslx.streamlit.app**

Klaar. Je kunt direct een PDF uploaden en aan de slag.

> **Let op:** als je de app online gebruikt worden API keys niet opgeslagen.
> Bij lokaal gebruik kun je API keys wel opslaan zodat je ze niet elke keer opnieuw hoeft in te vullen.

---

## Optie 2: Lokaal installeren

### Windows

#### 1. Python installeren

1. Ga naar **https://www.python.org/downloads/**
2. Klik op de grote gele knop **"Download Python 3.x.x"**
3. Open het gedownloade bestand
4. **Belangrijk:** Zet het vinkje aan bij **"Add Python to PATH"** (staat onderaan het scherm)
5. Klik op **"Install Now"**
6. Wacht tot de installatie klaar is en klik op **"Close"**

#### 2. De app downloaden

1. Open **PowerShell** (druk op de Windows-toets, typ `powershell`, druk Enter)
2. Kopieer en plak dit commando en druk Enter:

```
git clone https://github.com/ezero-m/ELOpdftoXSLX.git
```

3. Ga naar de map:

```
cd ELOpdftoXSLX
```

4. Installeer de benodigde software:

```
pip install -r requirements.txt
```

Dit kan een paar minuten duren. Wacht tot het klaar is.

> **Geen git?** Download dan de code als ZIP: ga naar
> https://github.com/ezero-m/ELOpdftoXSLX → klik op de groene knop **"Code"** →
> **"Download ZIP"**. Pak het ZIP-bestand uit en open PowerShell in die map.

#### 3. De app starten

In dezelfde PowerShell, typ:

```
python -m streamlit run app.py
```

Je browser opent automatisch. Als dat niet gebeurt, ga dan naar **http://localhost:8501**

#### Volgende keer starten

Je hoeft stap 1 en 2 maar een keer te doen. De volgende keer open je alleen PowerShell en typ:

```
cd ELOpdftoXSLX
python -m streamlit run app.py
```

---

### Mac

#### 1. Terminal openen

Druk op `⌘ + spatie` (Command + spatiebalk), typ **Terminal**, druk Enter.

#### 2. Python installeren

Typ in Terminal:

```
python3 --version
```

- Zie je een versienummer (bijv. `Python 3.12.0`)? Ga door naar stap 3.
- Zie je een foutmelding? Installeer Python:

```
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

Wacht tot dit klaar is (kan een paar minuten duren). Typ daarna:

```
brew install python
```

#### 3. De app downloaden

Typ in Terminal:

```
git clone https://github.com/ezero-m/ELOpdftoXSLX.git
```

Ga naar de map:

```
cd ELOpdftoXSLX
```

Installeer de benodigde software:

```
pip3 install -r requirements.txt
```

> **Geen git?** Download de code als ZIP: ga in Safari naar
> https://github.com/ezero-m/ELOpdftoXSLX → klik op de groene knop **"Code"** →
> **"Download ZIP"**. Pak uit en open Terminal in die map.

#### 4. De app starten

```
python3 -m streamlit run app.py
```

Je browser opent automatisch. Als dat niet gebeurt, ga dan naar **http://localhost:8501**

#### Volgende keer starten

```
cd ELOpdftoXSLX
python3 -m streamlit run app.py
```

---

## Veelgestelde vragen

### Wat is een API key?

Een API key is een soort wachtwoord waarmee de app verbinding maakt met een AI-dienst
(zoals ChatGPT of Claude). Je maakt er eentje aan op de website van de dienst.
De key is een lange reeks letters en cijfers die er ongeveer zo uitziet:
`sk-ant-abc123def456...`

### Waar maak ik een API key aan?

| Dienst | Website |
|--------|---------|
| Claude (Anthropic) | https://console.anthropic.com/settings/keys |
| OpenAI (ChatGPT) | https://platform.openai.com/api-keys |
| Google Gemini | https://aistudio.google.com/apikey |

Bij Ollama heb je geen API key nodig (draait lokaal).

### Worden mijn API keys veilig bewaard?

Ja. Als je bij lokaal gebruik kiest voor "API key onthouden", wordt de key opgeslagen
in een bestand op je eigen computer (`~/.elo_pdf_to_xlsx_keys.json`).
De key wordt nergens anders heen gestuurd dan naar de AI-dienst die je hebt gekozen.

### Kan ik de app ook zonder AI gebruiken?

Ja! Bij stap 2 in de app kun je kiezen voor "Handmatig". Dan download je de lesstof-tekst
of een kant-en-klare prompt, plak je die in een AI naar keuze (bijv. ChatGPT.com of
Claude.ai), en plak je het antwoord terug in de app. Zo heb je geen API key nodig.

### Ik krijg een foutmelding bij het installeren

- **"pip niet gevonden"**: probeer `pip3` in plaats van `pip`
- **"git niet gevonden"**: download de ZIP in plaats daarvan (zie instructies hierboven)
- **"Permission denied"**: op Mac, zet `sudo` voor het commando
  (bijv. `sudo pip3 install -r requirements.txt`) en vul je Mac-wachtwoord in
