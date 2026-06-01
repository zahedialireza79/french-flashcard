# French Flashcards 

Alireza Zahedi - 1 June 2026

A simple desktop flashcard app built with **PyQt6** for studying French vocabulary from an Excel file. Optionally fills in English meanings using a local **Ollama** model.

## Features

- Load French words from an `.xlsx` / `.xls` file
- Flip cards to reveal the meaning (and an optional example sentence)
- Previous / Next navigation, shuffle mode
- Optional: auto-fetch missing meanings from a local Ollama server (runs in the background so the UI stays responsive)
- Fetched meanings are cached to `meanings_cache.json`, so each word is only translated once

## Requirements

- Python 3.9+
- (Optional) [Ollama](https://ollama.com) running locally if you want auto-translation

## Installation

```bash
git clone https://github.com/zahedialireza79/french-flashcard.git
cd french-flashcard

python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate

pip install -r requirements.txt
```

## Usage

```bash
python flashcards.py
```

Then click **Load Excel…** and pick your vocabulary file.

### Excel format

The file should have a header row. Column names are case-insensitive. Only `french` is required.

| french   | meaning            | example                     |
|----------|--------------------|-----------------------------|
| bonjour  | hello, good day    | Bonjour, comment ça va ?    |
| chat     | cat                | Le chat dort sur le canapé. |
| pomme    |                    |                             |

- If `meaning` is empty for a row, you can fetch it from Ollama with the **"Fetch missing meanings from Ollama"** button.
- If the first column isn't named `french`, the app will treat it as the French column anyway.

## Optional: Ollama setup

Install Ollama, then pull a model:

```bash
ollama serve
ollama pull llama3.2
```

The model is configured at the top of `flashcards.py`:

```python
OLLAMA_MODEL = "llama3.2"   # or "mistral", "llama3.1", etc.
```

### Performance notes

Each word is one Ollama call, so timing scales linearly:

| Model         | Approx. time per word | 100 words |
|---------------|-----------------------|-----------|
| llama3.2:3b   | 1–2 s                 | ~2–3 min  |
| llama3.1:8b   | 2–5 s                 | ~5–8 min  |
| mistral:7b    | 2–4 s                 | ~4–7 min  |

(Times vary widely with hardware — much faster on Apple Silicon or a discrete GPU.)

Because results are cached in `meanings_cache.json`, you only pay this cost once per word. If you already know the meanings, putting them directly in the Excel is instant and 100% accurate.

## Project structure

```
french-flashcards/
├── .gitignore
├── README.md
├── requirements.txt
└── flashcards.py
```
