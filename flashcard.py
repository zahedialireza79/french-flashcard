"""
Alireza Zahedi - 1 June 2026
French Flashcards - PyQt6 app
- Loads French words from an Excel file.
- Expected columns (case-insensitive): 'french' (required), 'meaning' (optional), 'example' (optional).
- If 'meaning' is missing/empty for a row, you can fetch it from a local Ollama server.
- Meanings are cached to meanings_cache.json AND written back to the source Excel file,
  so the Excel becomes the long-term source of truth.

Run:
    pip install -r requirements.txt
    python flashcard.py
"""

import json
import sys
from pathlib import Path

import pandas as pd
import requests
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QApplication, QCheckBox, QFileDialog, QHBoxLayout, QLabel,
    QMainWindow, QMessageBox, QProgressBar, QPushButton, QVBoxLayout, QWidget,
)

OLLAMA_URL = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "gemma3"   # change to whatever `ollama list` shows
CACHE_FILE = Path("meanings_cache.json")

SYSTEM_PROMPT = (
    "You are a French-to-English dictionary. "
    "For each French word the user sends, respond with ONLY the English translation. "
    "Rules: "
    "1-3 words maximum. "
    "If multiple common meanings, separate with commas. "
    "No markdown, no quotes, no punctuation at the end, no explanations, no full sentences. "
    "Examples: "
    "Input: chat  Output: cat. "
    "Input: pomme  Output: apple. "
    "Input: voler  Output: to fly, to steal."
)


# ---------- Ollama worker (background thread) ----------

class OllamaWorker(QThread):
    progress = pyqtSignal(int, int, str)      # current, total, current word
    word_done = pyqtSignal(str, str)          # french word, meaning
    finished_all = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, words, model=OLLAMA_MODEL):
        super().__init__()
        self.words = words
        self.model = model
        self._stop = False

    def stop(self):
        self._stop = True

    def run(self):
        total = len(self.words)
        for i, word in enumerate(self.words, start=1):
            if self._stop:
                break
            self.progress.emit(i, total, word)
            try:
                resp = requests.post(
                    OLLAMA_URL,
                    json={
                        "model": self.model,
                        "system": SYSTEM_PROMPT,
                        "prompt": word,
                        "stream": False,
                        "keep_alive": "10m",
                        "options": {
                            "temperature": 0.1,
                            "num_predict": 30,
                        },
                    },
                    timeout=60,
                )
                resp.raise_for_status()
                meaning = resp.json().get("response", "").strip()
                # Light cleanup of common model artifacts
                meaning = meaning.replace("**", "").replace("*", "")
                meaning = meaning.split("\n")[0].strip(" \"'.,;:\n")
                self.word_done.emit(word, meaning or "(no answer)")
            except requests.exceptions.ConnectionError:
                self.error.emit(
                    "Could not reach Ollama at http://localhost:11434.\n"
                    "Run 'ollama serve' to start it."
                )
                return
            except requests.exceptions.HTTPError as e:
                body = ""
                try:
                    body = e.response.text
                except Exception:
                    pass
                self.error.emit(
                    f"Ollama returned {e.response.status_code} for model '{self.model}'.\n\n"
                    f"Response: {body}\n\n"
                    f"Run `ollama list` and make sure OLLAMA_MODEL matches an installed model."
                )
                return
            except Exception as e:
                self.word_done.emit(word, f"(error: {e})")
        self.finished_all.emit()


# ---------- Main window ----------

class FlashcardApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("French Flashcards")
        self.resize(680, 520)

        self.cards = []              # list of dicts: {french, meaning, example}
        self.index = 0
        self.showing_back = False
        self.cache = self._load_cache()
        self.worker = None

        # Excel state — used to write meanings back to the source file
        self.excel_path = None       # Path or None
        self.df = None               # original DataFrame (preserves user's column casing)
        self.french_col = None       # actual column name for french in df
        self.meaning_col = None      # actual column name for meaning in df (created if missing)
        self.example_col = None      # actual column name for example, or None

        self._build_ui()
        self._refresh_card()

    # ---- UI ----
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(15)

        # Top bar
        top = QHBoxLayout()
        self.load_btn = QPushButton("Load Excel…")
        self.load_btn.clicked.connect(self.load_excel)
        top.addWidget(self.load_btn)

        self.fetch_btn = QPushButton("Fetch from Ollama")
        self.fetch_btn.clicked.connect(self.fetch_with_ollama)
        self.fetch_btn.setEnabled(False)
        top.addWidget(self.fetch_btn)

        self.save_btn = QPushButton("Save to Excel")
        self.save_btn.clicked.connect(self.save_to_excel)
        self.save_btn.setEnabled(False)
        top.addWidget(self.save_btn)

        self.shuffle_cb = QCheckBox("Shuffle")
        self.shuffle_cb.stateChanged.connect(self._on_shuffle)
        top.addWidget(self.shuffle_cb)
        top.addStretch()
        root.addLayout(top)

        # Card
        self.card_label = QLabel("Load an Excel file to begin.\n\nClick the card to flip.")
        self.card_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.card_label.setWordWrap(True)
        self.card_label.setStyleSheet(
            "background:#fafafa; border:2px solid #ddd; border-radius:12px; padding:30px;"
        )
        self.card_label.setFont(QFont("Helvetica", 28))
        self.card_label.mousePressEvent = lambda e: self.flip()
        root.addWidget(self.card_label, stretch=1)

        # Counter
        self.counter = QLabel("")
        self.counter.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(self.counter)

        # Nav
        nav = QHBoxLayout()
        self.prev_btn = QPushButton("◀ Previous")
        self.prev_btn.clicked.connect(self.prev_card)
        self.flip_btn = QPushButton("Flip")
        self.flip_btn.clicked.connect(self.flip)
        self.next_btn = QPushButton("Next ▶")
        self.next_btn.clicked.connect(self.next_card)
        for b in (self.prev_btn, self.flip_btn, self.next_btn):
            nav.addWidget(b)
        root.addLayout(nav)

        # Progress
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        # Status
        self.status = QLabel("")
        root.addWidget(self.status)

    # ---- Excel loading ----
    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel file", "", "Excel files (*.xlsx *.xls)"
        )
        if not path:
            return
        try:
            df = pd.read_excel(path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not read file:\n{e}")
            return

        # Find columns case-insensitively, preserve original names for writing back
        col_map = {str(c).strip().lower(): c for c in df.columns}
        french_col = col_map.get("french")
        if french_col is None:
            # Fall back to first column
            french_col = df.columns[0]
        meaning_col = col_map.get("meaning")
        example_col = col_map.get("example")

        # If there's no meaning column yet, add one so we can write back into it
        if meaning_col is None:
            meaning_col = "meaning"
            df[meaning_col] = ""

        # Build cards
        self.cards = []
        from_cache = 0
        for _, row in df.iterrows():
            word = str(row.get(french_col, "")).strip()
            if not word or word.lower() == "nan":
                continue
            meaning = str(row.get(meaning_col, "")).strip()
            if meaning.lower() == "nan":
                meaning = ""
            if not meaning and word in self.cache:
                meaning = self.cache[word]
                from_cache += 1
            example = ""
            if example_col is not None:
                example = str(row.get(example_col, "")).strip()
                if example.lower() == "nan":
                    example = ""
            self.cards.append({"french": word, "meaning": meaning, "example": example})

        if not self.cards:
            QMessageBox.warning(self, "Empty", "No valid French words found.")
            return

        # Remember Excel state for write-back
        self.excel_path = Path(path)
        self.df = df
        self.french_col = french_col
        self.meaning_col = meaning_col
        self.example_col = example_col

        self.index = 0
        self.showing_back = False

        missing = sum(1 for c in self.cards if not c["meaning"])
        self.fetch_btn.setEnabled(True)
        self.save_btn.setEnabled(True)

        parts = [f"Loaded {len(self.cards)} cards"]
        if from_cache:
            parts.append(f"{from_cache} from cache")
        if missing:
            parts.append(f"{missing} missing")
        self.status.setText("  —  ".join(parts))
        self._refresh_card()

    # ---- Excel write-back ----
    def save_to_excel(self, silent=False):
        """Write current meanings back into the source Excel file."""
        if self.df is None or self.excel_path is None:
            return False

        # Build lookup: french word -> meaning
        meanings_by_word = {c["french"]: c["meaning"] for c in self.cards}

        def lookup(val):
            key = str(val).strip()
            return meanings_by_word.get(key, "")

        # Only overwrite cells where we have a non-empty meaning,
        # so we don't accidentally blank out user-edited rows.
        new_meanings = self.df[self.french_col].map(lookup)
        mask = new_meanings.astype(str).str.len() > 0
        self.df.loc[mask, self.meaning_col] = new_meanings[mask]

        try:
            self.df.to_excel(self.excel_path, index=False)
        except PermissionError:
            QMessageBox.critical(
                self, "Save failed",
                f"Could not write to {self.excel_path.name}.\n\n"
                "The file may be open in Excel. Close it and try again."
            )
            return False
        except Exception as e:
            QMessageBox.critical(self, "Save failed", f"Could not save Excel:\n{e}")
            return False

        if not silent:
            self.status.setText(f"Saved to {self.excel_path.name}")
        return True

    # ---- Card display ----
    def _refresh_card(self):
        if not self.cards:
            self.counter.setText("")
            return
        card = self.cards[self.index]
        if self.showing_back:
            text = card["meaning"] or "(no meaning yet)"
            if card["example"]:
                text += f"\n\n— {card['example']}"
            self.card_label.setStyleSheet(
                "background:#eef7ff; border:2px solid #4a90e2; border-radius:12px; padding:30px;"
            )
        else:
            text = card["french"]
            self.card_label.setStyleSheet(
                "background:#fafafa; border:2px solid #ddd; border-radius:12px; padding:30px;"
            )
        self.card_label.setText(text)
        self.counter.setText(f"{self.index + 1} / {len(self.cards)}")

    def flip(self):
        if not self.cards:
            return
        self.showing_back = not self.showing_back
        self._refresh_card()

    def next_card(self):
        if not self.cards:
            return
        self.index = (self.index + 1) % len(self.cards)
        self.showing_back = False
        self._refresh_card()

    def prev_card(self):
        if not self.cards:
            return
        self.index = (self.index - 1) % len(self.cards)
        self.showing_back = False
        self._refresh_card()

    def _on_shuffle(self):
        if not self.cards:
            return
        import random
        if self.shuffle_cb.isChecked():
            random.shuffle(self.cards)
        else:
            self.cards.sort(key=lambda c: c["french"].lower())
        self.index = 0
        self.showing_back = False
        self._refresh_card()

    # ---- Ollama integration ----
    def fetch_with_ollama(self):
        if not self.cards:
            return
        missing_words = [c["french"] for c in self.cards if not c["meaning"]]

        if missing_words:
            target_words = missing_words
            reply = QMessageBox.question(
                self,
                "Fetch from Ollama",
                f"Fetch meanings for {len(missing_words)} missing words "
                f"from Ollama ({OLLAMA_MODEL})?"
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
        else:
            # Everything already has a meaning — offer to re-fetch all
            reply = QMessageBox.question(
                self,
                "Re-fetch all?",
                f"All {len(self.cards)} cards already have meanings.\n\n"
                f"Re-fetch all of them from Ollama? This will overwrite both "
                f"the cache and the Excel file."
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
            target_words = [c["french"] for c in self.cards]
            for w in target_words:
                self.cache.pop(w, None)
            for c in self.cards:
                c["meaning"] = ""
            self._save_cache()
            self._refresh_card()

        self.progress.setVisible(True)
        self.progress.setRange(0, len(target_words))
        self.progress.setValue(0)
        self.fetch_btn.setEnabled(False)
        self.load_btn.setEnabled(False)
        self.save_btn.setEnabled(False)

        self.worker = OllamaWorker(target_words)
        self.worker.progress.connect(self._on_progress)
        self.worker.word_done.connect(self._on_word_done)
        self.worker.finished_all.connect(self._on_fetch_done)
        self.worker.error.connect(self._on_fetch_error)
        self.worker.start()

    def _on_progress(self, current, total, word):
        self.progress.setValue(current)
        self.status.setText(f"Fetching {current}/{total}: {word}")

    def _on_word_done(self, word, meaning):
        for c in self.cards:
            if c["french"] == word:
                c["meaning"] = meaning
        self.cache[word] = meaning
        self._save_cache()
        if self.cards and self.cards[self.index]["french"] == word:
            self._refresh_card()

    def _on_fetch_done(self):
        self.progress.setVisible(False)
        self.fetch_btn.setEnabled(True)
        self.load_btn.setEnabled(True)
        self.save_btn.setEnabled(True)
        # Auto-save back to Excel
        saved = self.save_to_excel(silent=True)
        if saved:
            self.status.setText(
                f"Done. Cached + saved to {self.excel_path.name}."
            )
        else:
            self.status.setText("Done. Cache saved (Excel write failed).")

    def _on_fetch_error(self, msg):
        self.progress.setVisible(False)
        self.fetch_btn.setEnabled(True)
        self.load_btn.setEnabled(True)
        self.save_btn.setEnabled(True)
        QMessageBox.critical(self, "Ollama error", msg)

    # ---- Cache ----
    def _load_cache(self):
        if CACHE_FILE.exists():
            try:
                return json.loads(CACHE_FILE.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def _save_cache(self):
        try:
            CACHE_FILE.write_text(
                json.dumps(self.cache, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        except Exception as e:
            print(f"Cache save failed: {e}")


def main():
    app = QApplication(sys.argv)
    win = FlashcardApp()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()