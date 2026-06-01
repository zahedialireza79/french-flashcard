"""
French Flashcards - PyQt6 app
- Loads French words from an Excel file.
- Expected columns (case-insensitive): 'french' (required), 'meaning' (optional), 'example' (optional).
- If 'meaning' is missing/empty for a row, you can fetch it from a local Ollama server.
- Meanings are cached to meanings_cache.json so you only pay the Ollama cost once.

Run:
    pip install PyQt6 pandas openpyxl requests
    python flashcards.py
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
OLLAMA_MODEL = "gemma3"  # change to whatever you have pulled, e.g. "mistral", "llama3.1"
CACHE_FILE = Path("meanings_cache.json")


# ---------- Ollama worker (runs in a background thread so UI stays responsive) ----------

class OllamaWorker(QThread):
    """Fetches meanings for a list of French words from a local Ollama server."""
    progress = pyqtSignal(int, int, str)        # current, total, current word
    word_done = pyqtSignal(str, str)            # french word, meaning
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
                prompt = (
                    f"Give the English meaning of the French word '{word}'. "
                    f"Respond with ONLY the English translation, no extra explanation, "
                    f"no quotes, no French. If the word has multiple common meanings, "
                    f"give the 2 or 3 most common, separated by commas."
                )
                resp = requests.post(
                    OLLAMA_URL,
                    json={"model": self.model, "prompt": prompt, "stream": False},
                    timeout=60,
                )
                resp.raise_for_status()
                meaning = resp.json().get("response", "").strip()
                # Clean up common LLM artifacts
                meaning = meaning.strip(' "\'.\n')
                self.word_done.emit(word, meaning or "(no answer)")
            except requests.exceptions.ConnectionError:
                self.error.emit(
                    "Could not reach Ollama at http://localhost:11434.\n"
                    "Make sure Ollama is running ('ollama serve') and the model is pulled "
                    f"('ollama pull {self.model}')."
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
        self.resize(640, 480)

        self.cards = []          # list of dicts: {french, meaning, example}
        self.index = 0
        self.showing_back = False
        self.cache = self._load_cache()
        self.worker = None

        self._build_ui()
        self._refresh_card()

    # ---- UI ----
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(15)

        # Top bar: load + ollama
        top = QHBoxLayout()
        self.load_btn = QPushButton("Load Excel…")
        self.load_btn.clicked.connect(self.load_excel)
        top.addWidget(self.load_btn)

        self.fetch_btn = QPushButton("Fetch missing meanings from Ollama")
        self.fetch_btn.clicked.connect(self.fetch_with_ollama)
        self.fetch_btn.setEnabled(False)
        top.addWidget(self.fetch_btn)

        self.shuffle_cb = QCheckBox("Shuffle")
        self.shuffle_cb.stateChanged.connect(self._on_shuffle)
        top.addWidget(self.shuffle_cb)
        top.addStretch()
        root.addLayout(top)

        # Card area (clickable label)
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

        # Nav buttons
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

        # Progress bar (hidden until needed)
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        self.status = QLabel("")
        root.addWidget(self.status)

    # ---- Data loading ----
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

        # Normalize column names
        df.columns = [str(c).strip().lower() for c in df.columns]
        if "french" not in df.columns:
            # Fall back: use first column as the French word
            df = df.rename(columns={df.columns[0]: "french"})

        self.cards = []
        for _, row in df.iterrows():
            word = str(row.get("french", "")).strip()
            if not word or word.lower() == "nan":
                continue
            meaning = str(row.get("meaning", "")).strip() if "meaning" in df.columns else ""
            if meaning.lower() == "nan":
                meaning = ""
            # Apply cached meaning if Excel has none
            if not meaning and word in self.cache:
                meaning = self.cache[word]
            example = str(row.get("example", "")).strip() if "example" in df.columns else ""
            if example.lower() == "nan":
                example = ""
            self.cards.append({"french": word, "meaning": meaning, "example": example})

        if not self.cards:
            QMessageBox.warning(self, "Empty", "No valid French words found.")
            return

        self.index = 0
        self.showing_back = False
        missing = sum(1 for c in self.cards if not c["meaning"])
        self.fetch_btn.setEnabled(missing > 0)
        self.status.setText(
            f"Loaded {len(self.cards)} cards" + (f" — {missing} missing meanings" if missing else "")
        )
        self._refresh_card()

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
        missing = [c["french"] for c in self.cards if not c["meaning"]]
        if not missing:
            QMessageBox.information(self, "Done", "All cards already have meanings.")
            return

        reply = QMessageBox.question(
            self,
            "Fetch from Ollama",
            f"Fetch meanings for {len(missing)} words from Ollama ({OLLAMA_MODEL})?\n"
            f"This may take a while (~1–3s per word).",
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        self.progress.setVisible(True)
        self.progress.setRange(0, len(missing))
        self.progress.setValue(0)
        self.fetch_btn.setEnabled(False)
        self.load_btn.setEnabled(False)

        self.worker = OllamaWorker(missing)
        self.worker.progress.connect(self._on_progress)
        self.worker.word_done.connect(self._on_word_done)
        self.worker.finished_all.connect(self._on_fetch_done)
        self.worker.error.connect(self._on_fetch_error)
        self.worker.start()

    def _on_progress(self, current, total, word):
        self.progress.setValue(current)
        self.status.setText(f"Fetching {current}/{total}: {word}")

    def _on_word_done(self, word, meaning):
        # Update both the cards and the cache
        for c in self.cards:
            if c["french"] == word:
                c["meaning"] = meaning
        self.cache[word] = meaning
        self._save_cache()
        # If the user is looking at this card, refresh
        if self.cards and self.cards[self.index]["french"] == word:
            self._refresh_card()

    def _on_fetch_done(self):
        self.progress.setVisible(False)
        self.fetch_btn.setEnabled(False)
        self.load_btn.setEnabled(True)
        self.status.setText("Done. Meanings cached to meanings_cache.json.")

    def _on_fetch_error(self, msg):
        self.progress.setVisible(False)
        self.fetch_btn.setEnabled(True)
        self.load_btn.setEnabled(True)
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
