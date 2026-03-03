# 🛡️ AI-Powered Client Redactor

A high-performance tool that uses **Deep Learning (Transformer-based NLP)** to identify and redact client/company names across PowerPoint (`.pptx`) and Excel (`.xlsx`) files. It ensures that the same client is consistently renamed (e.g., `[client1]`) across all your documents.

## ✨ Key Features

* **AI Entity Detection:** Uses spaCy's `en_core_web_trf` (Transformer) model for high-accuracy organization detection.
* **Smart Normalization:** Automatically handles variations like "Acme Corp," "Acme Inc," and "Acme" as the same entity.
* **Multi-Format Support:** Processes both Excel and PowerPoint files in a single pass.
* **Persistent Mapping:** Remembers client IDs across sessions via `client_mapping.json`.

---

## ⚙️ Setup Instructions (Windows)

Follow these steps to get your environment ready:

### 1. Create a Virtual Environment

Open **PowerShell** or **Command Prompt** in your project folder:

```powershell
# Create the environment
python -m venv venv

# Activate it
.\venv\Scripts\activate

```

### 2. Install Dependencies

```powershell
pip install -r requirements.txt

```

### 3. Download the AI Model

The script requires the heavy-duty Transformer model to work:

```powershell
python -m spacy download en_core_web_trf

```

---

## 🚀 How to Run

1. Place your files in the `./source_data` folder.
2. Open `main.py` to adjust your configuration (see below).
3. Run the script:
```powershell
python main.py

```


4. Find your redacted files in the `./sanitized_data` folder.

---

## 🛠️ Configuration Guide

Inside `main.py`, you will find the `GLOBAL_CLIENT_LIST`. This determines how the AI behaves:

### Mode A: Auto-Discovery (Recommended for broad cleaning)

**Set to:** `GLOBAL_CLIENT_LIST = []`

* **Why:** If the list is empty, the AI will redact **every** organization it finds (Microsoft, Google, local vendors, etc.).
* **Best for:** Cleaning a deck entirely of all third-party references.

### Mode B: Targeted Redaction

**Set to:** `GLOBAL_CLIENT_LIST = ["Acme", "Globex"]`

* **Why:** The AI will only look for these specific clients. It will ignore "Microsoft" or "Google" if they are not in your list.
* **Best for:** When you only need to hide specific high-value client names while keeping general tech names visible.

---

## 🧠 System Logic

1. **Extract:** The script "heals" text from PowerPoint runs or Excel cells.
2. **Detect:** The Transformer model identifies "ORG" (Organization) entities.
3. **Fuzzy Match:** It checks if "Acme" on Slide 1 matches "Acme Corp" in an Excel cell.
4. **Replace:** It swaps the text for a unique ID and saves a new version of the file, preserving as much formatting as possible.
