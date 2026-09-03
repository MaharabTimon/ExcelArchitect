# ExcelArchitect — AI Spreadsheet Co-Pilot

A desktop spreadsheet manipulation co-pilot application built with PyQt6, pandas, openpyxl, and Groq LLMs.

## Features
- **Natural Language Execution:** Type plain English instructions (e.g. *"Calculate 15% VAT in column G and generate a pie chart"*) to automatically update workbooks.
- **Sandboxed Execution:** Executes generated code in an isolated execution scope with state snapshotting.
- **Undo / Redo Stack:** Full multi-step undo/redo support.
- **Matplotlib & OpenPyXL Charts:** Automated chart creation and styling.

## Setup & Usage
1. Install dependencies:
   ```bash
   pip install PyQt6 pandas openpyxl matplotlib groq python-dotenv
   ```
2. Set your Groq API key in `.env` (see `.env.example`).
3. Run the application:
   ```bash
   python app.py
   ```
