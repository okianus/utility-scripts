# utility-scripts

A collection of small utility scripts and tools. Each subfolder is self-contained with its own setup and usage.

## Scripts

| Script | Description |
|--------|-------------|
| [pulse-survey-analysis](pulse-survey-analysis/) | Visualizes engagement pulse survey results: sentiment breakdown and horizontal bar charts. |

## General notes

- **Per-script setup:** Open the script’s folder and follow its README (e.g. venv, `pip install -r requirements.txt`).
- **Python:** Scripts may use their own virtual environment (e.g. `pulse-survey-analysis/.venv`). Run with that project’s Python or activate its venv before running.
- **Secrets:** Don’t commit API keys or sensitive data. Use environment variables or local `.env` files (and keep `.env` in `.gitignore`).
