# Pulse Survey Analysis

## Setup (use this so you have OpenAI and other deps)

1. **Create the virtual environment** (once):
   ```bash
   cd pulse-survey-analysis
   python3 -m venv .venv
   ```

2. **Install dependencies from requirements.txt** (once, or after changing requirements.txt):
   ```bash
   .venv/bin/pip install -r requirements.txt
   ```
   This installs everything listed in `requirements.txt`, including `openai`.

3. **Use the project’s Python when you run the script**  
   The ImportError usually happens because the script is run with a different Python (e.g. system or another venv) that doesn’t have these packages.

   - **From terminal:** run with the venv’s Python:
     ```bash
     .venv/bin/python visualize_survey.py
     ```
     Or activate the venv first, then run:
     ```bash
     source .venv/bin/activate
     pip install -r requirements.txt   # if you ever need to reinstall
     python visualize_survey.py
     ```
   - **From Cursor/VS Code:** choose the `.venv` interpreter so “Run” uses it:
     - Command Palette → “Python: Select Interpreter”
     - Pick the one under `pulse-survey-analysis/.venv/bin/python` (e.g. `Python 3.9.x ('.venv': venv)`).

## Run

```bash
.venv/bin/python visualize_survey.py
```

Or use the run script:

```bash
./run.sh
```

Optional: set `OPENAI_API_KEY` in your environment for LLM-generated chart labels.
