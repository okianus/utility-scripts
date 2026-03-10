#!/usr/bin/env bash
# Run the survey script with the project's venv (so openai and other deps are available).
cd "$(dirname "$0")"
.venv/bin/python visualize_survey.py
