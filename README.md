# MKTG 411 Knowledge Checks

This repository contains print-ready versions of the three MKTG 411 knowledge
checks copied from the existing Canvas New Quizzes. The original Canvas quiz
content was not modified.

## Print-ready files

The `print-ready` folder contains a student copy and a separate instructor
answer key for each knowledge check.

## Source and rebuilding

- `quiz_content.json` contains the questions, response choices, and keyed
  answers transcribed from Canvas.
- `build_quizzes.py` generates the six Word documents.

Run the builder with Python and `python-docx` available:

```powershell
python .\build_quizzes.py
```
