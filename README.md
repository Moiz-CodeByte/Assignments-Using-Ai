# Assignments Using AI

This repository helps students quickly create clean, professional assignment files in Word format by using Python scripts.

## Why This Is Useful for Students

- Saves time by generating a fully formatted assignment in seconds.
- Keeps layout consistent (headings, spacing, references, title page).
- Makes it easy to create multiple assignments with the same professional style.
- Reduces formatting stress so you can focus on content quality.
- Helps beginners learn simple automation with## Project StructureS documents)
- `Resources/` : Notes, source material, and reference content
- `Scripts/` : Assignment generation scripts organiziles (`.docx`)

Inside `Scripts/`:

- `Principal of Psycology/`
  - `create_assignment4.py`
  - `create_assignment5.py`
- `Professional Practices/`
  - `create_assignment.py`
  - `create_assignment2.py`
  - `create_assignment3.py`

## Requirements

- Python 3.10+
- Package: `python-docx`

Install dependency:

```powershell
py -m pip install python-docx
```

## How to Use

Run any script from the project root:

```powershell
py "Scripts\Principal of Psycology\create_assignment4.py"
py "Scripts\Principal of Psycology\create_assignment5.py"
py "Scripts\Professional Practices\create_assignment.py"
py "Scripts\Professional Practices\create_assignment2.py"
py "Scripts\Professional Practices\create_assignment3.py"
```

Each script creates a `.docx` assignment file in `Word Files/` (or the output path defined in that script).

## Important: Update Scripts According to Your Topic

Before running a script for your class submission, edit it based on your own topic.

Update these parts in the selected script:

- Main title text (assignment topic)
- Subject/course line
- Introduction paragraph
- Main sections and headings
- Bullet points or key arguments
- Conclusion
- References
- Output file name and path (`out = ...`)

Example workflow:

1. Copy an existing script (for example `create_assignment5.py`) and rename it for your topic.
2. Replace all content sections with your own topic details.
3. Keep the formatting helper functions (`h`, `body`) for consistent design.
4. Change the output filename so your new file is easy to identify.
5. Run the script and check the generated document in `Word Files/`.

## Best Practices for Students

- Use clear academic headings and short paragraphs.
- Add at least 3 to 5 credible references.
- Verify spellings in headings and file names before submission.
- Do not submit without reviewing the generated Word file.
- Keep one script per topic so future edits are easier.

## Troubleshooting

- If `python` is not recognized on Windows, use `py`.
- If you get `FileNotFoundError`, ensure the output folder exists and path is correct.
- If module error appears for `docx`, install dependency again:

```powershell
py -m pip install python-docx
```
recreate new assignments with a consistent design.
