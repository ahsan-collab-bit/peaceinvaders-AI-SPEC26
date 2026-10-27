# PeaceInvaders-AI-Spec26 — CNC Part Analyzer

**Team Peace Invaders** — Ahsan (team lead), Basit, Ezan · SPEC'26 AI domain · NED University

**Level reached: 3 (Excel + PDF drawing + IGES 3D model, cross-checked) + stretch (any part number).**
**LLM used: OpenAI gpt-oss-120b via the Groq API (temperature 0).**

- Technical document: [Technical_Document.pdf](Technical_Document.pdf)
- Presentation: uploaded through the SPEC'26 submission form

## Screenshots

Cross-check for part 2280 (all three files agree):

![check 2280](screenshots/check_2280.png)

Cross-check for part 2283 (planted material mismatch is caught):

![check 2283](screenshots/check_2283.png)

Sample engineering drawing from the dataset (part 2280):

![drawing 2280](screenshots/drawing_2280.png)

## Repository layout

| Path | What it is |
|---|---|
| `cnc_app.py` | The whole app (one file) |
| `data/` | Sample dataset: Excel record, 5 PDF drawings, 5 IGES models, `SOURCES.txt` |
| `screenshots/` | App output and a drawing preview |
| `Technical_Document.pdf` | Architecture, how each file is read, how the LLM is used, limitations |

Type a part number and a question. The app finds that part's Excel manufacturing record,
PDF engineering drawing and IGES 3D model, cross-checks them, and asks an LLM to answer
using ONLY those facts. Every answer names the file (and sheet/row/page) it came from.
If the answer is not in the files, it says "I can't find that in the files."

## How to run (Windows)
1. Install Python 3.10+ from python.org.
2. Install the two libraries:   py -m pip install openpyxl pypdf
3. Put the Excel (.xlsx or .csv), PDF drawings (.pdf) and IGES models (.igs/.iges) in the `data` folder.
4. Run:   py cnc_app.py
5. First time only: paste a free Groq API key (console.groq.com -> API Keys). It is saved in llm_key.txt.
6. Type a part number (e.g. 2280), then ask questions.
   Commands: `facts` = show all evidence, `new` = another part, `quit` = exit.

Other commands:
- py cnc_app.py --check 2280        show files found, facts and cross-check (no AI needed)
- py cnc_app.py --data D:\files    use files from another folder (e.g. new files from the judges)

## Sample dataset (data folder)
5 parts (2280-2284): one Excel record (2 sheets), 5 PDF drawings, 5 IGES models. See data/SOURCES.txt.
Part 2283 has a deliberate material mismatch; part 2282 has no surface finish in Excel (honesty test).

## How it works
1. Excel (openpyxl): finds the column-title row automatically, then every row on every sheet
   that contains the part number (exact match first, then forms like P-2280). Keeps sheet + row.
2. Finding the PDF/IGES for a part: file name contains the part number -> file named in the
   Excel row -> part number inside the file -> only file of that type (shown as a warning).
3. PDF (pypdf): text lines and numbers from every page; title-block lines are marked.
4. IGES (own parser, no CAD software needed): units from the Global section; real edges
   (lines, arcs, B-spline curves) with transformation matrices applied give the overall size;
   full circles give hole/round diameters.
5. Cross-check (plain Python, not the AI): part number on drawing, material Excel vs drawing,
   Excel sizes vs drawing, 3D size vs drawing, Excel sizes vs 3D size, circle diameters vs drawing.
   Results: OK / MISMATCH / CHECK.
6. LLM: Groq API (OpenAI-compatible); the model is chosen automatically and printed at start.
   Temperature 0. The LLM only receives the extracted facts with their sources plus strict
   rules (answer only from facts, cite sources, say when something is missing).
   If the AI cannot be reached, an offline mode shows the matching facts instead.

## Limitations
- Size is the bounding box along the model's X/Y/Z axes (a rotated part gives a bigger box).
- Circles are reported as "circle edges", not guaranteed holes (could be round outer faces).
- Old IGES types (e.g. 112/114 parametric splines) are not measured; the app says so.
- Scanned PDFs have no text; facts can be typed into <pdf name>_notes.txt as a workaround.
- Drawing numbers are matched without knowing which dimension line they belong to.
- .xls (old Excel) is not supported: save as .xlsx or .csv.
- The LLM can still misread a fact, so the app always shows the evidence it was given.
