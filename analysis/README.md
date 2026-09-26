# Analysis scripts

- `stats_core.py` — loads the de-identified cohort (81 ankles / 80 patients) from the study Excel file and provides the statistical helpers.
- `build_doc.py` — computes every number and table and writes the manuscript draft (.docx).

The Excel file contains patient identifiers and is **not** stored in this repository. Point the scripts to it with:

```bash
pip install openpyxl pandas scipy python-docx
ARTHRO_XLSX=/path/to/update_ver2.5.xlsx python3 build_doc.py ../manuscript/output.docx
```
