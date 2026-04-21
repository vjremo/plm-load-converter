# PLM Data Loader — Excel to Tab / CSV Converter

Converts a PLM load template (`.xlsx`) into a tab-delimited (`.txt`) or comma-separated (`.csv`) file.
Two delivery options: a **browser app** (no install) and a **Python CLI**.

---

## Browser App — `convert.html`

### Requirements
- Any modern browser (Chrome, Edge, Firefox, Safari)
- Internet connection on first load (SheetJS loaded from CDN)

### Usage
1. Open `convert.html` in your browser (double-click or `File → Open`)
2. Click **Choose a .xlsx file** or drag and drop your template onto the drop zone
3. The output preview appears immediately below
4. Click **⬇ datafile.txt** to download as tab-delimited text, or **⬇ datafile.csv** to download as CSV
5. Click **✕ Reset** to clear and load a different file

Files are processed entirely in the browser. Nothing is uploaded to any server. Data is discarded when the tab is closed.

---

## Python CLI — `convert.py`

### Requirements
- Python 3.9+
- `openpyxl` library

```bash
pip install -r requirements.txt
```

### Usage

```bash
# defaults: Load-Template.xlsx → datafile.txt
python convert.py

# explicit paths
python convert.py path/to/input.xlsx path/to/output.txt
```

---

## Excel File Structure

### Sheet 1 — Data

| Column | Header example | Rule |
|--------|---------------|------|
| A | `Object` | Plain string |
| B | `Object Type` | Plain string |
| C | `Name` | Plain string |
| D | `Department-SingleList` | Single lookup |
| E | `Class-ColorChoice` | Single lookup |
| F | `Age Range-MultiList` | Multi lookup |
| G | `Style-Composite` | Multi lookup |
| H | `Short Description` | Plain string |
| I | `Target Retail Price-Float` | Float validation |
| J | `Certified-Boolean` | Boolean conversion |
| K | `Sort Order-Integer` | Integer validation |
| L | `Tags-MultiEntry` | Multi plain-text entry |

- **Row 1** — column headers (define the transformation rule via suffix)
- **Row 2+** — data rows (one output line per row)

### Sheet 2 — References

Lookup table mapping **display names** to **internal names** for list fields.

| Department-Internal Name | Department-Display Name | Class-Internal Name | Class-Display Name | … |
|--------------------------|-------------------------|---------------------|--------------------|---|
| xyzDept1 | Dept 1 | xyzClass1 | Class 1 | … |
| xyzDept2 | Dept 2 | xyzClass2 | Class 2 | … |

Headers must follow the pattern `<Category>-Internal Name` / `<Category>-Display Name`.
The category must match the prefix before the suffix in Sheet 1 headers
(e.g. `Department-SingleList` matches `Department-Internal Name`).

---

## Column Suffix Rules

| Suffix | Canonical rule | Behaviour |
|--------|---------------|-----------|
| `-SingleList` | SingleList | Looks up the display name in References → outputs internal name. Error if not found. |
| `-ColorChoice` | SingleList | Same as SingleList. |
| `-MultiList` | MultiList | Comma-separated display names → each looked up → joined with `\|~*~\|`. Error if any value not found. |
| `-Composite` | MultiList | Same as MultiList. |
| `-Float` | Float | Validates the value is a decimal number. Error if not. |
| `-Boolean` | Boolean | `Yes` → `true`, `No` → `false`. Error on any other value. |
| `-Integer` | Integer | Must be a whole number with no decimal point. Error otherwise. |
| `-MultiEntry` | MultiEntry | Comma-separated plain text → joined with `\|~*~\|`. No reference lookup. |
| *(no suffix)* | Plain | Output as-is. |

---

## Output Format

Tab-delimited, one line per data row, no header row, no row numbers.

```
Product	Product\Type1	Test Product	xyzDept1	xyzClass1	xyz6plus	Test Descrition	3.05	true
```

**MultiList / Composite example** — input cell `3 to 6 years,6 plus` becomes:

```
xyz4to6|~*~|xyz6plus
```

**MultiEntry example** — input cell `Test1, Test2` becomes:

```
Test1|~*~|Test2
```

---

## Error Handling

Errors include the row number and column name for easy location.

| Scenario | Error message example |
|----------|-----------------------|
| List value not in References | `Row 2, column "Department-SingleList": "Unknown Dept" not found in References for "Department". Valid: Dept 1, Dept 2, Dept 3` |
| Invalid float | `Row 3, column "Target Retail Price-Float": "abc" is not a valid floating-point number` |
| Invalid boolean | `Row 4, column "Certified-Boolean": Boolean must be "Yes" or "No", got "maybe"` |
| Invalid integer | `Row 5, column "Sort Order-Integer": "3.5" is not a valid whole number (no decimals allowed)` |

In the browser app, all errors are shown in a red panel beneath the drop zone and rows with errors are skipped in the output. In the CLI, errors are raised and the process exits with a non-zero code.

---

## File Overview

| File | Purpose |
|------|---------|
| `convert.html` | Browser app — open directly, no install required |
| `convert.py` | Python CLI — requires `openpyxl` |
| `requirements.txt` | Python dependencies |
| `Load-Template.xlsx` | Example input file |
| `datafile.txt` | Example output file |
