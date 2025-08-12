 NDJsonToXlsx

Convert **NDJSON** (newline-delimited JSON; one JSON object per line) into **Excel `.xlsx`**.  
Nested objects/arrays are **flattened** into columns.

---

## Features
- Streams input line-by-line — suitable for large files.
- Two passes: **1st** collects all column keys, **2nd** writes the workbook.
- Object/array **flattening**:
  - Objects → `a.b.c`
  - Arrays → `arr[0]`, `arr[1]`, …
- Unsupported Excel types (lists/dicts, etc.) are saved as **JSON strings**.
- Malformed JSON lines are **skipped** and counted.

---

## Requirements
- Python **3.8+**
- Package: `xlsxwriter`

```bash
pip install xlsxwriter
```

---

## Installation
Clone the repo (or copy `converter.py` into your project):

```bash
git clone https://github.com/<your-username>/ndjson-to-xlsx.git
cd ndjson-to-xlsx
pip install xlsxwriter
```

---

## Usage

```bash
python converter.py input.ndjson output.xlsx
```

> `output.xlsx` must have the `.xlsx` extension.

---

## Example

**Input (`input.ndjson`)**
```ndjson
{"id":1,"user":{"name":"Ana","tags":["vip","beta"]},"score":9.5}
{"id":2,"user":{"name":"João","tags":[]},"score":7}
{"id":3,"score":null,"extra":{"a":1,"b":{"c":2}}}
```

**Output columns (`output.xlsx`)**
```
id | user.name | user.tags[0] | user.tags[1] | score | extra.a | extra.b.c
```

Notes:
- Arrays are expanded by index across all lines that contain them.
- Empty arrays do **not** create columns.
- Any remaining list/dict value is written as a **JSON text** (e.g., `"[1,2,3]"`, `"{}"`).

---

## Limits & Behavior
- **Excel row limit:** 1,048,576. Extra lines are ignored with a warning.
- **Excel column limit:** 16,384 (XFD). Excess unique fields will be ignored.
- **Types**
  - `str`, `int`, `float`, `bool`, `None` → written natively
  - `date/datetime/time` → ISO-8601 text (`.isoformat()`)
  - `list` / `dict` → JSON text
- **Invalid lines** (non-JSON) are skipped and reported at the end.

---

## Troubleshooting
- **`TypeError: Unsupported type <class 'list'> in write()`**  
  The writer can’t place lists/dicts directly in cells. This tool converts such values to **JSON strings** before writing. If you modified the script, restore that conversion step.

---

## Performance
- First pass scans to discover the union of keys (columns).
- Second pass streams rows to the worksheet.
- Highly heterogeneous NDJSON (many unique keys) will produce many columns and may hit Excel’s column limit.

---

## Development
Ideas and improvements welcome:
- Include/exclude column filters
- Partitioning to multiple sheets based on a key (e.g., a `type` field)
- Schema validation / typing hints

---

## Contributing
1. Fork the repository  
2. Create a feature branch: `git checkout -b feat/my-feature`  
3. Commit: `git commit -m "feat: add my feature"`  
4. Push: `git push origin feat/my-feature`  
5. Open a Pull Request