# excelcompare

A command-line tool for API regression testing: it compares expected values maintained in an
Excel workbook against the actual values returned in an XML response, and generates a
Markdown report.

Built for a QA specialist verifying a financial-sector API. The manual alternative was
reading an XML response field by field against a spreadsheet — slow, and easy to miss a
single changed digit.

---

## What it does

1. Loads expected values from an `.xlsb` workbook — one row per field.
2. Loads an XML response and extracts the payload embedded inside it.
3. Matches every expected key against the response, section by section.
4. Compares values, normalising numeric formats and translating Excel error codes.
5. Prints a colour-coded result per field and writes a timestamped Markdown report.

Each field ends up in one of three states:

| | Meaning |
| :-- | :-- |
| ✅ | Value found in the response and identical to the expected one |
| ❌ | Value found but different |
| ⚠️ | Key not present in the response at all |

---

## Requirements

- Python 3.8+
- [`pyxlsb`](https://pypi.org/project/pyxlsb/) — reads the binary `.xlsb` format

```bash
pip install -r requirements.txt
```

---

## Project layout

```
excelcompare/
├─ excelcompare.py
├─ requirements.txt
├─ <workbook>.xlsb        # expected values — one workbook, next to the script
├─ xml/                   # XML responses to check
└─ reports/               # generated reports (created automatically)
```

---

## Input format

### Excel workbook

A single `.xlsb` file in the script directory. Temporary lock files (`~$…`) are ignored, so
you can keep the workbook open while running the tool.

The tool reads a sheet named `OUTBOUND`, using two columns:

| Column | Content |
| :-- | :-- |
| **A** | Field key in dot notation — `root.Section.FieldName` |
| **B** | Expected value; multiple values separated by `\|` |

Rules:

- **Dot notation drives the lookup.** The element before the last one is treated as the XML
  section, so `outbound.Summary.Total` is looked for inside `<Summary>`. A two-part key such
  as `outbound.Status` is searched at the top level of the payload.
- **A pipe separates repeated elements.** `A10 | A20` matches a section containing `<Code>`
  twice.
- **An empty cell in column B means the field is expected to be absent.** It is rendered as
  🚫 in the report, and any value found in the response is flagged as a difference.
- **Excel error codes are translated, not compared raw.** A cell holding `#VALUE!`,
  `#DIV/0!`, `#N/A` or `#NAME?` is shown as a readable label, so a broken formula in the
  expectations sheet is visible instead of silently failing the comparison.
- **Numbers are normalised.** Numeric strings become numbers and integral floats become
  integers, so `60` and `60.0` compare equal.

### XML response

Any number of `.xml` files in `xml/`. The tool handles responses that carry the actual
payload as an escaped string inside the envelope — it extracts and unescapes that payload
before parsing, so you can save the raw response exactly as received. Sections present but
empty in the response are recognised as such rather than reported as missing.

---

## Worked example

The data below is invented — it only shows the shape the tool expects.

**Expectations sheet (`OUTBOUND`):**

| A — key | B — expected value |
| :-- | :-- |
| `outbound.RequestId` | `100015` |
| `outbound.Status` | `OK` |
| `outbound.Items.Code` | `A10 \| A20` |
| `outbound.Items.Label` | *(empty)* |
| `outbound.Summary.Total` | `1500` |
| `outbound.Summary.Ratio` | `#VALUE!` |

**XML response (`xml/my_test_case.xml`):**

```xml
<m:xmlString>&lt;outbound>
  &lt;RequestId>100015&lt;/RequestId>
  &lt;Status>PENDING&lt;/Status>
  &lt;Items>
    &lt;Code>A10&lt;/Code>
    &lt;Code>A20&lt;/Code>
    &lt;Label>COS&lt;/Label>
    &lt;Label>-9&lt;/Label>
  &lt;/Items>
  &lt;Summary>
    &lt;Total>1500&lt;/Total>
  &lt;/Summary>
&lt;/outbound></m:xmlString>
```

**Resulting report:**

```markdown
## API REGRESSION TEST REPORT
+ Test case: my_test_case
+ Generated: 2026-08-05 10:13:01

#### Field comparison

| 🎭 | KEY           | EXCEL                | XML             |
| :- | :------------ | :------------------- | :-------------- |
| ✅ | RequestId     | 100015               | 100015          |
| ❌ | Status        | OK                   | PENDING         |
| ✅ | Items.Code    | A10  \|  A20         | A10  \|  A20    |
| ❌ | Items.Label   | 🚫                   | COS  \|  -9     |
| ✅ | Summary.Total | 1500                 | 1500            |
| ⚠️ | Summary.Ratio | ❗SYNTAX ERROR (#VALUE!)❗ | not found in xml |
```

Reading the interesting rows:

- **`Items.Code`** — matched inside `<Items>`, both repeated values in order.
- **`Items.Label`** — expected to be absent (empty cell → 🚫), but the response returned two
  values. Flagged as a difference rather than passed over.
- **`Summary.Ratio`** — the expectations sheet holds a broken formula, and the field is also
  missing from the response. Both facts are visible in one row.

Because the report is Markdown, it pastes directly into a ticket or a pull request.

---

## Usage

Set the test case at the top of `excelcompare.py`:

```python
XML_FILE_NAME = 'my_test_case'   # the .xml extension is optional
```

Then run:

```bash
python excelcompare.py
```

The tool validates its inputs before doing any work and stops with an explicit message if
the `xml/` folder is missing, if there is not exactly one workbook, or if the named XML file
is not there. Console output lists every field with its expected and actual values, followed
by a summary of keys missing from the response.

---

## Sample data

**No sample workbook or XML response is included in this repository.** The tool was built
for a commissioned project, and both the field names and the message structure belong to
that engagement — publishing them is not something the project allows.

The input format is documented above and is fully reflected in the source code, which is
enough to adapt the tool to your own data: put one `.xlsb` with an `OUTBOUND` sheet next to
the script, drop an XML response into `xml/`, and run it.

---

## Limitations

- Expects exactly one `.xlsb` workbook in the script directory.
- The sheet name and the response envelope structure are fixed in code.
- The first row of the sheet is treated as a header and skipped.
- Test case selection is edited in the source rather than passed as an argument.

## Roadmap

- Command-line arguments for the test case and paths
- Configurable sheet name and envelope handling
- Unit tests with pytest
- Batch mode — run every XML in the folder in one pass

---

Built by [Tomasz Litwicki](https://github.com/TomaszLitwicki)
