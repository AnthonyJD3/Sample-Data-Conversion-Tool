# Sample Data Conversion Tool

A Python-based data migration utility that converts legacy sample-tube box exports from a grid-style Excel layout into the row-based format required for import into a newer database.

The project was created to support a database migration in which historical sample data was stored in multiple physical box layouts and with different numbers of data values associated with each sample position. Rather than manually restructuring thousands of records, the scripts automate the conversion while preserving each sample's box and position information.

## Overview

The legacy database exported sample information in a layout that mirrored the physical organization of sample-tube boxes. Depending on the source data:

- Boxes contain either **9 × 9 positions (81 samples)** or **10 × 10 positions (100 samples)**.
- Each sample position contains **4, 5, 6, 7, or 8 rows of associated data**.
- Multiple boxes are included in a single worksheet.
- The current templates and scripts are structured around **13 boxes per conversion**.

The new database requires the information in a normalized table where each sample position occupies a single row. The output contains the sample's data values followed by dedicated **Box** and **Position** columns.

For example, an 8-value sample becomes:

| A | B | C | D | E | F | G | H | Box | Position |
|---|---|---|---|---|---|---|---|---|---|
| Sample value 1 | Sample value 2 | Sample value 3 | Sample value 4 | Sample value 5 | Sample value 6 | Sample value 7 | Sample value 8 | Box ID | Position # |

A 4-value sample uses columns `A` through `D`, while the 5-, 6-, 7-, and 8-value versions use the corresponding number of data columns.

## Why Multiple Scripts Are Used

The source files are not all structured the same way. Two characteristics determine which conversion script is required:

1. **Physical box size**
   - `9x9` — 81 tube positions
   - `10x10` — 100 tube positions

2. **Number of data values stored for each sample position**
   - 4 values
   - 5 values
   - 6 values
   - 7 values
   - 8 values

Each combination uses a matching script and Excel template. Existing script names follow the pattern:

```text
<number_of_values>_Cells_<box_size>.py
```

Examples:

```text
4_Cells_9x9.py
4_Cells_10x10.py
8_Cells_9x9.py
8_Cells_10x10.py
```

> **Note:** In the existing filenames, `Cells` refers to the number of data values/rows associated with each sample position, not the number of physical positions in the sample box.

## Template Structure

Before running a conversion, the exported legacy data must be placed into the matching template in `Old_Sample_File.xlsx`.

The conversion scripts read the workbook's **first worksheet (`Sheet1`)**, so the appropriate template layout must be used there.

The templates do not alter the actual sample data. They provide structural markers that allow the scripts to separate one box from the next.

### Box marker rows

Before the data for each box, the template contains a row in which every physical box column contains a label such as:

```text
Box D1
```

The next box begins with another marker row, for example:

```text
Box D2
```

These labels are carried with the corresponding sample records so that the converted output retains box identification.

### End marker row

After the final box, the template contains a row with:

```text
End of column
```

repeated across the physical box columns.

This final row is important to the current parsing logic. The scripts divide each source column into fixed-size box blocks and intentionally ignore the final trailing block. In other words, the current scripts do **not** dynamically search for the text `End of column`; the row acts as the structural final block that is skipped by the conversion logic.

Do not remove the box marker rows or the final `End of column` row.

## How the Conversion Works

At a high level, each script performs the following process:

1. Reads `Old_Sample_File.xlsx` with no header row.
2. Loads each physical box column into a Python list.
3. Splits each column into fixed-size blocks representing individual boxes.
4. Removes the box marker row from the sample values while retaining the marker as the box identifier.
5. Groups the correct number of values together for each sample position.
6. Combines the converted data from all physical columns into a pandas `DataFrame`.
7. Creates a `Position` column and assigns the correct physical position number for each record.
8. Writes the converted data to `New_Data_File.xlsx`.
9. Creates a new timestamped worksheet so previous conversion results are not overwritten.

Output worksheets use names similar to:

```text
NewDataPositions [date + time]
```

## Expected Output Size

Because the current scripts are configured for 13 boxes:

| Box type | Positions per box | Number of boxes | Output records |
|---|---:|---:|---:|
| 9 × 9 | 81 | 13 | 1,053 |
| 10 × 10 | 100 | 13 | 1,300 |

The number of data columns changes according to the selected script, but the number of physical sample positions is determined by the box size.

## Requirements

- Python 3
- pandas
- openpyxl
- Microsoft Excel or another application capable of viewing `.xlsx` files

Install the Python dependencies with:

```bash
pip install pandas openpyxl
```

## File Paths

The current scripts use Windows-style paths:

```python
Old_Path = "C:\\Folder\\Old_Sample_File.xlsx"
New_Path = "C:\\Folder\\New_Data_File.xlsx"
```

Either:

- place the Excel files in `C:\Folder`, or
- edit `Old_Path` and `New_Path` in the selected script so they point to the correct files on your computer.

`New_Data_File.xlsx` must already exist because the scripts open it in append mode and add a new worksheet rather than creating a new workbook from scratch.

## Usage

### 1. Determine the source format

Identify:

- whether the physical sample box is **9 × 9** or **10 × 10**, and
- whether each sample position contains **4, 5, 6, 7, or 8 data values**.

### 2. Prepare `Old_Sample_File.xlsx`

Use the matching template and place the legacy export into `Sheet1`.

Preserve:

- each `Box <BoxName>` marker row,
- the exact row structure required by that template, and
- the final `End of column` row.

Because the scripts use fixed block sizes, adding or deleting rows within the template can shift the data and produce an incorrect conversion.

### 3. Select the matching script

For example, a 10 × 10 box with 8 data values per sample uses:

```text
8_Cells_10x10.py
```

A 9 × 9 box with 4 data values per sample uses:

```text
4_Cells_9x9.py
```

### 4. Verify the paths

Update the input and output file paths at the top of the script if necessary.

### 5. Run the script

Example:

```bash
python 8_Cells_10x10.py
```

### 6. Review the converted workbook

Open `New_Data_File.xlsx`.

A new worksheet will have been appended with a timestamp in its name. Review the `Box` and `Position` columns and verify that the expected number of records was generated before importing the data into the new database.

## Example Transformation

### Legacy layout

The old format groups samples according to their physical location in a tube box. A sample's multiple data values are stored vertically, and positions are distributed across the box's physical columns.

Conceptually:

```text
Box D1          Box D1          Box D1
Position 1 (1)  Position 2 (1)  Position 3 (1)
Position 1 (2)  Position 2 (2)  Position 3 (2)
Position 1 (3)  Position 2 (3)  Position 3 (3)
...
```

### Converted layout

The conversion rotates/restructures those values so that one physical position becomes one database-ready record:

```text
A              B              C              ...   Box   Position
Position 1 (1) Position 1 (2) Position 1 (3) ...   D1    1
Position 2 (1) Position 2 (2) Position 2 (3) ...   D1    2
Position 3 (1) Position 3 (2) Position 3 (3) ...   D1    3
```

The resulting format is easier to sort, filter, validate, and import into a relational database.

## Current Assumptions and Limitations

This project was written for a specific migration workflow, so the current scripts intentionally make several assumptions:

- The source workbook follows the matching template exactly.
- The current templates/scripts are designed around 13 boxes.
- Box dimensions are limited to 9 × 9 and 10 × 10.
- Supported sample record lengths are 4 through 8 values.
- Input and output paths are configured directly in each script.
- The scripts use fixed row/block sizes rather than automatically detecting the source format.
- There is currently no interactive validation for missing rows, malformed boxes, or an incorrect template.
- The box identifier is derived from the template's box marker row.
- Each run appends another worksheet to the destination workbook.

These constraints were acceptable for the original controlled migration process, where the source format was known in advance.

## Possible Future Improvements

A future version could consolidate the individual scripts into a single configurable conversion utility by adding:

- command-line arguments for box size and field count,
- automatic detection of 9 × 9 versus 10 × 10 layouts,
- automatic detection of the number of values per sample,
- dynamic handling of any number of boxes,
- validation of box markers and the final end marker,
- configurable input/output paths,
- clearer error messages for malformed source data,
- automated sorting by box and position,
- unit tests for each supported source format, and
- a simple graphical interface for non-technical users.

## Technologies Used

- **Python** — conversion logic
- **pandas** — spreadsheet data loading, restructuring, and DataFrame creation
- **openpyxl** — Excel workbook output/append engine
- **Excel (`.xlsx`)** — legacy input templates and converted output

## Data Privacy

The example files in a public portfolio repository should contain only synthetic, anonymized, or otherwise non-sensitive data. Production database exports, confidential sample identifiers, credentials, and proprietary records should not be committed to GitHub.

## Project Purpose

This project demonstrates practical experience with:

- data migration,
- data transformation and normalization,
- Python scripting,
- Excel automation,
- pandas DataFrames,
- working with inconsistent legacy data structures,
- translating physical storage layouts into database-ready records, and
- reducing manual work during a database transition.

## License

No license has been specified for this portfolio project.
