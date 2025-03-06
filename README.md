# Levenshtein Microsoft Excel Formula

A feature-rich and easy-to-use Levenshtein distance formula written entirely in Excel without VBA macros.

## Key Features

- Case sensitivity management
- Configurable output format (percentage or integer distance)
- No length limit (successfully tested with words up to 10,000 characters)
- Optional output table for debugging and visualization

## Installation

### Prerequisites
**Microsoft 365 is required** to ensure full compatibility with dynamic array formulas.

### Steps
1. Open Excel.
2. Go to **Formulas** > **Name Manager**.
3. Create a new name and paste the formula.

### Recommended Settings
- **Name:** `levenshtein_V1`
- **Description:**
  ```
  caseSensitive:
  - true/t
  - false/f

  outputType:
  - percent/p/%
  - integer/i/numeric/number

  printTable:
  - true/t
  - false/f
  - numeric only/no
  ```

## Usage
The formula accepts three parameters:

```excel
=levenshtein_V1(text1, text2, caseSensitive, outputType, printTable)
```
- `text1`: First text string
- `text2`: Second text string
- `caseSensitive`: Boolean or keyword to enable case sensitivity (`true`/`t`) or disable it (`false`/`f`)
- `outputType`: Choose between percentage (`percent`/`p/%`) or integer distance (`integer`/`i/numeric/number`)
- `printTable`: Boolean or keyword to print the comparison table (`true`/`t`) or only return the result (`false`/`f`)

## Current Version
- Version: **1.0**
- Repository: [GitHub](https://github.com/coulombevin/levenshtein_excel_formula/blob/main/levenshtein_V1.0.md)

## License
This project is licensed under the MIT License - see the LICENSE file for details.

