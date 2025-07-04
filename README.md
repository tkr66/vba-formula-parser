# vba formula formatter

A library to format formulas

## Examples

Format

```
in:
=CONCAT("R",MOD(ROW()-6,2)*2+1,"C",INT((ROW()-6)/2)*2+1)

out:
=CONCAT(
  "R",
  MOD(
    ROW() - 6,
    2
  ) * 2 + 1,
  "C",
  INT(
    (
      ROW() - 6
    ) / 2
  ) * 2 + 1
)
```

## Installation

1. Import Formulas.bas into your project
1. Include "Microsoft Scripting Runtime"

## Usage

To pretty-print a formula, use the `Pretty` function:

```vba
Dim originalFormula As String
originalFormula = "=CONCAT(""R"",MOD(ROW()-6,2)*2+1,""C"",INT((ROW()-6)/2)*2+1)"

Dim fmt As Formulas.Formatter
fmt = Formulas.NewFormatter( _
    indent:=" ", _
    indentLength:=2, _
    newLine:=vbCrLf, _
    eqAtStart:=True, _
    newLineAtEof:=True _
)

Dim formattedFormula As String
formattedFormula = Formulas.Pretty(originalFormula, fmt)
debug.Print formattedFormula
```

## Supported Syntax

- Operators
  - Arithmetic (+, -, \*, /)
  - Comparison (=, <>, <, >, <=, >=)
  - Concatenation (&)
  - Range Reference (:)
- Function Calls
- CellReferences
  - Absolute
  - Relative
  - Mixed (e.g., $A1, A$1)
- Constants
  - Numeric
  - String
  - Logical (TRUE, FALSE)
  - Array constants (e.g., {1,2;3,4}).

## Features

- Formula Formatting
