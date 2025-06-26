# vba formula parser

A library to parse formulas

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

Import Formulas.bas into your project
Include "Microsoft Scripting Runtime"

## Supported Syntax

- Operators
  - Arithmetic (+, -, \*, /, ^)
  - Comparison (=, <>, <, >, <=, >=)
  - Concatenation (&)
- Function Calls
- Constants
  - Numeric
  - String
  - Logical (TRUE, FALSE)
  - Array constants (e.g., {1,2;3,4}).

## Features

- Formula Formatting
- AST Generation
