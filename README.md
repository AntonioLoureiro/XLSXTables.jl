# XLSXTables.jl

Export any [Tables.jl](https://github.com/JuliaData/Tables.jl) table to an Excel
workbook with reusable formatting, powered by
[XLSX.jl](https://github.com/JuliaData/XLSX.jl).

Supports column tables, vectors of named tuples, DataFrames, CSV.File, and other
Tables.jl sources. DataFrames and CSV are optional inputs, not dependencies.

## Installation

Requires Julia 1.10 or later. This package is not yet registered in General.

```julia
using Pkg
Pkg.add(url="https://github.com/AntonioLoureiro/XLSXTables.jl")
```

For a local checkout, use `Pkg.develop(path="/path/to/XLSXTables.jl")`.

## Quick start

```julia
using Dates, XLSXTables

data = (
    date=[Date(2026, 10, 1), Date(2026, 10, 2)],
    description=["Consulting", "Software"],
    amount=[1250.0, 89.90],
    tax_rate=[0.23, 0.23],
)

style = TableStyle(
    column_styles=Dict(
        :amount => CellStyle(number_format="#,##0.00 \"€\""),
        :tax_rate => CellStyle(number_format="0.0%"),
    ),
    column_widths=Dict(:description => 28, :amount => 18),
)

write_xlsx("report.xlsx", data; sheetname="Transactions", style)
```

Existing output files require `overwrite=true`. Number formatting changes the
display of values; a percentage such as `0.23` remains numeric `0.23` in Excel.

## Default number formats

| Julia column type | Excel format |
| --- | --- |
| `Bool` | `General` |
| `Integer` | `#,##0` |
| `AbstractFloat` | `#,##0.00` |
| `Date` | `yyyy-mm-dd` |
| `DateTime` | `yyyy-mm-dd hh:mm:ss` |
| `Time` | `hh:mm:ss` |
| `AbstractString` | `@` |

`missing` and `nothing` become blank cells. They are excluded from column type
inference. For abstract or `Any` columns, the common type of nonblank values is
used. A truly heterogeneous column may have no matching rule, in which case its
body format is retained. An explicit column style always applies.

## Formatting layers

Data cells receive **body style → first matching type rule → column override**.
Named-tuple attributes merge, so setting a column's font to `(bold=true,)`
retains the body font name and size. Number formats are replaced as a whole;
`nothing` inherits, while `"General"` explicitly resets the number format.
Headers use their own style and do not inherit body formatting.

```julia
style = TableStyle(
    header=CellStyle(
        font=(bold=true, color="FFFFFFFF"),
        fill=(pattern="solid", fgColor="FF24476B"),
        alignment=(wrapText=true,),
    ),
    body=CellStyle(font=(name="Calibri", size=11)),
    type_formats=[Float64 => "#,##0.0000", default_type_formats()...],
    column_styles=Dict(
        :amount => CellStyle(number_format="#,##0.00 \"€\"", font=(bold=true,)),
    ),
)
```

Type rules are ordered; **the first match wins**. Prepend specific rules to
`default_type_formats()` to override defaults, or supply your own complete
list. Set `type_formats=[]` to use only body and column formatting.

`CellStyle` supports `number_format`, `font`, `fill`, `border`, and `alignment`.
The four named-tuple fields forward keywords to XLSX.jl's corresponding public
formatting functions. For example:

```julia
CellStyle(
    border=(bottom=["style" => "thin", "color" => "FFDDDDDD"],),
    alignment=(horizontal="right", wrapText=true),
)
```

See the [XLSX formatting API](https://juliadata.org/XLSX.jl/stable/api/formats/)
for accepted attribute names and values. `column_styles`, `column_widths`, and
`column_labels` accept dictionaries or named tuples with original column names;
unknown names raise an error to catch typos. Widths are in Excel character units.

## Multiple worksheets

```julia
write_xlsx("report.xlsx", "Transactions" => data, "Summary" => (total=[1339.90],))
```

Shared keyword options apply to every sheet. Use `write_table!` when sheets
need different styles or layouts:

```julia
import XLSX

book = XLSX.newxlsx()
sheet = book[1]
XLSX.renamesheet!(sheet, "Report")
sheet["A1"] = "Monthly report"

write_table!(sheet, data;
    anchor="B3", style,
    column_labels=Dict(:amount => "Amount (€)", :tax_rate => "Tax rate"),
)

summary = XLSX.addsheet!(book, "Summary")
write_table!(summary, (total=[1339.90],); style=TableStyle())
XLSX.writexlsx("custom-report.xlsx", book; overwrite=true)
```

`write_table!` returns the worksheet and only writes inside the table's
rectangle; an explicit column width affects the whole worksheet column.
Use `header=false` to omit headers. Tables with no rows but known columns can
write just their headers. Tables without columns are rejected.

## Empty workbook, reusable header style, and a DataFrame

The runnable example [examples/dataframe_workbook.jl](examples/dataframe_workbook.jl)
creates an empty file, defines a reusable header style, then adds a sheet from
a DataFrame and saves the completed workbook.

From the repository root, install the optional example dependencies and run it:

```sh
julia --project=examples -e 'using Pkg; Pkg.develop(path="."); Pkg.instantiate()'
julia --project=examples examples/dataframe_workbook.jl
```

It creates `dataframe_report.xlsx` in the current directory. You can supply a
different filename as the script's first argument. The initial save refuses
to overwrite an existing file.

The main steps are:

```julia
using DataFrames, Dates, XLSXTables
import XLSX

workbook = XLSX.newxlsx()
placeholder = workbook[1]
XLSX.writexlsx("dataframe_report.xlsx", workbook)

default_header = CellStyle(
    font=(name="Calibri", size=11, bold=true, color="FFFFFFFF"),
    fill=(pattern="solid", fgColor="FF24476B"),
    alignment=(horizontal="left", vertical="center", wrapText=true),
)
default_style = TableStyle(header=default_header)

df = DataFrame(
    date=[Date(2026, 10, 1), Date(2026, 10, 2)],
    description=["Consulting", "Software"],
    quantity=[10, 1],
    amount=[1250.0, 89.90],
)

sheet = XLSX.addsheet!(workbook, "Transactions")
write_table!(sheet, df; style=default_style)

XLSX.deletesheet!(placeholder)
XLSX.writexlsx("dataframe_report.xlsx", workbook; overwrite=true)
```

`default_style` is a reusable configuration passed to each `write_table!`
call; the current API does not register defaults on a workbook. Body number
formats still come from the standard type rules. The initial blank sheet is
removed after adding `Transactions`, leaving just the populated sheet.

## Scope and data handling

- Exports formatted worksheet cells; does not create native Excel Table objects.
- Materialises input columns in memory; this is not a streaming writer.
- Strings starting with `=` remain text. Formula creation is left to XLSX.jl.
- Integers with more than 15 decimal digits are stored as text to preserve
  identifiers exactly. Other integer and real values are converted to Excel's
  supported scalar representations. Real values use `Float64`, so arbitrary
  precision decimals can lose precision; convert them to strings if required.
- Symbols are converted to strings. Unsupported values, NaN, and infinity are
  rejected before writing table values.
- The destination is written only after the workbook has been built and saved
  to a temporary file. An error during formatting leaves an existing output
  file intact. In-memory `write_table!` calls can be partially modified if an
  XLSX formatting keyword is invalid.
- Updating an existing complex Excel workbook is subject to XLSX.jl's own
  round-trip limitations. Keep the original when working with such files.

## Development

```sh
julia --project -e 'using Pkg; Pkg.instantiate(); Pkg.test()'
julia --project examples/basic.jl
```

CI tests Julia 1.10 and the latest stable Julia, including Linux, Windows, and
macOS. Licensed under the MIT License.
