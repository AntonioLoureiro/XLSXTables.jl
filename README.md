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

## Default appearance

`TableStyle()` includes the following defaults, also used when no `style` is
passed to `write_xlsx` or `write_table!`:

- Headers: Calibri 11, bold white text, solid orange `#FD5108` background,
  horizontally and vertically centered, with text wrapping enabled.
- Data: Calibri 11 with a solid light-grey `#F2F2F2` background.
- Automatic column widths enabled, with number formats chosen by column type.

No style setup is needed to get this appearance:

```julia
write_xlsx("report.xlsx", df)
# Or, in an existing workbook:
write_table!(sheet, df)
```

Custom `header`, `body`, `type_formats`, and `column_styles` remain available
through `TableStyle`. Colors are stored as opaque ARGB (`FFFD5108` and
`FFF2F2F2`).

## Default number formats

| Julia column type | Excel format |
| --- | --- |
| `Bool` | `General` |
| `Integer` | `0` |
| `AbstractFloat` | `#,##0.00` |
| `Date` | `yyyy-mm-dd` |
| `DateTime` | `yyyy-mm-dd hh:mm:ss` |
| `Time` | `hh:mm` |
| `AbstractString` | `@` |

`missing` and `nothing` become blank cells. They are excluded from column type
inference. For abstract or `Any` columns, the common type of nonblank values is
used. A truly heterogeneous column may have no matching rule, in which case its
body format is retained. An explicit column style always applies.

### Define your own defaults by type

Pass an ordered list of `JuliaType => "Excel format code"` rules to
`TableStyle(type_formats=...)`:

```julia
using Dates, XLSXTables

default_formats = [
    Bool => "General",
    Integer => "0",
    AbstractFloat => "#,##0.000",
    Date => "dd-mm-yyyy",
    DateTime => "dd-mm-yyyy hh:mm:ss",
    Time => "hh:mm",
    AbstractString => "@",
]

default_style = TableStyle(
    header=CellStyle(font=(bold=true,)),
    type_formats=default_formats,
)

write_table!(sheet, df; style=default_style)
```

`Integer` covers `Int`, `Int32`, `Int64`, and unsigned integer types.
`AbstractFloat` covers floating-point types such as `Float32` and `Float64`.
You can also use a concrete type, such as `Float32`, for a more specific rule.
**The first matching rule wins**; put `Bool` before `Integer` because
`Bool <: Integer`. Likewise, put `Float32` before `AbstractFloat` if both appear.

These are Excel display format codes, not Julia `Dates.DateFormat` patterns.
Rules apply to each column's nonmissing type, including columns such as
`Union{Missing, Float64}`. Supplying `type_formats` replaces the built-in rule
list; reuse the same `TableStyle` in subsequent exports to keep your defaults.
Use `column_styles` for exceptions, for example
`:tax_rate => CellStyle(number_format="0.0%")`.

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
list. When prepending an `Integer` rule, put a `Bool` rule before it if booleans
should keep their own format. Set `type_formats=[]` to use only body and column
formatting.

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

## Automatic column widths

Column sizing is **on by default** for both `write_xlsx` and `write_table!`.
Each column uses its header label and formatted values to estimate the width,
with 2 units of padding, a minimum of 10, and a maximum of 50. No extra option
is needed:

```julia
write_xlsx("report.xlsx", data)
```

Configure reusable defaults through `TableStyle`:

```julia
style = TableStyle(
    auto_width=true,
    min_width=10,
    max_width=45,
    width_padding=2,
    wrap_long_text=true,
    column_widths=Dict(:description => 28), # Explicit widths always win.
)
write_xlsx("report.xlsx", data; style)
```

`column_widths` overrides the estimate even outside the automatic limits.
Text exceeding `max_width` is wrapped unless its alignment explicitly sets
`wrapText`. Set `wrap_long_text=false` to disable this automatic wrapping;
explicit header/body alignment still applies. Manual width overrides leave
wrapping unchanged.

Turn sizing off for a single export, or in a reusable style:

```julia
write_xlsx("report.xlsx", data; auto_width=false)
write_table!(sheet, data; style=TableStyle(auto_width=false))
```

The per-call `auto_width` keyword overrides `style.auto_width`. When sizing is
off, existing worksheet widths are preserved except for explicit
`column_widths`. With sizing on, widths are calculated from the current table,
and affect the **entire worksheet column**, including cells outside the table.
They are saved in the workbook and do not recalculate after edits in Excel.

Estimates account for common numeric formats (grouping, decimal places,
percentages, currency, positive/negative/zero sections), date/time formats,
font size, bold text, Unicode width, and the longest line of multiline text.
`column_labels` is used for displayed headers; `header=false` excludes them.
Empty tables size their headers, while empty tables without headers leave
automatic widths alone.

This approximates Excel AutoFit. Font metrics, localized date/currency text,
uncommon custom formats, and conditional-formatting rules added after export
can differ from the estimate. Set an explicit width for precise layouts.
Wrapping does not set a fixed row height; the spreadsheet viewer controls it.

See [examples/automatic_widths.jl](examples/automatic_widths.jl) for a runnable
workbook comparing automatic sizing, reusable limits and overrides, and the
disabled option:

```sh
julia --project examples/automatic_widths.jl
```

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

`write_table!` returns the worksheet and writes cell values and styles inside
the table's rectangle; automatic and explicit widths affect the whole worksheet
column. Use `auto_width=false` to preserve existing widths unless overridden.
Use `header=false` to omit headers. Tables with no rows but known columns can
write just their headers. Tables without columns are rejected.

## Empty workbook and a DataFrame with default styling

The runnable example [examples/dataframe_workbook.jl](examples/dataframe_workbook.jl)
creates an empty file, adds a sheet from a DataFrame using the package defaults,
and saves the completed workbook. It also demonstrates a percentage override
for one column.

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

default_style = TableStyle(
    column_styles=Dict(:tax_rate => CellStyle(number_format="0.0%")),
)

df = DataFrame(
    date=[Date(2026, 10, 1), Date(2026, 10, 2)],
    description=["Consulting", "Software"],
    quantity=[10, 1],
    amount=[1250.0, 89.90],
    tax_rate=[0.23, 0.23],
    recorded_at=[DateTime(2026, 10, 1, 9, 30), DateTime(2026, 10, 2, 14, 15)],
    start_time=[Time(9, 30), Time(14, 15)],
    approved=[true, false],
)

sheet = XLSX.addsheet!(workbook, "Transactions")
write_table!(sheet, df; style=default_style)

XLSX.deletesheet!(placeholder)
XLSX.writexlsx("dataframe_report.xlsx", workbook; overwrite=true)
```

`default_style` is a reusable configuration passed to each `write_table!`
call; the current API does not register defaults on a workbook. Header and body
styles, number formats, and automatic widths come from the package defaults,
with `tax_rate` displayed as a percentage through its column override. Omit
`style` entirely to use all built-in defaults. The initial blank sheet is removed
after adding `Transactions`, leaving just the populated sheet.

## Conditional formatting of numbers

The runnable [examples/conditional_numbers.jl](examples/conditional_numbers.jl)
demonstrates both number-format sections and native Excel conditional rules.
Run it with the examples environment configured above:

```sh
julia --project=examples examples/conditional_numbers.jl
```

### Different display for positive, negative, and zero values

Excel number formats can contain `positive;negative;zero` sections. These can
be used as your type defaults:

```julia
style = TableStyle(
    type_formats=[
        Bool => "General",
        Integer => "#,##0;[Red](#,##0);\"-\"",
        AbstractFloat => "#,##0.00;[Red](#,##0.00);\"-\"",
        default_type_formats()...,
    ],
)
```

Positive numbers use the first section, negative numbers appear in red with
parentheses, and zero appears as a dash. Stored values remain numeric. For a
percentage column, use `CellStyle(number_format="0.0%;[Red](0.0%);\"-\"")`.

### Highlight numbers against a threshold

After `write_table!`, use XLSX.jl's public conditional-formatting API on the
data cells. For a DataFrame written at `A1` with headers:

```julia
import XLSX

write_table!(sheet, df; style)
amount_column = findfirst(==(:amount), propertynames(df))

if nrow(df) > 0
    data_rows = 2:(nrow(df) + 1)
    XLSX.setConditionalFormat(sheet, data_rows, amount_column, :cellIs;
        operator="lessThan", value="0", dxStyle="redfilltext",
    )
    XLSX.setConditionalFormat(sheet, data_rows, amount_column, :cellIs;
        operator="greaterThan", value="1000", dxStyle="greenfilltext",
    )
end

XLSX.writexlsx("conditional_numbers.xlsx", workbook)
```

Negative amounts get a light red fill and dark red text; amounts greater than
1,000 get a light green fill and dark green text. The header is excluded.
Adjust the row and column offsets if you use a different anchor or omit the
header. XLSX.jl expects comparison values as strings, such as `"1000"`.

Conditional rules are applied through XLSX.jl after table export; they are not
fields of `TableStyle`. They are stored in the workbook and evaluated by Excel
when values change within the specified range. Rows added beyond that range
need an updated rule range. XLSX.jl also supports formulas, colour scales, and
data bars; see its [conditional-formatting guide](https://juliadata.org/XLSX.jl/stable/formatting/conditionalFormatting/).

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
