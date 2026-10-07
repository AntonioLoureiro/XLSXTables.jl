# Run from the repository root:
#   julia --project examples/automatic_widths.jl [output.xlsx]
using Dates, XLSXTables
import XLSX

output = isempty(ARGS) ? "automatic_widths.xlsx" : first(ARGS)
data = (
    date=[Date(2026, 10, 1), Date(2026, 10, 2), Date(2026, 10, 3)],
    description=["Consulting", "Software subscription",
        "Annual support including deployment, onboarding, and team training"],
    quantity=[10, 1, 25],
    amount=[1234567.89, -2500.5, 0.0],
    tax_rate=[0.23, 0.125, 0.0],
)
formats = [
    Bool => "General",
    Integer => "#,##0;[Red](#,##0);\"-\"",
    AbstractFloat => "#,##0.00;[Red](#,##0.00);\"-\"",
    default_type_formats()...,
]
columns = Dict(:tax_rate => CellStyle(number_format="0.0%"))

workbook = XLSX.newxlsx()
sheet = workbook[1]
XLSX.renamesheet!(sheet, "Automatic")
# No auto_width argument: sizing is on by default, based on the displayed data.
write_table!(sheet, data; style=TableStyle(type_formats=formats, column_styles=columns))

sheet = XLSX.addsheet!(workbook, "Configured")
defaults = TableStyle(
    auto_width=true,
    min_width=10,
    max_width=35,
    width_padding=3,
    wrap_long_text=true,
    type_formats=formats,
    column_styles=columns,
    column_widths=Dict(:amount => 22), # Overrides the automatic estimate.
)
write_table!(sheet, data; style=defaults)

sheet = XLSX.addsheet!(workbook, "Disabled")
sheet["B1"] = "Existing header" # Establish a column, as in a template.
XLSX.setColumnWidth(sheet, "B"; width=30)
# The per-call switch overrides defaults.auto_width. Description keeps its
# existing width of 30; the explicit amount width of 22 still applies.
write_table!(sheet, data; style=defaults, auto_width=false)

XLSX.writexlsx(output, workbook) # Refuse to overwrite an existing file.
println("Created ", abspath(output))
