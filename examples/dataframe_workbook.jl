using DataFrames, Dates, XLSXTables
import XLSX

# Run from the repository root after setting up the examples environment:
# julia --project=examples examples/dataframe_workbook.jl [output.xlsx]
output = isempty(ARGS) ? "dataframe_report.xlsx" : only(ARGS)

# 1. Create and save an empty workbook. Excel requires at least one sheet,
#    so XLSX.newxlsx() starts with one blank placeholder sheet.
workbook = XLSX.newxlsx()
placeholder = workbook[1]
XLSX.writexlsx(output, workbook)  # Refuse to overwrite an existing file.

# 2. Define the header style and type formats to reuse for every table.
default_header = CellStyle(
    font=(name="Calibri", size=11, bold=true, color="FFFFFFFF"),
    fill=(pattern="solid", fgColor="FF24476B"),
    alignment=(horizontal="left", vertical="center", wrapText=true),
)
default_formats = [
    Bool => "General",                 # Before Integer: Bool is an Integer subtype.
    Integer => "0",                    # Int, Int32, Int64, unsigned integers, etc.
    AbstractFloat => "#,##0.000",       # Float32, Float64, etc.; three decimal places.
    Date => "dd-mm-yyyy",
    DateTime => "dd-mm-yyyy hh:mm:ss",
    Time => "hh:mm",
    AbstractString => "@",
]
default_style = TableStyle(
    header=default_header,
    type_formats=default_formats,
    column_styles=Dict(
        # A column override takes precedence over its type's default format.
        :tax_rate => CellStyle(number_format="0.0%"),
    ),
)

# 3. Create a DataFrame. Formats are chosen from each column's nonmissing type;
#    formatting changes the displayed value, not the stored value.
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

# 4. Add a new sheet and write the DataFrame, using the reusable style.
#    Styles are passed explicitly; they are not attached to the workbook.
sheet = XLSX.addsheet!(workbook, "Transactions")
write_table!(sheet, df; style=default_style)

# 5. Remove the initial blank sheet now that a populated sheet exists,
#    then save over the empty file created above.
XLSX.deletesheet!(placeholder)
XLSX.writexlsx(output, workbook; overwrite=true)

println("Created ", abspath(output))
