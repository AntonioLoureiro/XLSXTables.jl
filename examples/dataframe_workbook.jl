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

# 2. Define the header style to reuse for every table in this workbook.
default_header = CellStyle(
    font=(name="Calibri", size=11, bold=true, color="FFFFFFFF"),
    fill=(pattern="solid", fgColor="FF24476B"),
    alignment=(horizontal="left", vertical="center", wrapText=true),
)
default_style = TableStyle(header=default_header)

# 3. Create a DataFrame. TableStyle's type rules automatically format dates,
#    integers, decimal numbers, and text in the body.
df = DataFrame(
    date=[Date(2026, 10, 1), Date(2026, 10, 2)],
    description=["Consulting", "Software"],
    quantity=[10, 1],
    amount=[1250.0, 89.90],
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
