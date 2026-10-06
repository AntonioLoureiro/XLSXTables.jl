using DataFrames, XLSXTables
import XLSX

# Run from the repository root after setting up the examples environment:
# julia --project=examples examples/conditional_numbers.jl [output.xlsx]
output = isempty(ARGS) ? "conditional_numbers.xlsx" : only(ARGS)

# 1. Display positive, negative, and zero values differently using the three
#    sections of an Excel number format: positive;negative;zero.
#    Negatives appear in red and parentheses; zeros appear as a dash.
style = TableStyle(
    type_formats=[
        Bool => "General",
        Integer => "#,##0;[Red](#,##0);\"-\"",
        AbstractFloat => "#,##0.00;[Red](#,##0.00);\"-\"",
        default_type_formats()...,
    ],
    column_styles=Dict(
        :margin => CellStyle(number_format="0.0%;[Red](0.0%);\"-\""),
    ),
    column_widths=Dict(:description => 24, :amount => 18, :quantity => 14, :margin => 14),
)

df = DataFrame(
    description=["Invoice", "Refund", "No activity", "Small invoice", "Large invoice"],
    amount=[1500.0, -250.0, 0.0, 900.0, 2500.0],
    quantity=[15, -2, 0, 9, 25],
    margin=[0.25, -0.10, 0.0, 0.05, 0.30],
)

workbook = XLSX.newxlsx()
sheet = workbook[1]
XLSX.renamesheet!(sheet, "Numbers")
write_table!(sheet, df; style)

# 2. Add native Excel conditional-formatting rules after writing the data.
#    Find the column by name. With the default A1 anchor and header=true,
#    data starts on row 2; the header is excluded from both rules.
amount_column = findfirst(==(:amount), propertynames(df))
threshold = 1000.0
if nrow(df) > 0
    data_rows = 2:(nrow(df) + 1)
    XLSX.setConditionalFormat(sheet, data_rows, amount_column, :cellIs;
        operator="lessThan", value="0", dxStyle="redfilltext",
    )
    XLSX.setConditionalFormat(sheet, data_rows, amount_column, :cellIs;
        operator="greaterThan", value=string(threshold), dxStyle="greenfilltext",
    )
end

# These rules are saved in the workbook. Excel evaluates them, including when
# you edit values in the formatted range. XLSXTables does not evaluate them.
XLSX.writexlsx(output, workbook)  # Refuse to overwrite an existing file.
println("Created ", abspath(output))
