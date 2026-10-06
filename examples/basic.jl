using Dates, XLSXTables

transactions = (
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
    column_widths=Dict(:date => 14, :description => 28, :amount => 18, :tax_rate => 14),
)

write_xlsx("transactions.xlsx", transactions;
    sheetname="Transactions", style, overwrite=true,
    column_labels=Dict(:date => "Date", :description => "Description",
        :amount => "Amount (€)", :tax_rate => "Tax rate"),
)
