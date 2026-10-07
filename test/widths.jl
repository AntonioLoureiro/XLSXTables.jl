function iswrapped(sheet, cell)
    alignment = XLSX.getAlignment(sheet, cell)
    return !isnothing(alignment) && get(alignment.alignment["alignment"], "wrapText", "0") == "1"
end

@testset "Automatic column widths" begin
    @test TableStyle().auto_width
    @testset "Formatted display estimates" begin
        display(value, format) = only(XLSXTables._display_candidates(value, XLSXTables._width_sections(format)))
        @test display(1234.5, "#,##0.00") == "1,234.50"
        @test display(-1234.5, "#,##0.00;[Red](#,##0.00);\"-\"") == "(1,234.50)"
        @test display(-12, "0") == "-12"
        @test display(0, "0;[Red](0);\"-\"") == "-"
        @test display(999.999, "#,##0.00") == "1,000.00"
        @test display(0.125, "0.0%") == "12.5%"
        @test display(12.5, "0.0\"%\"") == "12.5%"
        @test display(12.5, "0.0\\%") == "12.5%"
        @test display(12, "00000") == "00012"
        @test display(1234567, "0.0,,") == "1.2"
        @test display(12345.0, "0.00E+000") == "1.23E+004"
        @test display(12, "\"USD; \"#,##0.00") == "USD; 12.00"
        @test display(1234.5, "[\$€-407]#,##0.00") == "€1,234.50"
        @test display(Date(2026, 9, 30), "dddd, mmmm d, yyyy") == "Wednesday, September 30, 2026"
        @test display(DateTime(2026, 9, 30, 14, 5, 6), "dd-mm-yyyy hh:mm:ss AM/PM") == "30-09-2026 02:05:06 PM"
        @test display(Time(14, 5, 6), "hh:mm:ss") == "14:05:06"
        @test display(missing, "0") == ""
        @test display(true, "General") == "TRUE"
        @test display("sample", "0;0;0;\"Label: \"@") == "Label: sample"
        # Unsupported fractional masks still produce an estimate, without
        # affecting the format written into the workbook.
        @test !isempty(display(1.5, "# ?/?"))
        conditional = XLSXTables._display_candidates(-1200.5,
            XLSXTables._width_sections("[>=1000]#,##0.00;[Red](#,##0.00)"))
        @test any(s -> occursin("1,200.50", s), conditional)
    end

    mktempdir() do dir
        @testset "Written widths, formatted values, and overrides" begin
            data = (description=[repeat("Long text ", 12)], amount=[-1234567.89],
                day=[Date(2026, 9, 30)], identifier=[big(10)^25], fixed=[repeat("x", 80)])
            style = TableStyle(
                column_styles=Dict(
                    :amount => CellStyle(number_format="\"USD \"#,##0.00;[Red](\"USD \"#,##0.00);\"-\""),
                    :day => CellStyle(number_format="dddd, mmmm d, yyyy"),
                ),
                column_widths=Dict(:fixed => 18),
            )
            path = joinpath(dir, "widths.xlsx")
            write_xlsx(path, data; anchor="B3", style)
            sheet = XLSX.opentemplate(path)[1]
            @test 50 <= XLSX.getColumnWidth(sheet, "B4") < 51
            @test 18 <= XLSX.getColumnWidth(sheet, "C4") < 23
            @test 26 <= XLSX.getColumnWidth(sheet, "D4") < 35
            @test 28 <= XLSX.getColumnWidth(sheet, "E4") < 29
            @test 18 <= XLSX.getColumnWidth(sheet, "F4") < 19
            @test iswrapped(sheet, "B4")
            @test !iswrapped(sheet, "F4") # Explicit width leaves wrapping alone.
            @test sheet["C4"] == -1234567.89
            @test sheet["D4"] == Date(2026, 9, 30)
            @test sheet["E4"] == string(big(10)^25)
            @test XLSX.getFormat(sheet, "C4").format["numFmt"]["formatCode"] == style.column_styles[:amount].number_format
        end

        @testset "Default on, style defaults, and per-call switch" begin
            data = (x=["12345678901234567890"], y=[1])
            path = joinpath(dir, "switch.xlsx")
            write_xlsx(path, "One" => data, "Two" => data)
            book = XLSX.opentemplate(path)
            @test all(22 <= XLSX.getColumnWidth(book[name], "A1") < 23 for name in ("One", "Two"))
            @test 10 <= XLSX.getColumnWidth(book[1], "B1") < 11

            write_xlsx(path, "One" => data, "Two" => data; overwrite=true, auto_width=false)
            book = XLSX.opentemplate(path)
            @test all(isnothing(XLSX.getColumnWidth(book[name], "A1")) for name in ("One", "Two"))

            off = TableStyle(auto_width=false, column_widths=Dict(:y => 23))
            write_xlsx(path, data; overwrite=true, style=off)
            sheet = XLSX.opentemplate(path)[1]
            @test isnothing(XLSX.getColumnWidth(sheet, "A1"))
            @test 23 <= XLSX.getColumnWidth(sheet, "B1") < 24
            write_table!(sheet, data; style=off, auto_width=true)
            @test 22 <= XLSX.getColumnWidth(sheet, "A1") < 23
            @test 23 <= XLSX.getColumnWidth(sheet, "B1") < 24

            XLSX.setColumnWidth(sheet, "A"; width=37)
            original = XLSX.getColumnWidth(sheet, "A1")
            write_table!(sheet, data; auto_width=false)
            @test XLSX.getColumnWidth(sheet, "A1") == original
            write_table!(sheet, data; style=TableStyle(auto_width=false))
            @test XLSX.getColumnWidth(sheet, "A1") == original
            write_table!(sheet, (x=String[],); header=false)
            @test XLSX.getColumnWidth(sheet, "A1") == original
        end

        @testset "Labels, fonts, bounds, wrapping, and empty columns" begin
            sheet = XLSX.newxlsx()[1]
            labels = Dict(:x => repeat("H", 30))
            write_table!(sheet, (x=[1],); column_labels=labels)
            @test XLSX.getColumnWidth(sheet, "A1") > 30
            write_table!(sheet, (x=[1],); column_labels=labels, header=false)
            @test 10 <= XLSX.getColumnWidth(sheet, "A1") < 11
            write_table!(sheet, (x=Int[],); column_labels=labels)
            @test XLSX.getColumnWidth(sheet, "A1") > 30

            data = (x=["12345678901234567890"],)
            write_table!(sheet, data)
            ordinary = XLSX.getColumnWidth(sheet, "A1")
            write_table!(sheet, data; style=TableStyle(width_padding=5))
            @test XLSX.getColumnWidth(sheet, "A1") ≈ ordinary + 3
            write_table!(sheet, data; style=TableStyle(body=CellStyle(font=(size=22,))))
            @test XLSX.getColumnWidth(sheet, "A1") > ordinary * 1.8
            write_table!(sheet, (x=["界"^15],))
            @test XLSX.getColumnWidth(sheet, "A1") > ordinary
            write_table!(sheet, (x=["1234567890\n1234567890"],))
            @test 12 <= XLSX.getColumnWidth(sheet, "A1") < 13
            write_table!(sheet, (x=[missing],); style=TableStyle(min_width=17))
            @test 17 <= XLSX.getColumnWidth(sheet, "A1") < 18

            for (style, expected) in (
                (TableStyle(max_width=12), true),
                (TableStyle(max_width=12, wrap_long_text=false), false),
                (TableStyle(max_width=12, body=CellStyle(alignment=(wrapText=false,))), false),
                (TableStyle(max_width=12, column_styles=(x=CellStyle(alignment=(wrapText=false,)),)), false),
            )
                sheet = XLSX.newxlsx()[1]
                write_table!(sheet, data; style)
                @test 12 <= XLSX.getColumnWidth(sheet, "A1") < 13
                @test iswrapped(sheet, "A2") == expected
            end
            sheet = XLSX.newxlsx()[1]
            write_table!(sheet, data; column_labels=labels,
                style=TableStyle(header=CellStyle(), max_width=12))
            @test iswrapped(sheet, "A1")
            sheet = XLSX.newxlsx()[1]
            write_table!(sheet, data; column_labels=labels,
                style=TableStyle(header=CellStyle(alignment=(wrapText=false,)), max_width=12))
            @test !iswrapped(sheet, "A1")
        end
    end

    @testset "Width configuration validation" begin
        for options in ((min_width=-1,), (max_width=256,), (min_width=20, max_width=10),
            (min_width=NaN,), (max_width=Inf,), (width_padding=-1,), (width_padding=Inf,))
            @test_throws ArgumentError TableStyle(; options...)
        end
        original = TableStyle()
        positional = TableStyle(original.header, original.body, original.type_formats,
            original.column_styles, original.column_widths)
        @test positional.auto_width
    end
end
