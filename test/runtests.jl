using Test, Dates, XLSXTables
import XLSX, Tables

# Row-access-only table: proves that the writer does not require DataFrames.
struct RowSource{T}
    rows::Vector{T}
end
Tables.istable(::Type{<:RowSource}) = true
Tables.rowaccess(::Type{<:RowSource}) = true
Tables.rows(source::RowSource) = source.rows

formatcode(sheet, cell) = XLSX.getFormat(sheet, cell).format["numFmt"]["formatCode"]

@testset "XLSXTables" begin
    mktempdir() do dir
        @testset "Values, default formats, missing data, and headers" begin
            table = (
                id=[1, 2], amount=Union{Missing,Float64}[missing, 1234.5],
                day=[Date(2026, 10, 6), Date(2026, 10, 7)],
                stamp=[DateTime(2026, 10, 6, 9, 30), DateTime(2026, 10, 7, 10)],
                clock=[Time(9, 30), Time(10)], active=[true, false],
                label=["0012", "=1+1"], blank=[nothing, nothing],
            )
            path = joinpath(dir, "default.xlsx")
            @test write_xlsx(path, table; sheetname="Data") == path
            sheet = XLSX.readxlsx(path)["Data"]
            @test sheet["A1"] == "id"
            @test sheet["A2"] == 1
            @test ismissing(sheet["B2"])
            @test sheet["B3"] == 1234.5
            @test sheet["C2"] == Date(2026, 10, 6)
            @test sheet["D2"] == DateTime(2026, 10, 6, 9, 30)
            @test sheet["E2"] == Time(9, 30)
            @test sheet["F2"] === true
            @test sheet["F3"] === false
            @test sheet["G2"] == "0012"
            @test sheet["G3"] == "=1+1"
            @test ismissing(sheet["H2"])
            @test formatcode(sheet, "A2") == "#,##0"
            @test formatcode(sheet, "B3") == "#,##0.00"
            @test formatcode(sheet, "C2") == "yyyy-mm-dd"
            @test formatcode(sheet, "D2") == "yyyy-mm-dd hh:mm:ss"
            @test formatcode(sheet, "E2") == "hh:mm:ss"
            @test formatcode(sheet, "G2") == "@"
            @test haskey(XLSX.getFont(sheet, "A1").font, "b")
            @test XLSX.getFill(sheet, "A1").fill["patternFill"]["fgrgb"] == "FF24476B"
        end

        @testset "Layered style overrides and anchors" begin
            style = TableStyle(
                body=CellStyle(font=(name="Arial", size=12), alignment=(horizontal="right",)),
                type_formats=[Float64 => "0.0000", default_type_formats()...],
                column_styles=Dict(
                    :rate => CellStyle(number_format="0.0%", font=(bold=true,)),
                    :count => CellStyle(number_format="General"),
                ),
                column_widths=Dict("rate" => 24),
            )
            path = joinpath(dir, "overrides.xlsx")
            write_xlsx(path, (rate=[0.125], value=[3.14], count=[4]);
                anchor="B3", column_labels=Dict(:rate => "Rate (%)"), style)
            sheet = XLSX.readxlsx(path)[1]
            @test sheet["B3"] == "Rate (%)"
            @test sheet["B4"] ≈ 0.125
            @test sheet["C3"] == "value"
            @test formatcode(sheet, "B4") == "0.0%"
            @test formatcode(sheet, "C4") == "0.0000"
            @test formatcode(sheet, "D4") == "General"
            font = XLSX.getFont(sheet, "B4").font
            @test font["name"]["val"] == "Arial"
            @test font["sz"]["val"] == "12"
            @test haskey(font, "b")
            # XLSX adds cell padding to the requested character width on disk.
            @test 24 <= XLSX.getColumnWidth(XLSX.opentemplate(path)[1], "B4") < 25
            @test XLSX.getAlignment(sheet, "B4").alignment["alignment"]["horizontal"] == "right"
        end

        @testset "Table sources, worksheets and empty tables" begin
            rows = [(name="Ana", score=1.5), (name="Rui", score=2.5)]
            path = joinpath(dir, "sheets.xlsx")
            write_xlsx(path, "Rows" => rows, :Custom => RowSource(rows),
                "Empty" => (id=Int[], value=Float64[]))
            book = XLSX.readxlsx(path)
            @test XLSX.sheetnames(book) == ["Rows", "Custom", "Empty"]
            @test book["Rows"]["A3"] == "Rui"
            @test book["Custom"]["B3"] == 2.5
            @test book["Empty"]["B1"] == "value"

            book = XLSX.newxlsx()
            sheet = book[1]
            sheet["A1"] = "Keep me"
            @test write_table!(sheet, (x=[3],); anchor="AA4", header=false) === sheet
            @test sheet["AA4"] == 3
            @test sheet["A1"] == "Keep me"
            @test write_table!(sheet, (x=Int[],); anchor="C5", header=false) === sheet
            @test write_table!(sheet, (x=[7],); anchor="XFD1048576", header=false) === sheet
            @test sheet["XFD1048576"] == 7
        end

        @testset "Abstract columns and scalar conversion" begin
            path = joinpath(dir, "types.xlsx")
            write_xlsx(path, (mixed=Any[missing, nothing, 1.25],
                small=Float32[1.5, 2.5, 3.5],
                big=[big(10)^25, big(2), big(3)],
                symbols=[:a, :b, :c], unsigned=UInt32[1, 2, 3]))
            sheet = XLSX.readxlsx(path)[1]
            @test formatcode(sheet, "A4") == "#,##0.00"
            @test sheet["B2"] == 1.5
            @test sheet["C2"] == string(big(10)^25)
            @test sheet["C3"] == 2
            @test sheet["D2"] == "a"
            @test sheet["E2"] == 1
        end

        @testset "Validation and protection of existing files" begin
            path = joinpath(dir, "protected.xlsx")
            table = (x=[1],)
            write_xlsx(path, table)
            original = read(path)
            @test_throws ArgumentError write_xlsx(path, table)
            @test read(path) == original
            @test_throws ArgumentError write_xlsx(path, (x=[NaN],); overwrite=true)
            @test read(path) == original
            @test_throws ArgumentError write_xlsx(path, table;
                overwrite=true, style=TableStyle(column_styles=Dict(:typo => CellStyle())))
            @test read(path) == original
            write_xlsx(path, (x=[2],); overwrite=true)
            @test XLSX.readxlsx(path)[1]["A2"] == 2
            @test_throws ArgumentError write_xlsx(joinpath(dir, "bad.xlsx"), "A" => table, "a" => table)
            for name in ("", "bad/name", "bad:name", "'bad", "bad'", repeat("a", 32))
                @test_throws ArgumentError write_xlsx(joinpath(dir, "bad.xlsx"), table; sheetname=name)
            end
            @test_throws ArgumentError TableStyle(column_widths=Dict(:x => -1))
            @test_throws ArgumentError TableStyle(column_widths=Dict(:x => Inf))
            @test_throws ArgumentError TableStyle(column_styles=["x" => CellStyle(), :x => CellStyle()])
            @test_throws ArgumentError TableStyle(type_formats=["bad" => "0"])
            for anchor in ("A0", "A1:B2", "XFE1", "A1048577", "!A1")
                @test_throws ArgumentError write_xlsx(joinpath(dir, "bad.xlsx"), table; anchor)
            end
            @test_throws ArgumentError write_xlsx(joinpath(dir, "bad.xlsx"), table; anchor="A1048576")
            @test_throws ArgumentError write_xlsx(joinpath(dir, "bad.xlsx"), (x=[1], y=[2]); anchor="XFD1")
            @test_throws ArgumentError write_xlsx(joinpath(dir, "bad.xlsx"), table; column_labels=Dict(:z => "Z"))
            @test_throws ArgumentError write_xlsx(joinpath(dir, "bad.xlsx"), (x=[Inf],))
            @test_throws ArgumentError write_xlsx(joinpath(dir, "bad.xlsx"), (x=[1 + 2im],))
            @test_throws ArgumentError write_xlsx(joinpath(dir, "bad.xlsx"), NamedTuple())
            @test_throws ArgumentError write_xlsx(joinpath(dir, "bad.xlsx"), 42)
            @test_throws ArgumentError write_xlsx(joinpath(dir, "bad.xlsx"), (x=[1], y=[2, 3]))
            @test !isfile(joinpath(dir, "bad.xlsx"))
        end
    end
end

include("widths.jl")
