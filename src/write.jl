const MAX_ROWS = 1_048_576
const MAX_COLS = 16_384

function _anchor(anchor::AbstractString)
    m = match(r"^([A-Za-z]{1,3})([1-9][0-9]{0,6})$", anchor)
    isnothing(m) && throw(ArgumentError("anchor must be a cell address such as A1 or C5"))
    col = 0
    for ch in uppercase(m[1])
        col = 26col + Int(ch - 'A') + 1
    end
    row = parse(Int, m[2])
    row <= MAX_ROWS && col <= MAX_COLS || throw(ArgumentError("anchor exceeds Excel limits"))
    return row, col
end

function _column_name(col::Int)
    result = ""
    while col > 0
        col, digit = divrem(col - 1, 26)
        result = string(Char(Int('A') + digit), result)
    end
    return result
end

_cell(row, col) = string(_column_name(col), row)
_range(r1, c1, r2, c2) = string(_cell(r1, c1), ":", _cell(r2, c2))

function _validate_keys(entries, names, option)
    unknown = setdiff(collect(keys(entries)), names)
    isempty(unknown) || throw(ArgumentError("unknown columns in $option: $(join(unknown, ", "))"))
end

# Normalise values explicitly, since XLSX otherwise stringifies some numeric types.
_excel_value(::Nothing) = missing
_excel_value(x::Union{Missing,Date,DateTime,Time}) = x
_excel_value(x::Bool) = x
_excel_value(x::AbstractString) = String(x)
_excel_value(x::Symbol) = String(x)
function _excel_value(x::Integer)
    # Excel numeric cells retain only about 15 significant decimal digits.
    # Preserve larger identifiers exactly as text instead of silently rounding.
    return abs(big(x)) > 999_999_999_999_999 ? string(x) : Int64(x)
end
function _excel_value(x::Real)
    value = Float64(x)
    isfinite(value) || throw(ArgumentError("Excel numeric cells cannot contain NaN or Inf"))
    return value
end
_excel_value(x) = throw(ArgumentError("unsupported Excel value of type $(typeof(x)); convert it to a string or supported scalar first"))

function _prepare(table, anchor, header, labels, style)
    Tables.istable(table) || throw(ArgumentError("table must implement the Tables.jl interface"))
    columns = Tables.columns(table)
    names = Symbol.(collect(Tables.columnnames(columns)))
    isempty(names) && throw(ArgumentError("a table must have at least one column"))
    length(unique(names)) == length(names) || throw(ArgumentError("column names must be unique"))
    data = [collect(Tables.getcolumn(columns, name)) for name in names]
    nrows = length(first(data))
    all(col -> length(col) == nrows, data) || throw(ArgumentError("columns must have equal lengths"))
    row, col = _anchor(anchor)
    row + nrows + Int(header) - 1 <= MAX_ROWS || throw(ArgumentError("table exceeds Excel's row limit"))
    col + length(names) - 1 <= MAX_COLS || throw(ArgumentError("table exceeds Excel's column limit"))
    labelmap = _named_dict(String, labels)
    _validate_keys(labelmap, names, "column_labels")
    _validate_keys(style.column_styles, names, "column_styles")
    _validate_keys(style.column_widths, names, "column_widths")
    styles = [_data_style(style, name, data[i]) for (i, name) in enumerate(names)]
    # Validate and convert all cells before mutating a supplied worksheet.
    values = [map(_excel_value, column) for column in data]
    return (; names, values, styles, labels=[get(labelmap, n, string(n)) for n in names],
        row, col, nrows, header, style)
end

function _write_prepared!(sheet::XLSX.Worksheet, prepared)
    (; names, values, styles, labels, row, col, nrows, header, style) = prepared
    for (j, column) in enumerate(values)
        target_col = col + j - 1
        header && (sheet[row, target_col] = labels[j])
        for (i, value) in enumerate(column)
            sheet[row + Int(header) + i - 1, target_col] = value
        end
        if nrows > 0
            _apply_style!(sheet, _range(row + Int(header), target_col,
                row + Int(header) + nrows - 1, target_col), styles[j])
        end
        if haskey(style.column_widths, names[j])
            XLSX.setColumnWidth(sheet, _column_name(target_col); width=style.column_widths[names[j]])
        end
    end
    if header
        _apply_style!(sheet, _range(row, col, row, col + length(names) - 1), style.header)
    end
    return sheet
end

"""
    write_table!(sheet, table; anchor="A1", header=true, column_labels=Dict(), style=TableStyle())

Write a Tables.jl source into an existing writable XLSX worksheet and return
that worksheet. Data outside the destination rectangle is left alone; existing
cells inside it are overwritten. Column widths, when provided, affect the whole
worksheet column. Empty tables with known columns can still write headers.
This writes formatted cells, not a native Excel Table object.
"""
function write_table!(sheet::XLSX.Worksheet, table;
    anchor::AbstractString="A1", header::Bool=true,
    column_labels=Dict{Symbol,String}(), style::TableStyle=TableStyle(),
)
    return _write_prepared!(sheet, _prepare(table, anchor, header, column_labels, style))
end

function _sheet_name(name)
    name isa Union{Symbol,AbstractString} || throw(ArgumentError("sheet names must be strings or symbols"))
    s = String(name)
    isempty(s) && throw(ArgumentError("sheet name cannot be empty"))
    # Excel measures the 31-character limit in UTF-16 code units.
    length(transcode(UInt16, s)) <= 31 || throw(ArgumentError("sheet name exceeds 31 characters"))
    occursin(r"[\\/\?\*\[\]:\x00-\x1f]", s) && throw(ArgumentError("invalid character in sheet name: $s"))
    (startswith(s, "'") || endswith(s, "'")) && throw(ArgumentError("sheet names cannot start or end with an apostrophe"))
    return s
end

"""
    write_xlsx(path, table; sheetname="Sheet1", overwrite=false, kwargs...)
    write_xlsx(path, "Sheet A" => table_a, "Sheet B" => table_b; overwrite=false, kwargs...)

Create an Excel workbook with one or more sheets. Formatting keywords are the
same as `write_table!` and apply to each table. For individual sheet settings,
use `write_table!` with an XLSX workbook directly. Returns `path` as a string.
Existing files are protected unless `overwrite=true`. The complete workbook is
written to a temporary file in the destination directory before replacement.
"""
function write_xlsx(path::AbstractString, table;
    sheetname::Union{Symbol,AbstractString}="Sheet1", overwrite::Bool=false, kwargs...,
)
    return write_xlsx(path, sheetname => table; overwrite, kwargs...)
end

function write_xlsx(path::AbstractString, first_sheet::Pair, other_sheets::Pair...;
    overwrite::Bool=false, anchor::AbstractString="A1", header::Bool=true,
    column_labels=Dict{Symbol,String}(), style::TableStyle=TableStyle(),
)
    destination = abspath(path)
    ispath(destination) && !overwrite && throw(ArgumentError("file already exists; pass overwrite=true: $path"))
    isdir(destination) && throw(ArgumentError("destination is a directory: $path"))
    sheets = (first_sheet, other_sheets...)
    names = [_sheet_name(first(pair)) for pair in sheets]
    length(unique(lowercase.(names))) == length(names) ||
        throw(ArgumentError("sheet names must be unique (case-insensitive)"))
    prepared = [_prepare(last(pair), anchor, header, column_labels, style) for pair in sheets]
    workbook = XLSX.newxlsx()
    for (i, name) in enumerate(names)
        sheet = i == 1 ? workbook[1] : XLSX.addsheet!(workbook, name)
        i == 1 && XLSX.renamesheet!(sheet, name)
        _write_prepared!(sheet, prepared[i])
    end
    temporary, io = mktemp(dirname(destination))
    close(io)
    try
        XLSX.writexlsx(temporary, workbook; overwrite=true)
        mv(temporary, destination; force=overwrite)
    finally
        isfile(temporary) && rm(temporary)
    end
    return String(path)
end
