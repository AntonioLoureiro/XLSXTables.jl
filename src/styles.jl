"""
    CellStyle(; number_format=nothing, font=(;), fill=(;), border=(;), alignment=(;))

Reusable cell formatting. The named tuples accept the keyword arguments of
`XLSX.setFont`, `setFill`, `setBorder`, and `setAlignment`, respectively.
Unspecified attributes inherit the previous formatting layer. Set a number
format to `"General"` to explicitly reset it.
"""
Base.@kwdef struct CellStyle
    number_format::Union{Nothing,String} = nothing
    font::NamedTuple = (;)
    fill::NamedTuple = (;)
    border::NamedTuple = (;)
    alignment::NamedTuple = (;)
end

"""
    default_type_formats()

Return a fresh, ordered list of Julia type => Excel number format rules.
The first matching rule wins, so put specific types before abstract types.
`Missing` and `Nothing` are ignored when inferring a column's type.
"""
default_type_formats() = Pair{Type,String}[
    Bool => "General",
    Integer => "#,##0",
    AbstractFloat => "#,##0.00",
    Date => "yyyy-mm-dd",
    DateTime => "yyyy-mm-dd hh:mm:ss",
    Time => "hh:mm:ss",
    AbstractString => "@",
]

"""
    TableStyle(; header, body, type_formats, column_styles, column_widths)

Configure a whole table. Data cells receive, in order: `body`, the first
matching `type_formats` rule, then the named `column_styles` override.
Headers use `header` independently. Column keys refer to the original column
names and can be strings or symbols. Widths use Excel character units and
affect the entire worksheet column.
"""
struct TableStyle
    header::CellStyle
    body::CellStyle
    type_formats::Vector{Pair{Type,String}}
    column_styles::Dict{Symbol,CellStyle}
    column_widths::Dict{Symbol,Float64}
end

function TableStyle(;
    header::CellStyle=CellStyle(
        font=(bold=true, color="FFFFFFFF"),
        fill=(pattern="solid", fgColor="FF24476B"),
        alignment=(vertical="center", wrapText=true),
    ),
    body::CellStyle=CellStyle(font=(name="Calibri", size=11)),
    type_formats=default_type_formats(),
    column_styles=Dict{Symbol,CellStyle}(),
    column_widths=Dict{Symbol,Float64}(),
)
    rules = Pair{Type,String}[]
    for (T, fmt) in type_formats
        T isa Type || throw(ArgumentError("type_formats keys must be Julia types"))
        fmt isa AbstractString || throw(ArgumentError("number formats must be strings"))
        push!(rules, T => String(fmt))
    end
    overrides = _named_dict(CellStyle, column_styles)
    widths = _named_dict(Float64, column_widths)
    all(w -> isfinite(w) && 0 <= w <= 255, values(widths)) ||
        throw(ArgumentError("column widths must be finite and between 0 and 255"))
    return TableStyle(header, body, rules, overrides, widths)
end

function _named_dict(::Type{V}, entries) where V
    result = Dict{Symbol,V}()
    for (key, value) in pairs_or_entries(entries)
        key isa Union{Symbol,AbstractString} ||
            throw(ArgumentError("column keys must be strings or symbols"))
        name = Symbol(key)
        haskey(result, name) && throw(ArgumentError("duplicate column key: $name"))
        result[name] = value
    end
    return result
end

pairs_or_entries(entries::Union{AbstractDict,NamedTuple}) = pairs(entries)
pairs_or_entries(entries) = entries

function _merge_style(base::CellStyle, override::CellStyle)
    return CellStyle(
        number_format=isnothing(override.number_format) ? base.number_format : override.number_format,
        font=merge(base.font, override.font),
        fill=merge(base.fill, override.fill),
        border=merge(base.border, override.border),
        alignment=merge(base.alignment, override.alignment),
    )
end

function _column_type(column)
    T = Base.nonmissingtype(eltype(column))
    T === Nothing && (T = Union{})
    # For Any/abstract/mixed columns infer from the actual, nonblank values.
    if T === Union{} || !isconcretetype(T)
        T = Union{}
        for value in column
            (ismissing(value) || isnothing(value)) && continue
            T = typejoin(T, typeof(value))
        end
    end
    return T
end

function _data_style(style::TableStyle, name::Symbol, column)
    result = style.body
    T = _column_type(column)
    if T !== Union{}
        for (rule, fmt) in style.type_formats
            if T <: rule
                result = _merge_style(result, CellStyle(number_format=fmt))
                break
            end
        end
    end
    return _merge_style(result, get(style.column_styles, name, CellStyle()))
end

function _apply_style!(sheet, range::String, style::CellStyle)
    isnothing(style.number_format) || XLSX.setFormat(sheet, range; format=style.number_format)
    isempty(style.font) || XLSX.setFont(sheet, range; style.font...)
    isempty(style.fill) || XLSX.setFill(sheet, range; style.fill...)
    isempty(style.border) || XLSX.setBorder(sheet, range; style.border...)
    isempty(style.alignment) || XLSX.setAlignment(sheet, range; style.alignment...)
    return nothing
end
