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
    Integer => "0",
    AbstractFloat => "#,##0.00",
    Date => "yyyy-mm-dd",
    DateTime => "yyyy-mm-dd hh:mm:ss",
    Time => "hh:mm",
    AbstractString => "@",
]

"""
    TableStyle(; header, body, type_formats, column_styles, column_widths,
                 auto_width=true, min_width=10, max_width=50, width_padding=2,
                 wrap_long_text=true)

Configure a whole table. Data cells receive, in order: `body`, the first
matching `type_formats` rule, then the named `column_styles` override.
Headers use `header` independently. Column keys refer to the original column
names and can be strings or symbols. Widths use Excel character units and
affect the entire worksheet column. Automatic widths estimate formatted values
and headers, with padding and limits. Explicit `column_widths` always win.
`wrap_long_text` wraps text exceeding the automatic width cap unless `wrapText`
was explicitly configured. Sizing approximates Excel AutoFit; it does not use
an Excel rendering engine. Disable with `auto_width=false`.

Default headers use Calibri 11, bold white text, a solid orange (#FD5108)
background, centered horizontal/vertical alignment, and wrapped text. Default
data cells use Calibri 11 and a solid light-grey (#F2F2F2) background.
"""
struct TableStyle
    header::CellStyle
    body::CellStyle
    type_formats::Vector{Pair{Type,String}}
    column_styles::Dict{Symbol,CellStyle}
    column_widths::Dict{Symbol,Float64}
    auto_width::Bool
    min_width::Float64
    max_width::Float64
    width_padding::Float64
    wrap_long_text::Bool
end

function TableStyle(;
    header::CellStyle=CellStyle(
        font=(name="Calibri", size=11, bold=true, color="FFFFFFFF"),
        fill=(pattern="solid", fgColor="FFFD5108"),
        alignment=(horizontal="center", vertical="center", wrapText=true),
    ),
    body::CellStyle=CellStyle(
        font=(name="Calibri", size=11),
        fill=(pattern="solid", fgColor="FFF2F2F2"),
    ),
    type_formats=default_type_formats(),
    column_styles=Dict{Symbol,CellStyle}(),
    column_widths=Dict{Symbol,Float64}(),
    auto_width::Bool=true,
    min_width::Real=10,
    max_width::Real=50,
    width_padding::Real=2,
    wrap_long_text::Bool=true,
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
    all(isfinite, (min_width, max_width, width_padding, Float64(width_padding))) &&
        0 <= min_width <= max_width <= 255 && width_padding >= 0 ||
        throw(ArgumentError("width limits must satisfy 0 <= min_width <= max_width <= 255; padding must be finite and nonnegative"))
    return TableStyle(header, body, rules, overrides, widths, auto_width,
        Float64(min_width), Float64(max_width), Float64(width_padding), wrap_long_text)
end

# Preserve the original positional constructor, applying the same validation.
TableStyle(header::CellStyle, body::CellStyle, type_formats, column_styles, column_widths) =
    TableStyle(; header, body, type_formats, column_styles, column_widths)

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
