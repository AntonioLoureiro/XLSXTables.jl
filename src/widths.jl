# Approximate display measurement only. This code never changes cell values or
# number formats, and intentionally does not implement all Excel formatting.
const _WIDTH_FORMAT_NAMES = Dict(
    "General" => "General", "Number" => "0.00", "Currency" => "\$#,##0.00_);(\$#,##0.00)",
    "Percentage" => "0%", "ShortDate" => "m/d/yyyy", "LongDate" => "d-mmm-yy",
    "Time" => "h:mm:ss", "Scientific" => "##0.0E+0",
    "1" => "0", "2" => "0.00", "3" => "#,##0", "4" => "#,##0.00",
    "5" => "\$#,##0_);(\$#,##0)", "6" => "\$#,##0_);[Red](\$#,##0)",
    "7" => "\$#,##0.00_);(\$#,##0.00)", "8" => "\$#,##0.00_);[Red](\$#,##0.00)",
    "9" => "0%", "10" => "0.00%", "11" => "0.00E+00",
    "12" => "# ?/?", "13" => "# ??/??",
    "14" => "m/d/yy", "15" => "d-mmm-yy", "16" => "d-mmm", "17" => "mmm-yy",
    "18" => "h:mm AM/PM", "19" => "h:mm:ss AM/PM", "20" => "h:mm", "21" => "h:mm:ss",
    "22" => "m/d/yy h:mm", "37" => "#,##0;(#,##0)", "38" => "#,##0;[Red](#,##0)",
    "39" => "#,##0.00;(#,##0.00)", "40" => "#,##0.00;[Red](#,##0.00)",
    "45" => "mm:ss", "46" => "[h]:mm:ss", "47" => "mmss.0", "48" => "##0.0E+0", "49" => "@",
)

# Keep literals separate from format codes: a quoted percent or zero must not
# scale a number or change its precision. Split only on unquoted semicolons.
function _width_sections(format::AbstractString)
    chars = collect(get(_WIDTH_FORMAT_NAMES, format, format))
    sections = [Tuple{Symbol,String}[]]
    i = 1
    while i <= length(chars)
        ch = chars[i]
        if ch == '"'
            i += 1
            start = i
            while i <= length(chars) && chars[i] != '"'; i += 1; end
            push!(sections[end], (:literal, join(chars[start:i-1])))
        elseif ch in ('\\', '_', '*') && i < length(chars)
            i += 1
            # Fill characters (*) do not demand a larger width.
            ch != '*' && push!(sections[end], (:literal, ch == '_' ? " " : string(chars[i])))
        elseif ch == '['
            i += 1
            start = i
            while i <= length(chars) && chars[i] != ']'; i += 1; end
            tag = join(chars[start:i-1])
            if occursin(r"^[hHmMsS]+$", tag)
                push!(sections[end], (:elapsed, lowercase(tag)))
            elseif startswith(tag, "\$")
                currency = split(chop(tag; head=1, tail=0), '-'; limit=2)[1]
                push!(sections[end], (:literal, currency))
            elseif occursin(r"^[<>=]", tag)
                push!(sections[end], (:condition, tag))
            end # Colour and locale-only tags have no displayed width.
        elseif ch == ';'
            push!(sections, Tuple{Symbol,String}[])
        elseif i + 4 <= length(chars) && uppercase(join(chars[i:i+4])) == "AM/PM"
            push!(sections[end], (:ampm, "AM/PM")); i += 4
        else
            push!(sections[end], (:code, string(ch)))
        end
        i += 1
    end
    return sections
end

_tokens_text(tokens) = join(t[2] for t in tokens if t[1] != :condition)

function _date_display(value, tokens)
    output = IOBuffer()
    ampm = any(t -> t[1] == :ampm, tokens)
    previous_field = ""
    i = 1
    while i <= length(tokens)
        kind, text = tokens[i]
        if kind == :condition
            i += 1; continue
        elseif kind == :ampm
            print(output, value isa Date || Dates.hour(value) < 12 ? "AM" : "PM")
        elseif kind == :elapsed
            print(output, repeat("0", max(2, length(text))))
        elseif kind == :code && lowercase(text) in ("y", "m", "d", "h", "s")
            field = lowercase(text)
            stop = i
            while stop < length(tokens) && tokens[stop+1][1] == :code && lowercase(tokens[stop+1][2]) == field
                stop += 1
            end
            count = stop - i + 1
            nextfield = findfirst(t -> t[1] == :code && lowercase(t[2]) in ("y", "m", "d", "h", "s"), tokens[stop+1:end])
            following = isnothing(nextfield) ? "" : lowercase(tokens[stop+nextfield][2])
            minute = field == "m" && (value isa Time || previous_field == "h" || following == "s")
            n = if field == "y"
                value isa Time ? 1900 : Dates.year(value)
            elseif field == "d"
                value isa Time ? 1 : Dates.day(value)
            elseif field == "h"
                h = value isa Date ? 0 : Dates.hour(value)
                ampm ? mod(h - 1, 12) + 1 : h
            elseif field == "s"
                value isa Date ? 0 : Dates.second(value)
            elseif minute
                value isa Date ? 0 : Dates.minute(value)
            else
                value isa Time ? 1 : Dates.month(value)
            end
            if field == "m" && !minute && count >= 3
                name = count == 3 ? Dates.monthabbr(n) : Dates.monthname(n)
                print(output, count == 5 ? first(name) : name)
            elseif field == "d" && count >= 3 && !(value isa Time)
                print(output, count == 3 ? Dates.dayabbr(value) : Dates.dayname(value))
            elseif field == "y"
                print(output, count <= 2 ? lpad(mod(n, 100), 2, '0') : lpad(n, count, '0'))
            else
                print(output, lpad(n, count, '0'))
            end
            previous_field = field
            i = stop
        else
            print(output, text)
        end
        i += 1
    end
    return String(take!(output))
end

function _number_display(value::Real, tokens; signed=false)
    placeholder(t) = t[1] == :code && t[2] in ("0", "#", "?")
    firstdigit, lastdigit = findfirst(placeholder, tokens), findlast(placeholder, tokens)
    isnothing(firstdigit) && return _tokens_text(tokens)
    pattern = tokens[firstdigit:lastdigit]
    code = _tokens_text(pattern)
    # Fractions, embedded literals and uncommon masks use a conservative
    # fallback. They remain valid Excel formats and are still written as-is.
    if any(t -> t[1] != :code, pattern) || !occursin(r"^[#0?,]+(?:\.[#0?]*)?(?:[Ee][+-]?[0#?]+)?$", code)
        return string(value, " ", _tokens_text(tokens))
    end
    stop = lastdigit
    while stop < length(tokens) && tokens[stop+1] == (:code, ","); stop += 1; end
    scales = stop - lastdigit
    percents = count(==((:code, "%")), tokens)
    n = abs(Float64(value)) * 100.0^percents / 1000.0^scales
    isfinite(n) || return repeat("0", 256) # Saturate the supported width cap.
    scientific = occursin(r"[Ee]", code)
    mantissa = split(code, r"[Ee]"; limit=2)[1]
    parts = split(mantissa, '.'; limit=2)
    decimals = length(parts) == 2 ? count(c -> c in ('0', '#', '?'), parts[2]) : 0
    # Optional decimal placeholders are reserved conservatively for sizing.
    precision = min(decimals, 30)
    spec = Printf.Format("%." * string(precision) * (scientific ? "E" : "f"))
    numeric = Printf.format(spec, n)
    if scientific
        # Reserve all integer and exponent placeholders, including engineering
        # masks such as ##0.0E+0. Extra space is preferable to clipped digits.
        digits, exponent = split(numeric, 'E'; limit=2)
        mantissa_parts = split(digits, '.'; limit=2)
        whole = lpad(mantissa_parts[1], count(c -> c in ('0', '#', '?'), parts[1]), '0')
        digits = length(mantissa_parts) == 2 ? whole * "." * mantissa_parts[2] : whole
        exponent_mask = split(code, r"[Ee]"; limit=2)[2]
        exponent_digits = lpad(exponent[2:end], count(c -> c in ('0', '#', '?'), exponent_mask), '0')
        numeric = digits * "E" * first(exponent) * exponent_digits
    else
        pieces = split(numeric, '.'; limit=2)
        whole = lpad(pieces[1], count(==('0'), parts[1]), '0')
        if occursin(',', parts[1])
            whole = replace(whole, r"(\d)(?=(\d{3})+(?!\d))" => s"\1,")
        end
        numeric = length(pieces) == 2 ? whole * "." * pieces[2] : whole
    end
    decimals > precision && (numeric *= repeat("0", min(decimals - precision, 256)))
    prefix = _tokens_text(tokens[1:firstdigit-1])
    suffix = _tokens_text(tokens[stop+1:end])
    return (signed && value < 0 ? "-" : "") * prefix * numeric * suffix
end

function _display_candidates(value, sections)
    (ismissing(value) || isnothing(value)) && return [""]
    value isa Bool && return [value ? "TRUE" : "FALSE"]
    if value isa AbstractString
        length(sections) < 4 && return [String(value)]
        return [join(t == (:code, "@") ? value : t[2] for t in sections[4] if t[1] != :condition)]
    end
    isempty(sections) && return [string(value)]
    if value isa Union{Date,DateTime,Time}
        return _tokens_text(sections[1]) in ("General", "@") ? [string(value)] : [_date_display(value, sections[1])]
    elseif value isa Real
        # Conditional number-format sections are measured conservatively;
        # Excel itself determines which condition matches when displaying.
        any(s -> any(t -> t[1] == :condition, s), sections) &&
            return [_number_display(value, s; signed=true) for s in sections[1:min(3, end)]]
        idx = value == 0 && length(sections) >= 3 ? 3 : value < 0 && length(sections) >= 2 ? 2 : 1
        tokens = sections[idx]
        _tokens_text(tokens) in ("General", "@") && return [string(value)]
        return [_number_display(value, tokens; signed=idx == 1)]
    end
    return [string(value)]
end

function _text_units(text::AbstractString, style::CellStyle)
    mono = occursin(r"(?i)mono|courier|consolas|menlo", string(get(style.font, :name, "")))
    size = get(style.font, :size, 11)
    scale = size isa Real && isfinite(size) && size > 0 ? size / 11 : 1.0
    scale *= get(style.font, :bold, false) === true ? 1.08 : 1.0
    widest = 0.0
    for line in split(replace(text, '\r' => ""), '\n')
        width = 0.0
        for ch in line
            units = ch == '\t' ? 4 : textwidth(ch)
            weight = mono || units != 1 ? 1.0 : ch in "MWmw@%" ? 1.4 : ch in "il.,:;!'| `" ? 0.5 : isuppercase(ch) ? 1.15 : 1.0
            width += units * weight
        end
        widest = max(widest, width)
    end
    indent = get(style.alignment, :indent, 0)
    return widest * scale + (indent isa Real ? max(0, indent) * 3 : 0)
end

function _column_layout(values, label, body::CellStyle, style::TableStyle, header::Bool)
    head_width = header ? _text_units(label, style.header) : 0.0
    widest, text_width = head_width, 0.0
    format = body.number_format
    sections = isnothing(format) ? Vector{Tuple{Symbol,String}}[] : _width_sections(format)
    for value in values
        width = maximum(_text_units(s, body) for s in _display_candidates(value, sections))
        widest = max(widest, width)
        value isa AbstractString && (text_width = max(text_width, width))
    end
    width = clamp(widest + style.width_padding, style.min_width, style.max_width)
    wrap_body = style.wrap_long_text && text_width + style.width_padding > style.max_width && !haskey(body.alignment, :wrapText)
    wrap_header = style.wrap_long_text && header && head_width + style.width_padding > style.max_width && !haskey(style.header.alignment, :wrapText)
    return (; width, wrap_body, wrap_header)
end
