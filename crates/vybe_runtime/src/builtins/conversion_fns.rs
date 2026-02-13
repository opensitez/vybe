use crate::value::{RuntimeError, Value};
use chrono::{NaiveDate, NaiveTime, NaiveDateTime};

fn date_to_ole(dt: NaiveDateTime) -> f64 {
    let base_date = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap().and_hms_opt(0, 0, 0).unwrap();
    let duration = dt.signed_duration_since(base_date);
    let days = duration.num_days() as f64;
    let seconds = (duration.num_seconds() % 86400) as f64;
    days + (seconds / 86400.0)
}

fn parse_date(s: &str) -> Option<f64> {
    let s = s.trim();
    
    // Try full datetime formats
    let formats = [
        "%m/%d/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%Y-%m-%d %H:%M",
        "%m/%d/%Y",
        "%Y-%m-%d",
        "%H:%M:%S",
        "%H:%M"
    ];

    for fmt in formats {
        if let Ok(dt) = NaiveDateTime::parse_from_str(s, fmt) {
            return Some(date_to_ole(dt));
        }
        if let Ok(d) = NaiveDate::parse_from_str(s, fmt) {
            return Some(date_to_ole(d.and_hms_opt(0, 0, 0).unwrap()));
        }
        if let Ok(t) = NaiveTime::parse_from_str(s, fmt) {
            let base_date = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
            return Some(date_to_ole(base_date.and_time(t)));
        }
    }
    None
}

pub fn cstr_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CStr requires exactly one argument".to_string()));
    }
    Ok(Value::String(args[0].as_string()))
}

pub fn cint_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CInt requires exactly one argument".to_string()));
    }
    let val = args[0].as_integer()?;
    Ok(Value::Integer(val))
}

pub fn cdbl_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CDbl requires exactly one argument".to_string()));
    }
    let val = args[0].as_double()?;
    Ok(Value::Double(val))
}

pub fn cbool_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CBool requires exactly one argument".to_string()));
    }
    let val = args[0].as_bool()?;
    Ok(Value::Boolean(val))
}

/// CLng(expression) - Convert to Long
pub fn clng_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CLng requires exactly one argument".to_string()));
    }
    let val = args[0].as_long()?;
    Ok(Value::Long(val))
}

pub fn cbyte_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CByte requires exactly one argument".to_string()));
    }
    let val = args[0].as_byte()?;
    Ok(Value::Byte(val))
}

pub fn cchar_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CChar requires exactly one argument".to_string()));
    }
    let val = args[0].as_char()?;
    Ok(Value::Char(val))
}

/// CSng(expression) - Convert to Single
pub fn csng_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CSng requires exactly one argument".to_string()));
    }
    let val = args[0].as_double()? as f32;
    Ok(Value::Single(val))
}

/// Val(string) - Converts string to number, stops at first non-numeric character
pub fn val_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Val requires exactly one argument".to_string()));
    }
    let s = args[0].as_string();
    let trimmed = s.trim();

    if trimmed.is_empty() {
        return Ok(Value::Double(0.0));
    }

    // Val stops at first non-numeric character (but allows leading whitespace, sign, decimal)
    let mut end = 0;
    let mut has_dot = false;
    let chars: Vec<char> = trimmed.chars().collect();

    if !chars.is_empty() && (chars[0] == '-' || chars[0] == '+') {
        end = 1;
    }

    while end < chars.len() {
        if chars[end].is_ascii_digit() {
            end += 1;
        } else if chars[end] == '.' && !has_dot {
            has_dot = true;
            end += 1;
        } else {
            break;
        }
    }

    if end == 0 || (end == 1 && (chars[0] == '-' || chars[0] == '+')) {
        return Ok(Value::Double(0.0));
    }

    let num_str: String = chars[..end].iter().collect();
    let val: f64 = num_str.parse().unwrap_or(0.0);
    Ok(Value::Double(val))
}

/// Str(number) - Converts number to string (with leading space for positive)
pub fn str_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Str requires exactly one argument".to_string()));
    }
    let s = args[0].as_string();
    // VB Str() adds a leading space for positive numbers
    match &args[0] {
        Value::Integer(i) if *i >= 0 => Ok(Value::String(format!(" {}", i))),
        Value::Long(l) if *l >= 0 => Ok(Value::String(format!(" {}", l))),
        Value::Single(f) if *f >= 0.0 => Ok(Value::String(format!(" {}", f))),
        Value::Double(d) if *d >= 0.0 => Ok(Value::String(format!(" {}", d))),
        _ => Ok(Value::String(s)),
    }
}

/// Hex(number) - Converts to hexadecimal string
pub fn hex_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Hex requires exactly one argument".to_string()));
    }
    let val = args[0].as_integer()?;
    Ok(Value::String(format!("{:X}", val)))
}

/// Oct(number) - Converts to octal string
pub fn oct_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Oct requires exactly one argument".to_string()));
    }
    let val = args[0].as_integer()?;
    Ok(Value::String(format!("{:o}", val)))
}



/// CDate(expression) - Converts to date
pub fn cdate_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CDate requires exactly one argument".to_string()));
    }
    match &args[0] {
        Value::Date(d) => Ok(Value::Date(*d)),
        Value::String(s) => {
            if let Some(d) = parse_date(s) {
                Ok(Value::Date(d))
            } else {
                Err(RuntimeError::Custom(format!("Type mismatch: cannot convert '{}' to Date", s)))
            }
        },
        Value::Double(d) => Ok(Value::Date(*d)),
        Value::Integer(i) => Ok(Value::Date(*i as f64)),
        _ => Err(RuntimeError::Custom("Type mismatch".to_string()))
    }
}

/// CDec(expression) - Converts to Decimal (maps to Double)
pub fn cdec_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CDec requires exactly one argument".to_string()));
    }
    Ok(Value::Double(args[0].as_double()?))
}

/// Asc(string) returns numeric code; ChrW(code) returns Unicode char
/// FormatNumber(number[, decimals]) - Formats number with decimal places
pub fn formatnumber_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() || args.len() > 2 {
        return Err(RuntimeError::Custom("FormatNumber requires 1 or 2 arguments".to_string()));
    }
    let d = args[0].as_double()?;
    let decimals = if args.len() >= 2 { args[1].as_integer()?.max(0) as usize } else { 2 };
    Ok(Value::String(format!("{:.prec$}", d, prec = decimals)))
}

/// FormatCurrency(number[, decimals]) - Formats as currency
pub fn formatcurrency_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() || args.len() > 2 {
        return Err(RuntimeError::Custom("FormatCurrency requires 1 or 2 arguments".to_string()));
    }
    let d = args[0].as_double()?;
    let decimals = if args.len() >= 2 { args[1].as_integer()?.max(0) as usize } else { 2 };
    Ok(Value::String(format!("${:.prec$}", d, prec = decimals)))
}

/// FormatPercent(number[, decimals]) - Formats as percentage
pub fn formatpercent_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() || args.len() > 2 {
        return Err(RuntimeError::Custom("FormatPercent requires 1 or 2 arguments".to_string()));
    }
    let d = args[0].as_double()?;
    let decimals = if args.len() >= 2 { args[1].as_integer()?.max(0) as usize } else { 2 };
    Ok(Value::String(format!("{:.prec$}%", d * 100.0, prec = decimals)))
}

/// CObj(expression) - Convert to Object
pub fn cobj_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CObj requires exactly one argument".to_string()));
    }
    // In VB, CObj just returns the value as an Object type
    Ok(args[0].clone())
}

/// CShort(expression) - Convert to Short (16-bit integer)
pub fn cshort_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CShort requires exactly one argument".to_string()));
    }
    let val = args[0].as_integer()? as i16;
    Ok(Value::Integer(val as i32))
}

/// CUShort(expression) - Convert to UShort (16-bit unsigned integer)
pub fn cushort_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CUShort requires exactly one argument".to_string()));
    }
    let val = args[0].as_integer()?.max(0) as u16;
    Ok(Value::Integer(val as i32))
}

/// CUInt(expression) - Convert to UInteger (32-bit unsigned integer)
pub fn cuint_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CUInt requires exactly one argument".to_string()));
    }
    let val = args[0].as_integer()?.max(0) as u32;
    Ok(Value::Long(val as i64))
}

/// CULng(expression) - Convert to ULong (64-bit unsigned integer)
pub fn culng_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CULng requires exactly one argument".to_string()));
    }
    let val = args[0].as_long()?.max(0) as u64;
    Ok(Value::Long(val as i64))
}

/// CSByte(expression) - Convert to Signed Byte (-128 to 127)
pub fn csbyte_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CSByte requires exactly one argument".to_string()));
    }
    let val = args[0].as_integer()?;
    if val < i8::MIN as i32 || val > i8::MAX as i32 {
        return Err(RuntimeError::Custom("Overflow in CSByte conversion".to_string()));
    }
    Ok(Value::Integer(val))
}

/// AscW(string) - Returns Unicode code point of first character
pub fn ascw_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("AscW requires exactly one argument".to_string()));
    }
    let s = args[0].as_string();
    if let Some(ch) = s.chars().next() {
        Ok(Value::Integer(ch as u32 as i32))
    } else {
        Err(RuntimeError::Custom("AscW: empty string".to_string()))
    }
}

/// ChrW(code) - Returns Unicode character from code point
pub fn chrw_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("ChrW requires exactly one argument".to_string()));
    }
    let code = args[0].as_integer()? as u32;
    if let Some(ch) = char::from_u32(code) {
        Ok(Value::String(ch.to_string()))
    } else {
        Err(RuntimeError::Custom(format!("ChrW: invalid Unicode code point {}", code)))
    }
}

/// RGB(red, green, blue) - Returns color value from RGB components
pub fn rgb_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 3 {
        return Err(RuntimeError::Custom("RGB requires exactly 3 arguments".to_string()));
    }
    let r = (args[0].as_integer()? & 0xFF) as u32;
    let g = (args[1].as_integer()? & 0xFF) as u32;
    let b = (args[2].as_integer()? & 0xFF) as u32;
    
    // VB6 color format: &H00BBGGRR (blue in high byte, red in low byte)
    let color = (b << 16) | (g << 8) | r;
    Ok(Value::Integer(color as i32))
}

/// QBColor(color_number) - Returns color from QB color number (0-15)
pub fn qbcolor_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("QBColor requires exactly one argument".to_string()));
    }
    let color_num = args[0].as_integer()?;
    
    // QBasic color palette (16 colors)
    let rgb = match color_num {
        0 => (0, 0, 0),         // Black
        1 => (0, 0, 128),       // Blue
        2 => (0, 128, 0),       // Green
        3 => (0, 128, 128),     // Cyan
        4 => (128, 0, 0),       // Red
        5 => (128, 0, 128),     // Magenta
        6 => (128, 128, 0),     // Brown/Yellow
        7 => (192, 192, 192),   // Light Gray
        8 => (128, 128, 128),   // Dark Gray
        9 => (0, 0, 255),       // Bright Blue
        10 => (0, 255, 0),      // Bright Green
        11 => (0, 255, 255),    // Bright Cyan
        12 => (255, 0, 0),      // Bright Red
        13 => (255, 0, 255),    // Bright Magenta
        14 => (255, 255, 0),    // Bright Yellow
        15 => (255, 255, 255),  // White
        _ => return Err(RuntimeError::Custom(format!("QBColor: invalid color number {}", color_num))),
    };
    
    // Return as VB6 color format
    let color = (rgb.2 << 16) | (rgb.1 << 8) | rgb.0;
    Ok(Value::Integer(color as i32))
}

/// CCur(expression) - Convert to Currency type (stored as Double with 4 decimal places)
pub fn ccur_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CCur requires exactly one argument".to_string()));
    }
    // Currency in VB is fixed-point decimal with 4 decimal places
    // We'll use Double and round to 4 decimal places
    let val = args[0].as_double()?;
    let rounded = (val * 10000.0).round() / 10000.0;
    Ok(Value::Double(rounded))
}

/// CVar(expression) - Convert to Variant (just returns the value as-is)
pub fn cvar_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("CVar requires exactly one argument".to_string()));
    }
    // In VB, CVar converts to Variant type
    // Since our Value enum is already variant-like, just return as-is
    Ok(args[0].clone())
}
