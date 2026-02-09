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
