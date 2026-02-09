use crate::value::{RuntimeError, Value};

pub fn len_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Len requires exactly one argument".to_string()));
    }
    let s = args[0].as_string();
    Ok(Value::Integer(s.len() as i32))
}

pub fn left_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 2 {
        return Err(RuntimeError::Custom("Left requires exactly two arguments".to_string()));
    }
    let s = args[0].as_string();
    let count = args[1].as_integer()? as usize;
    Ok(Value::String(s.chars().take(count).collect()))
}

pub fn right_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 2 {
        return Err(RuntimeError::Custom("Right requires exactly two arguments".to_string()));
    }
    let s = args[0].as_string();
    let count = args[1].as_integer()? as usize;
    let chars: Vec<char> = s.chars().collect();
    let start = chars.len().saturating_sub(count);
    Ok(Value::String(chars[start..].iter().collect()))
}

pub fn mid_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 2 || args.len() > 3 {
        return Err(RuntimeError::Custom("Mid requires 2 or 3 arguments".to_string()));
    }
    let s = args[0].as_string();
    let start = (args[1].as_integer()? - 1).max(0) as usize; // VB is 1-indexed
    let chars: Vec<char> = s.chars().collect();

    let result = if args.len() == 3 {
        let length = args[2].as_integer()? as usize;
        chars[start.min(chars.len())..].iter().take(length).collect()
    } else {
        chars[start.min(chars.len())..].iter().collect()
    };
    Ok(Value::String(result))
}

pub fn ucase_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("UCase requires exactly one argument".to_string()));
    }
    Ok(Value::String(args[0].as_string().to_uppercase()))
}

pub fn lcase_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("LCase requires exactly one argument".to_string()));
    }
    Ok(Value::String(args[0].as_string().to_lowercase()))
}

pub fn trim_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Trim requires exactly one argument".to_string()));
    }
    Ok(Value::String(args[0].as_string().trim().to_string()))
}

pub fn ltrim_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("LTrim requires exactly one argument".to_string()));
    }
    Ok(Value::String(args[0].as_string().trim_start().to_string()))
}

pub fn rtrim_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("RTrim requires exactly one argument".to_string()));
    }
    Ok(Value::String(args[0].as_string().trim_end().to_string()))
}

/// InStr([start,] string1, string2) - Returns position of first occurrence (1-based), 0 if not found
pub fn instr_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 2 || args.len() > 3 {
        return Err(RuntimeError::Custom("InStr requires 2 or 3 arguments".to_string()));
    }

    let (start, haystack, needle) = if args.len() == 3 {
        let s = (args[0].as_integer()? - 1).max(0) as usize;
        (s, args[1].as_string(), args[2].as_string())
    } else {
        (0, args[0].as_string(), args[1].as_string())
    };

    if needle.is_empty() {
        return Ok(Value::Integer((start + 1) as i32));
    }

    match haystack[start.min(haystack.len())..].find(&needle) {
        Some(pos) => Ok(Value::Integer((start + pos + 1) as i32)), // 1-based
        None => Ok(Value::Integer(0)),
    }
}

/// InStrRev(string1, string2[, start]) - Returns position of last occurrence (1-based)
pub fn instrrev_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 2 || args.len() > 3 {
        return Err(RuntimeError::Custom("InStrRev requires 2 or 3 arguments".to_string()));
    }

    let haystack = args[0].as_string();
    let needle = args[1].as_string();
    let end = if args.len() == 3 {
        args[2].as_integer()? as usize
    } else {
        haystack.len()
    };

    if needle.is_empty() {
        return Ok(Value::Integer(end as i32));
    }

    match haystack[..end.min(haystack.len())].rfind(&needle) {
        Some(pos) => Ok(Value::Integer((pos + 1) as i32)), // 1-based
        None => Ok(Value::Integer(0)),
    }
}

/// Replace(string, find, replacement[, start[, count]]) - Replace occurrences
pub fn replace_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 3 || args.len() > 5 {
        return Err(RuntimeError::Custom("Replace requires 3 to 5 arguments".to_string()));
    }

    let source = args[0].as_string();
    let find = args[1].as_string();
    let replacement = args[2].as_string();

    let start = if args.len() >= 4 {
        (args[3].as_integer()? - 1).max(0) as usize
    } else {
        0
    };

    let count = if args.len() >= 5 {
        args[4].as_integer()?
    } else {
        -1
    };

    let working = if start > 0 {
        source[start.min(source.len())..].to_string()
    } else {
        source
    };

    let result = if count < 0 {
        working.replace(&find, &replacement)
    } else {
        working.replacen(&find, &replacement, count as usize)
    };

    Ok(Value::String(result))
}

/// Chr(charcode) - Returns character for ASCII code
pub fn chr_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Chr requires exactly one argument".to_string()));
    }
    let code = args[0].as_integer()?;
    if code < 0 || code > 127 {
        // Support extended range via char::from_u32
        if let Some(c) = char::from_u32(code as u32) {
            return Ok(Value::String(c.to_string()));
        }
        return Err(RuntimeError::Custom(format!("Invalid character code: {}", code)));
    }
    Ok(Value::String((code as u8 as char).to_string()))
}

/// Asc(string) - Returns ASCII code for first character
pub fn asc_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Asc requires exactly one argument".to_string()));
    }
    let s = args[0].as_string();
    if s.is_empty() {
        return Err(RuntimeError::Custom("Asc: string must not be empty".to_string()));
    }
    Ok(Value::Integer(s.chars().next().unwrap() as i32))
}

/// Split(expression[, delimiter]) - Splits string into array
pub fn split_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() || args.len() > 2 {
        return Err(RuntimeError::Custom("Split requires 1 or 2 arguments".to_string()));
    }
    let s = args[0].as_string();
    let delim = if args.len() == 2 {
        args[1].as_string()
    } else {
        " ".to_string()
    };

    let parts: Vec<Value> = s.split(&delim).map(|p| Value::String(p.to_string())).collect();
    Ok(Value::Array(parts))
}

/// Join(array[, delimiter]) - Joins array elements into string
pub fn join_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() || args.len() > 2 {
        return Err(RuntimeError::Custom("Join requires 1 or 2 arguments".to_string()));
    }
    let arr = match &args[0] {
        Value::Array(a) => a,
        _ => return Err(RuntimeError::TypeError {
            expected: "Array".to_string(),
            got: format!("{:?}", args[0]),
        }),
    };

    let delim = if args.len() == 2 {
        args[1].as_string()
    } else {
        " ".to_string()
    };

    let parts: Vec<String> = arr.iter().map(|v| v.as_string()).collect();
    Ok(Value::String(parts.join(&delim)))
}

/// StrReverse(string) - Reverses a string
pub fn strreverse_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("StrReverse requires exactly one argument".to_string()));
    }
    Ok(Value::String(args[0].as_string().chars().rev().collect()))
}

/// Space(number) - Returns a string of spaces
pub fn space_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Space requires exactly one argument".to_string()));
    }
    let count = args[0].as_integer()?.max(0) as usize;
    Ok(Value::String(" ".repeat(count)))
}

/// String(number, character) - Returns repeated character
pub fn string_repeat_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 2 {
        return Err(RuntimeError::Custom("String requires exactly two arguments".to_string()));
    }
    let count = args[0].as_integer()?.max(0) as usize;
    let ch = args[1].as_string();
    let c = ch.chars().next().unwrap_or(' ');
    Ok(Value::String(c.to_string().repeat(count)))
}

/// StrComp(string1, string2[, compare]) - Compares two strings
/// Returns -1, 0, or 1
pub fn strcomp_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 2 || args.len() > 3 {
        return Err(RuntimeError::Custom("StrComp requires 2 or 3 arguments".to_string()));
    }
    let s1 = args[0].as_string();
    let s2 = args[1].as_string();

    // compare mode: 0 = binary (default), 1 = text (case-insensitive)
    let text_compare = args.get(2).map(|v| v.as_integer().unwrap_or(0) == 1).unwrap_or(false);

    let result = if text_compare {
        s1.to_lowercase().cmp(&s2.to_lowercase())
    } else {
        s1.cmp(&s2)
    };

    Ok(Value::Integer(match result {
        std::cmp::Ordering::Less => -1,
        std::cmp::Ordering::Equal => 0,
        std::cmp::Ordering::Greater => 1,
    }))
}

/// Format(expression, format_string) - Basic formatting
pub fn format_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 1 || args.len() > 2 {
        return Err(RuntimeError::Custom("Format requires 1 or 2 arguments".to_string()));
    }

    if args.len() == 1 {
        return Ok(Value::String(args[0].as_string()));
    }

    let fmt_str = args[1].as_string().to_lowercase();

    match fmt_str.as_str() {
        "currency" | "c" => {
            let val = args[0].as_double()?;
            Ok(Value::String(format!("${:.2}", val)))
        }
        "fixed" | "f" => {
            let val = args[0].as_double()?;
            Ok(Value::String(format!("{:.2}", val)))
        }
        "standard" | "n" => {
            let val = args[0].as_double()?;
            // Add thousands separator
            let formatted = format!("{:.2}", val);
            let parts: Vec<&str> = formatted.split('.').collect();
            let int_part = parts[0];
            let dec_part = parts.get(1).unwrap_or(&"00");
            let chars: Vec<char> = int_part.chars().rev().collect();
            let mut with_commas = String::new();
            for (i, c) in chars.iter().enumerate() {
                if i > 0 && i % 3 == 0 && *c != '-' {
                    with_commas.push(',');
                }
                with_commas.push(*c);
            }
            let int_formatted: String = with_commas.chars().rev().collect();
            Ok(Value::String(format!("{}.{}", int_formatted, dec_part)))
        }
        "percent" | "p" => {
            let val = args[0].as_double()?;
            Ok(Value::String(format!("{:.2}%", val * 100.0)))
        }
        "yes/no" => {
            let val = args[0].as_bool()?;
            Ok(Value::String(if val { "Yes" } else { "No" }.to_string()))
        }
        "true/false" => {
            let val = args[0].as_bool()?;
            Ok(Value::String(if val { "True" } else { "False" }.to_string()))
        }
        "on/off" => {
            let val = args[0].as_bool()?;
            Ok(Value::String(if val { "On" } else { "Off" }.to_string()))
        }
        _ => {
            // For numeric format strings like "0.00", "#,##0", etc.
            // Basic handling: count decimal places
            if fmt_str.contains('.') {
                let dec_places = fmt_str.split('.').nth(1).map(|s| s.len()).unwrap_or(0);
                let val = args[0].as_double()?;
                Ok(Value::String(format!("{:.prec$}", val, prec = dec_places)))
            } else {
                Ok(Value::String(args[0].as_string()))
            }
        }
    }
}


