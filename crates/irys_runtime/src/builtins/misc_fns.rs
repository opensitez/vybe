use crate::value::{RuntimeError, Value};

/// DoEvents() - Yields to the event loop (stub: no-op in interpreter)
pub fn doevents_fn(_args: &[Value]) -> Result<Value, RuntimeError> {
    Ok(Value::Nothing)
}

/// RGB(red, green, blue) - Returns a color value
pub fn rgb_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 3 {
        return Err(RuntimeError::Custom("RGB requires exactly 3 arguments".to_string()));
    }
    let r = args[0].as_integer()?.clamp(0, 255);
    let g = args[1].as_integer()?.clamp(0, 255);
    let b = args[2].as_integer()?.clamp(0, 255);
    Ok(Value::Long((r as i64) | ((g as i64) << 8) | ((b as i64) << 16)))
}

/// Environ(name) - Returns environment variable value
pub fn environ_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Environ requires exactly one argument".to_string()));
    }
    let name = args[0].as_string();
    let val = std::env::var(&name).unwrap_or_default();
    Ok(Value::String(val))
}

/// QBColor(color_code) - Returns a color value from VB color codes 0-15
pub fn qbcolor_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("QBColor requires exactly one argument".to_string()));
    }
    let code = args[0].as_integer()?;
    let color: i64 = match code {
        0 => 0x000000,   // Black
        1 => 0x800000,   // Blue
        2 => 0x008000,   // Green
        3 => 0x808000,   // Cyan
        4 => 0x000080,   // Red
        5 => 0x800080,   // Magenta
        6 => 0x008080,   // Yellow
        7 => 0xC0C0C0,   // White
        8 => 0x808080,   // Gray
        9 => 0xFF0000,   // Light Blue
        10 => 0x00FF00,  // Light Green
        11 => 0xFFFF00,  // Light Cyan
        12 => 0x0000FF,  // Light Red
        13 => 0xFF00FF,  // Light Magenta
        14 => 0x00FFFF,  // Light Yellow
        15 => 0xFFFFFF,  // Bright White
        _ => 0x000000,
    };
    Ok(Value::Long(color))
}

/// Print/Debug.Print - Outputs to console (maps to MsgBox side effect)
pub fn print_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    let output: String = args.iter().map(|a| a.as_string()).collect::<Vec<_>>().join(" ");
    // This returns the string; the interpreter will handle it as a side effect
    Ok(Value::String(output))
}

/// Err.Number / Err.Description stubs
pub fn err_number_fn(_args: &[Value]) -> Result<Value, RuntimeError> {
    Ok(Value::Integer(0))
}

pub fn err_description_fn(_args: &[Value]) -> Result<Value, RuntimeError> {
    Ok(Value::String(String::new()))
}

/// String.IsNullOrEmpty(s) - Returns True if string is null or empty
pub fn isnullorempty_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("IsNullOrEmpty requires exactly one argument".to_string()));
    }
    let result = match &args[0] {
        Value::Nothing => true,
        Value::String(s) => s.is_empty(),
        _ => false,
    };
    Ok(Value::Boolean(result))
}

/// DateAdd(interval, number, date) - Stub: returns date unchanged
pub fn dateadd_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 3 {
        return Err(RuntimeError::Custom("DateAdd requires exactly 3 arguments".to_string()));
    }
    Ok(args[2].clone())
}

/// DateDiff(interval, date1, date2) - Stub: returns 0
pub fn datediff_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 3 {
        return Err(RuntimeError::Custom("DateDiff requires exactly 3 arguments".to_string()));
    }
    Ok(Value::Integer(0))
}

/// Strings.StrDup(count, char) - Duplicate a character
pub fn strdup_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 2 {
        return Err(RuntimeError::Custom("StrDup requires exactly 2 arguments".to_string()));
    }
    let count = args[0].as_integer()?.max(0) as usize;
    let s = args[1].as_string();
    let ch = s.chars().next().unwrap_or(' ');
    Ok(Value::String(std::iter::repeat(ch).take(count).collect()))
}
