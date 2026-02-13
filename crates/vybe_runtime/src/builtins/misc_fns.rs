use crate::value::{RuntimeError, Value};

/// DoEvents() - Yields to the event loop (stub: no-op in interpreter)
pub fn doevents_fn(_args: &[Value]) -> Result<Value, RuntimeError> {
    Ok(Value::Nothing)
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
