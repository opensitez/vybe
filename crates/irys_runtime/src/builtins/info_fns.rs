use crate::value::{RuntimeError, Value};

/// IsNumeric(expression) - Returns True if value can be evaluated as a number
pub fn isnumeric_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("IsNumeric requires exactly one argument".to_string()));
    }
    let result = match &args[0] {
        Value::Integer(_) | Value::Long(_) | Value::Single(_) | Value::Double(_) | Value::Byte(_) | Value::Date(_) => true,
        Value::String(s) => s.trim().parse::<f64>().is_ok(),
        Value::Boolean(_) => true,
        _ => false,
    };
    Ok(Value::Boolean(result))
}

/// IsArray(expression) - Returns True if value is an array
pub fn isarray_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("IsArray requires exactly one argument".to_string()));
    }
    Ok(Value::Boolean(matches!(&args[0], Value::Array(_))))
}

/// IsNothing(expression) - Returns True if value is Nothing
pub fn isnothing_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("IsNothing requires exactly one argument".to_string()));
    }
    Ok(Value::Boolean(matches!(&args[0], Value::Nothing)))
}

/// IsDate(expression) - Returns True if value can be converted to a date
pub fn isdate_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("IsDate requires exactly one argument".to_string()));
    }
    
    // Check if value is already a date type
    if matches!(&args[0], Value::Date(_)) {
        return Ok(Value::Boolean(true));
    }
    
    // Try parsing string as date
    let s = args[0].as_string();
    let formats = [
        "%m/%d/%Y", "%Y-%m-%d", "%m/%d/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S",
        "%m/%d/%y", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d",
        "%B %d, %Y", "%b %d, %Y", "%d %B %Y", "%d %b %Y",
    ];
    
    use chrono::NaiveDate;
    for fmt in &formats {
        if NaiveDate::parse_from_str(&s, fmt).is_ok() {
            return Ok(Value::Boolean(true));
        }
    }
    
    Ok(Value::Boolean(false))
}

/// TypeName(expression) - Returns a string describing the type
pub fn typename_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("TypeName requires exactly one argument".to_string()));
    }
    let name = match &args[0] {
        Value::Integer(_) => "Integer",
        Value::Long(_) => "Long",
        Value::Single(_) => "Single",
        Value::Double(_) => "Double",
        Value::Date(_) => "Date",
        Value::String(_) => "String",
        Value::Boolean(_) => "Boolean",
        Value::Byte(_) => "Byte",
        Value::Char(_) => "Char",
        Value::Array(_) => "Variant()",
        Value::Nothing => "Nothing",
        Value::Object(obj_ref) => {
            let borrowed = obj_ref.borrow();
            return Ok(Value::String(borrowed.class_name.clone()));
        }
        Value::Collection(_) => "Collection",
        Value::Lambda { .. } => "Lambda",
    };
    Ok(Value::String(name.to_string()))
}

/// VarType(expression) - Returns an integer indicating the type
pub fn vartype_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("VarType requires exactly one argument".to_string()));
    }
    let vt = match &args[0] {
        Value::Nothing => 1,    // vbNull
        Value::Integer(_) => 2, // vbInteger
        Value::Long(_) => 3,    // vbLong
        Value::Single(_) => 4,  // vbSingle
        Value::Double(_) => 5,  // vbDouble
        Value::Date(_) => 7,    // vbDate
        Value::String(_) => 8,  // vbString
        Value::Boolean(_) => 11, // vbBoolean
        Value::Byte(_) => 17,   // vbByte
        Value::Char(_) => 18,   // vbChar
        Value::Object(_) => 9,  // vbObject
        Value::Collection(_) => 9, // vbObject
        Value::Array(_) => 8192, // vbArray
        Value::Lambda { .. } => 9, // vbObject
    };
    Ok(Value::Integer(vt))
}

/// IIf(expression, truepart, falsepart) - Inline If
pub fn iif_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 3 {
        return Err(RuntimeError::Custom("IIf requires exactly 3 arguments".to_string()));
    }
    let condition = args[0].as_bool()?;
    Ok(if condition { args[1].clone() } else { args[2].clone() })
}

/// Choose(index, choice1, choice2, ...) - Returns value at index position
pub fn choose_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() < 2 {
        return Err(RuntimeError::Custom("Choose requires at least 2 arguments".to_string()));
    }
    let index = args[0].as_integer()?;
    if index < 1 || index as usize >= args.len() {
        return Ok(Value::Nothing);
    }
    Ok(args[index as usize].clone())
}

/// Switch(expr1, val1, expr2, val2, ...) - Returns value for first True expression
pub fn switch_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() % 2 != 0 {
        return Err(RuntimeError::Custom("Switch requires an even number of arguments".to_string()));
    }
    for i in (0..args.len()).step_by(2) {
        if args[i].as_bool()? {
            return Ok(args[i + 1].clone());
        }
    }
    Ok(Value::Nothing)
}

/// Array(elem1, elem2, ...) - Creates an array from arguments
pub fn array_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    Ok(Value::Array(args.to_vec()))
}

/// InputBox(prompt[, title[, default]]) - Shows native input dialog and returns user input
pub fn inputbox_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    
    if args.is_empty() || args.len() > 3 {
        return Err(RuntimeError::Custom("InputBox requires 1 to 3 arguments".to_string()));
    }
    
    let prompt = args[0].as_string();
    let title = if args.len() >= 2 {
        args[1].as_string()
    } else {
        "Input".to_string()
    };
    let default = if args.len() >= 3 {
        args[2].as_string()
    } else {
        String::new()
    };
    
    match show_native_input_dialog(&prompt, &title, &default) {
        Some(input) => Ok(Value::String(input)),
        None => Ok(Value::String(String::new())), // User cancelled
    }
}

/// Show native OS input dialog
fn show_native_input_dialog(prompt: &str, title: &str, default_value: &str) -> Option<String> {
    #[cfg(target_os = "macos")]
    {
        use std::process::Command;
        
        let script = format!(
            "display dialog \"{}\" default answer \"{}\" with title \"{}\" buttons {{\"OK\", \"Cancel\"}} default button \"OK\"",
            prompt.replace("\"", "\\\""),
            default_value.replace("\"", "\\\""),
            title.replace("\"", "\\\"")
        );
        
        match Command::new("osascript").arg("-e").arg(&script).output() {
            Ok(output) if output.status.success() => {
                let result = String::from_utf8_lossy(&output.stdout);
                // Parse result like "button returned:OK, text returned:Hello"
                if let Some(text_part) = result.split("text returned:").nth(1) {
                    return Some(text_part.trim().to_string());
                }
                None
            }
            _ => None,
        }
    }
    
    #[cfg(target_os = "linux")]
    {
        use std::process::Command;
        
        match Command::new("zenity")
            .arg("--entry")
            .arg(format!("--title={}", title))
            .arg(format!("--text={}", prompt))
            .arg(format!("--entry-text={}", default_value))
            .output()
        {
            Ok(output) if output.status.success() => {
                Some(String::from_utf8_lossy(&output.stdout).trim().to_string())
            }
            _ => None,
        }
    }
    
    #[cfg(target_os = "windows")]
    {
        use std::process::Command;
        
        // Use PowerShell with InputBox from VB assembly
        let script = format!(
            "[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic'); \
             [Microsoft.VisualBasic.Interaction]::InputBox('{}', '{}', '{}')",
            prompt.replace("'", "''"),
            title.replace("'", "''"),
            default_value.replace("'", "''")
        );
        
        match Command::new("powershell")
            .arg("-NoProfile")
            .arg("-Command")
            .arg(&script)
            .output()
        {
            Ok(output) if output.status.success() => {
                Some(String::from_utf8_lossy(&output.stdout).trim().to_string())
            }
            _ => None,
        }
    }
    
    #[cfg(not(any(target_os = "macos", target_os = "linux", target_os = "windows")))]
    {
        None
    }
}

/// IsEmpty(expression) - Returns True if value is empty (uninitialized Variant or empty string)
pub fn isempty_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("IsEmpty requires exactly one argument".to_string()));
    }
    let result = match &args[0] {
        Value::Nothing => true,
        Value::String(s) => s.is_empty(),
        _ => false,
    };
    Ok(Value::Boolean(result))
}

/// IsObject(expression) - Returns True if value is an object reference
pub fn isobject_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("IsObject requires exactly one argument".to_string()));
    }
    let result = matches!(&args[0], Value::Object(_) | Value::Collection(_));
    Ok(Value::Boolean(result))
}

/// IsError(expression) - Returns True if value is an error value
pub fn iserror_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("IsError requires exactly one argument".to_string()));
    }
    // In VB, CVErr creates error values. We don't have a dedicated Error type,
    // but we can check for special patterns or Nothing which might represent errors
    let result = match &args[0] {
        Value::Nothing => false, // Nothing is not an error, it's an absence of value
        Value::String(s) => {
            // Check for common error string patterns
            let lower = s.to_lowercase();
            lower.starts_with("error") || 
            lower.contains("exception") || 
            lower.starts_with("err:")
        }
        _ => false,
    };
    Ok(Value::Boolean(result))
}

/// IsDBNull(expression) - Returns True if value is DBNull (used with database operations)
pub fn isdbnull_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("IsDBNull requires exactly one argument".to_string()));
    }
    // DBNull is typically represented as Nothing in our system
    Ok(Value::Boolean(matches!(&args[0], Value::Nothing)))
}
