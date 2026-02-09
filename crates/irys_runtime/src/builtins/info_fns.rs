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

/// IsDate(expression) - Returns True if value can be converted to a date (stub)
pub fn isdate_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("IsDate requires exactly one argument".to_string()));
    }
    // Simplified: just check some common date patterns
    let s = args[0].as_string();
    let is_date = s.contains('/') || s.contains('-');
    Ok(Value::Boolean(is_date))
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

/// InputBox(prompt[, title[, default]]) - Returns input string (stub: returns default or empty)
pub fn inputbox_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() || args.len() > 3 {
        return Err(RuntimeError::Custom("InputBox requires 1 to 3 arguments".to_string()));
    }
    // In a console/headless environment, just return the default value
    let default = if args.len() >= 3 {
        args[2].as_string()
    } else {
        String::new()
    };
    Ok(Value::String(default))
}
