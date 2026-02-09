use crate::value::{RuntimeError, Value};

/// Abs(number) - Returns absolute value
pub fn abs_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Abs requires exactly one argument".to_string()));
    }
    match &args[0] {
        Value::Integer(i) => Ok(Value::Integer(i.abs())),
        Value::Long(l) => Ok(Value::Long(l.abs())),
        Value::Single(f) => Ok(Value::Single(f.abs())),
        Value::Double(d) => Ok(Value::Double(d.abs())),
        _ => {
            let d = args[0].as_double()?;
            Ok(Value::Double(d.abs()))
        }
    }
}

/// Int(number) - Returns the integer portion, rounding toward negative infinity
pub fn int_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Int requires exactly one argument".to_string()));
    }
    let d = args[0].as_double()?;
    Ok(Value::Integer(d.floor() as i32))
}

/// Fix(number) - Returns the integer portion, truncating toward zero
pub fn fix_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Fix requires exactly one argument".to_string()));
    }
    let d = args[0].as_double()?;
    Ok(Value::Integer(d.trunc() as i32))
}

/// Sgn(number) - Returns sign: -1, 0, or 1
pub fn sgn_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Sgn requires exactly one argument".to_string()));
    }
    let d = args[0].as_double()?;
    Ok(Value::Integer(if d < 0.0 { -1 } else if d > 0.0 { 1 } else { 0 }))
}

/// Sqr(number) - Returns square root
pub fn sqr_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Sqr requires exactly one argument".to_string()));
    }
    let d = args[0].as_double()?;
    if d < 0.0 {
        return Err(RuntimeError::Custom("Sqr: cannot take square root of negative number".to_string()));
    }
    Ok(Value::Double(d.sqrt()))
}

/// Rnd([number]) - Returns a random number between 0 and 1
pub fn rnd_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() > 1 {
        return Err(RuntimeError::Custom("Rnd requires 0 or 1 arguments".to_string()));
    }
    // Simple pseudo-random using system time
    let seed = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .subsec_nanos();
    let r = ((seed as f64 * 1.0) % 1000000.0) / 1000000.0;
    Ok(Value::Single(r as f32))
}

/// Round(number[, decimal_places]) - Rounds to specified decimal places
pub fn round_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() || args.len() > 2 {
        return Err(RuntimeError::Custom("Round requires 1 or 2 arguments".to_string()));
    }
    let d = args[0].as_double()?;
    let places = if args.len() == 2 {
        args[1].as_integer()?.max(0) as u32
    } else {
        0
    };

    let factor = 10f64.powi(places as i32);
    let rounded = (d * factor).round() / factor;

    if places == 0 {
        Ok(Value::Integer(rounded as i32))
    } else {
        Ok(Value::Double(rounded))
    }
}

/// Log(number) - Returns natural logarithm
pub fn log_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Log requires exactly one argument".to_string()));
    }
    let d = args[0].as_double()?;
    if d <= 0.0 {
        return Err(RuntimeError::Custom("Log: argument must be positive".to_string()));
    }
    Ok(Value::Double(d.ln()))
}

/// Exp(number) - Returns e raised to a power
pub fn exp_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Exp requires exactly one argument".to_string()));
    }
    let d = args[0].as_double()?;
    Ok(Value::Double(d.exp()))
}

/// Sin(number) - Returns sine (radians)
pub fn sin_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Sin requires exactly one argument".to_string()));
    }
    Ok(Value::Double(args[0].as_double()?.sin()))
}

/// Cos(number) - Returns cosine (radians)
pub fn cos_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Cos requires exactly one argument".to_string()));
    }
    Ok(Value::Double(args[0].as_double()?.cos()))
}

/// Tan(number) - Returns tangent (radians)
pub fn tan_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Tan requires exactly one argument".to_string()));
    }
    Ok(Value::Double(args[0].as_double()?.tan()))
}

/// Atn(number) - Returns arctangent (radians)
pub fn atn_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Atn requires exactly one argument".to_string()));
    }
    Ok(Value::Double(args[0].as_double()?.atan()))
}

/// Max(a, b) - Returns the larger of two values
pub fn max_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 2 {
        return Err(RuntimeError::Custom("Max requires exactly 2 arguments".to_string()));
    }
    let a = args[0].as_double()?;
    let b = args[1].as_double()?;
    if a >= b {
        Ok(args[0].clone())
    } else {
        Ok(args[1].clone())
    }
}

/// Min(a, b) - Returns the smaller of two values
pub fn min_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 2 {
        return Err(RuntimeError::Custom("Min requires exactly 2 arguments".to_string()));
    }
    let a = args[0].as_double()?;
    let b = args[1].as_double()?;
    if a <= b {
        Ok(args[0].clone())
    } else {
        Ok(args[1].clone())
    }
}

/// Ceiling(number) - Rounds up to nearest integer
pub fn ceiling_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Ceiling requires exactly one argument".to_string()));
    }
    let d = args[0].as_double()?;
    Ok(Value::Integer(d.ceil() as i32))
}

/// Floor(number) - Rounds down to nearest integer
pub fn floor_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 1 {
        return Err(RuntimeError::Custom("Floor requires exactly one argument".to_string()));
    }
    let d = args[0].as_double()?;
    Ok(Value::Integer(d.floor() as i32))
}

/// Pow(base, exponent) - Returns base raised to exponent power
pub fn pow_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 2 {
        return Err(RuntimeError::Custom("Pow requires exactly 2 arguments".to_string()));
    }
    let base = args[0].as_double()?;
    let exp = args[1].as_double()?;
    Ok(Value::Double(base.powf(exp)))
}

/// Randomize([seed]) - Initializes random number generator (stub)
pub fn randomize_fn(_args: &[Value]) -> Result<Value, RuntimeError> {
    // Stub: VB's Randomize seeds the RNG. We don't have persistent RNG state.
    Ok(Value::Nothing)
}

/// Atn2/Atan2(y, x) - Returns arctangent of y/x
pub fn atan2_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.len() != 2 {
        return Err(RuntimeError::Custom("Atan2 requires exactly 2 arguments".to_string()));
    }
    let y = args[0].as_double()?;
    let x = args[1].as_double()?;
    Ok(Value::Double(y.atan2(x)))
}
