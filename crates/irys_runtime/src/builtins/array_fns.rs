use crate::value::{RuntimeError, Value};

pub fn ubound_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("UBound requires at least one argument".to_string()));
    }

    match &args[0] {
        Value::Array(arr) => {
            if arr.is_empty() {
                Ok(Value::Integer(-1))
            } else {
                Ok(Value::Integer(arr.len() as i32 - 1))
            }
        }
        _ => Err(RuntimeError::TypeError {
            expected: "Array".to_string(),
            got: format!("{:?}", args[0]),
        }),
    }
}

pub fn lbound_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("LBound requires at least one argument".to_string()));
    }

    match &args[0] {
        Value::Array(_) => Ok(Value::Integer(0)), // Arrays are always 0-based
        _ => Err(RuntimeError::TypeError {
            expected: "Array".to_string(),
            got: format!("{:?}", args[0]),
        }),
    }
}
