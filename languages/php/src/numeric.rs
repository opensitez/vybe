//! PHP numeric strings consume the entire string, unlike parseFloat.
use vybe_runtime::value::ObjectKind;
use vybe_runtime::{Framework, Value};

fn php_truthy(value: &Value) -> bool {
    match value {
        Value::Null | Value::Undefined => false,
        Value::Bool(value) => *value,
        Value::I32(value) => *value != 0,
        Value::I64(value) => *value != 0,
        Value::F32(value) => *value != 0.0 && !value.is_nan(),
        Value::F64(value) => *value != 0.0 && !value.is_nan(),
        Value::String(value) => !value.is_empty() && value.as_ref() != "0",
        Value::Object(value) => match &value.lock().unwrap().kind {
            ObjectKind::Array(items) => !items.is_empty(),
            ObjectKind::Map(items) => !items.is_empty(),
            _ => true,
        },
        _ => true,
    }
}

pub(crate) fn is_numeric_string(value: &str) -> bool {
    let value = value.trim_matches(|c| matches!(c, ' ' | '\t' | '\n' | '\r' | '\x0b' | '\x0c'));
    let bytes = value.as_bytes();
    let mut i = usize::from(matches!(bytes.first(), Some(b'+' | b'-')));
    let start = i;
    while bytes.get(i).is_some_and(u8::is_ascii_digit) {
        i += 1;
    }
    let mut digits = i - start;
    if bytes.get(i) == Some(&b'.') {
        i += 1;
        let start = i;
        while bytes.get(i).is_some_and(u8::is_ascii_digit) {
            i += 1;
        }
        digits += i - start;
    }
    if digits == 0 {
        return false;
    }
    if matches!(bytes.get(i), Some(b'e' | b'E')) {
        i += 1;
        if matches!(bytes.get(i), Some(b'+' | b'-')) {
            i += 1;
        }
        let start = i;
        while bytes.get(i).is_some_and(u8::is_ascii_digit) {
            i += 1;
        }
        if i == start {
            return false;
        }
    }
    i == bytes.len()
}

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn(
        "php:value",
        "looseBoolEq",
        Box::new(|_, args| {
            Value::Bool(args.len() >= 2 && php_truthy(&args[0]) == php_truthy(&args[1]))
        }),
    );
    fw.register_host_fn(
        "php:value",
        "isNumeric",
        Box::new(|_, args| {
            Value::Bool(match args.first() {
                Some(Value::I32(_) | Value::I64(_) | Value::F64(_)) => true,
                Some(Value::String(value)) => is_numeric_string(value),
                _ => false,
            })
        }),
    );
}
