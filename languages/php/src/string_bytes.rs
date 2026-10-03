//! PHP string operations whose offsets are UTF-8 bytes. The VM's ordinary
//! string substring primitive indexes UTF-16 units, which corrupts template
//! source after a multibyte character when paired with PCRE byte offsets.
use std::sync::Arc;
use vybe_runtime::{Framework, Value};

fn php_string(value: &Value) -> String {
    match value {
        Value::Null | Value::Bool(false) => String::new(),
        Value::Bool(true) => "1".to_string(),
        _ => value.to_string(),
    }
}

fn integer(value: &Value) -> i64 {
    match value {
        Value::I32(n) => *n as i64,
        Value::I64(n) => *n,
        Value::F64(n) => *n as i64,
        Value::Bool(v) => i64::from(*v),
        _ => value.to_string().parse().unwrap_or(0),
    }
}

fn substr(args: &[Value]) -> Value {
    let Some(source) = args.first() else {
        return Value::Null;
    };
    let Some(offset) = args.get(1) else {
        return Value::Null;
    };
    let source = php_string(source);
    let size = source.len() as i64;
    let start = integer(offset);
    let start = if start < 0 {
        size.saturating_add(start)
    } else {
        start
    }
    .clamp(0, size);
    let end = match args.get(2) {
        None | Some(Value::Null) => size,
        Some(length) => {
            let length = integer(length);
            if length < 0 {
                size.saturating_add(length)
            } else {
                start.saturating_add(length)
            }
        }
    }
    .clamp(start, size);
    let bytes = &source.as_bytes()[start as usize..end as usize];
    // Value::String is UTF-8. A slice through the middle of a code point
    // cannot be represented yet; valid boundaries (including Twig's PCRE
    // offsets) round-trip exactly.
    Value::String(Arc::from(String::from_utf8_lossy(bytes).as_ref()))
}

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn("php:string", "substr", Box::new(|_, args| substr(args)));
}
