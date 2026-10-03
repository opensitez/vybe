//! PHP's `$_COOKIE` decodes values written by `setcookie()`. The shared HTTP
//! cookie primitive keeps the raw header values for other language adapters.

use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Framework, Value};

fn hex_digit(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn decode_value(value: &str) -> String {
    let source = value.as_bytes();
    let mut bytes = Vec::with_capacity(source.len());
    let mut index = 0;
    while index < source.len() {
        if source[index] == b'%' && index + 2 < source.len() {
            if let (Some(hi), Some(lo)) =
                (hex_digit(source[index + 1]), hex_digit(source[index + 2]))
            {
                bytes.push((hi << 4) | lo);
                index += 3;
                continue;
            }
        }
        bytes.push(if source[index] == b'+' {
            b' '
        } else {
            source[index]
        });
        index += 1;
    }
    // Binary PHP strings use one Latin-1 code point per byte in the VM.
    bytes.into_iter().map(char::from).collect()
}

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn(
        "php:cookie",
        "decodeRequest",
        Box::new(|_, args| {
            let Some(Value::Object(source)) = args.first() else {
                return Value::Null;
            };
            let source = source.lock().unwrap();
            let ObjectKind::Map(entries) = &source.kind else {
                return Value::Null;
            };
            let mut decoded = Object::new();
            decoded.kind = ObjectKind::Map(Default::default());
            let ObjectKind::Map(output) = &mut decoded.kind else {
                unreachable!()
            };
            for (name, value) in entries {
                let value = match value {
                    Value::String(text) => Value::String(Arc::from(decode_value(text))),
                    other => other.clone(),
                };
                output.insert(name.clone(), value);
            }
            Value::Object(Arc::new(Mutex::new(decoded)))
        }),
    );
}

#[cfg(test)]
mod tests {
    use super::decode_value;

    #[test]
    fn php_cookie_values_decode_percent_and_plus() {
        assert_eq!(decode_value("a%2Bb%3D%2Fc+test"), "a+b=/c test");
        assert_eq!(decode_value("%FF%00%25"), "ÿ\0%");
        assert_eq!(decode_value("bad%Q1"), "bad%Q1");
    }
}
