//! PHP HTML-special-character escaping, including existing-entity handling.
use vybe_runtime::{Framework, Value};

use super::html_entity_names::{HTML4_ENTITY_NAMES, HTML5_ENTITY_NAMES};

fn entity_length(text: &str, document: i32) -> Option<usize> {
    let end = text.bytes().take(64).position(|byte| byte == b';')?;
    let name = text.get(1..end)?;
    let valid = if let Some(number) = name.strip_prefix('#') {
        let number = if let Some(hex) = number.strip_prefix('x').or_else(|| number.strip_prefix('X')) {
            if hex.is_empty() || !hex.bytes().all(|byte| byte.is_ascii_hexdigit()) { return None; }
            u32::from_str_radix(hex, 16).ok()
        } else {
            if number.is_empty() || !number.bytes().all(|byte| byte.is_ascii_digit()) { return None; }
            number.parse::<u32>().ok()
        };
        number.is_some_and(|number| number <= 0x10ffff)
    } else {
        match document {
            16 => matches!(name, "amp" | "lt" | "gt" | "quot" | "apos"),
            48 => HTML5_ENTITY_NAMES.binary_search(&name).is_ok(),
            32 => name == "apos" || HTML4_ENTITY_NAMES.binary_search(&name).is_ok(),
            _ => HTML4_ENTITY_NAMES.binary_search(&name).is_ok(),
        }
    };
    valid.then_some(end + 1)
}

fn php_boolean(value: &Value) -> bool {
    match value {
        Value::Null | Value::Undefined => false,
        Value::Bool(value) => *value,
        Value::I32(value) => *value != 0,
        Value::I64(value) => *value != 0,
        Value::F64(value) => *value != 0.0,
        Value::String(value) => !value.is_empty() && value.as_ref() != "0",
        _ => true,
    }
}

fn disallowed(character: char, document: i32) -> bool {
    let code = character as u32;
    if document == 0 {
        matches!(code, 0..=8 | 11..=12 | 14..=31 | 127..=159)
    } else if document == 48 {
        matches!(code, 0..=8 | 11 | 14..=31 | 127..=159 | 0xfdd0..=0xfdef)
            || code & 0xffff >= 0xfffe
    } else {
        matches!(code, 0..=8 | 11..=12 | 14..=31 | 0xfffe..=0xffff)
    }
}

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn("php:html", "specialChars", Box::new(|_, args| {
        let Some(Value::String(input)) = args.first() else { return Value::String("".into()); };
        let flags = args.get(1).map_or(11, Value::as_i32);
        let document = flags & 48;
        let double_encode = args.get(3).is_none_or(php_boolean);
        let mut output = String::with_capacity(input.len());
        let mut position = 0;
        while position < input.len() {
            let rest = &input[position..];
            let character = rest.chars().next().unwrap();
            if character == '&' && !double_encode {
                if let Some(length) = entity_length(rest, document) {
                    output.push_str(&rest[..length]);
                    position += length;
                    continue;
                }
            }
            if flags & 128 != 0 && disallowed(character, document) {
                output.push('\u{fffd}');
            } else {
                match character {
                    '&' => output.push_str("&amp;"),
                    '<' => output.push_str("&lt;"),
                    '>' => output.push_str("&gt;"),
                    '"' if flags & 2 != 0 => output.push_str("&quot;"),
                    '\'' if flags & 1 != 0 => output.push_str(if document == 0 { "&#039;" } else { "&apos;" }),
                    _ => output.push(character),
                }
            }
            position += character.len_utf8();
        }
        Value::String(output.into())
    }));
}
