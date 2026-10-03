//! Byte-oriented PHP string scans. A mask is a set of bytes, not Unicode
//! characters; keep the tight scan out of generated guest bytecode.
use std::sync::Arc;
use vybe_runtime::{Framework, Value};

fn addcslashes(args: &[Value]) -> Value {
    let [Value::String(subject), Value::String(characters), ..] = args else {
        return Value::Null;
    };

    // PHP charlists and strings are byte-oriented. Expand ranges once, rather
    // than searching the charlist and comparing dynamic values for every byte.
    let chars = characters.as_bytes();
    let mut selected = [false; 256];
    let mut index = 0;
    while index < chars.len() {
        if index + 3 < chars.len()
            && chars[index + 1] == b'.'
            && chars[index + 2] == b'.'
            && chars[index] <= chars[index + 3]
        {
            for byte in chars[index]..=chars[index + 3] {
                selected[byte as usize] = true;
            }
            index += 4;
        } else {
            selected[chars[index] as usize] = true;
            index += 1;
        }
    }

    let mut output = Vec::with_capacity(subject.len());
    for &byte in subject.as_bytes() {
        if !selected[byte as usize] {
            output.push(byte);
            continue;
        }
        output.push(b'\\');
        match byte {
            0 => output.push(b'0'),
            7 => output.push(b'a'),
            8 => output.push(b'b'),
            9 => output.push(b't'),
            10 => output.push(b'n'),
            11 => output.push(b'v'),
            12 => output.push(b'f'),
            13 => output.push(b'r'),
            1..=31 | 127..=255 => {
                output.push(b'0' + (byte >> 6));
                output.push(b'0' + ((byte >> 3) & 7));
                output.push(b'0' + (byte & 7));
            }
            _ => output.push(byte),
        }
    }

    // Value::String is UTF-8. When a charlist selects only part of a UTF-8
    // sequence, PHP can return bytes the current Value type cannot represent.
    let output = String::from_utf8(output)
        .unwrap_or_else(|err| String::from_utf8_lossy(&err.into_bytes()).into_owned());
    Value::String(Arc::from(output))
}

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn(
        "php:string",
        "addcslashes",
        Box::new(|_, args| addcslashes(args)),
    );
    fw.register_host_fn(
        "php:string",
        "pregQuote",
        Box::new(|_, args| {
            let Some(Value::String(subject)) = args.first() else {
                return Value::String(Arc::from(""));
            };
            let delimiter = match args.get(1) {
                Some(Value::String(value)) => value.as_bytes().first().copied(),
                _ => None,
            };
            let mut quoted = String::with_capacity(subject.len());
            for ch in subject.chars() {
                if ch == '\0' {
                    quoted.push_str("\\000");
                    continue;
                }
                let mut encoded = [0; 4];
                let bytes = ch.encode_utf8(&mut encoded);
                if ".\\+*?[^]$(){}=!<>|:-#".contains(ch)
                    || delimiter == bytes.as_bytes().first().copied()
                {
                    quoted.push('\\');
                }
                quoted.push(ch);
            }
            Value::String(Arc::from(quoted))
        }),
    );
    fw.register_host_fn(
        "php:string",
        "byteLength",
        Box::new(|_, args| match args.first() {
            Some(Value::String(value)) => Value::I64(value.len() as i64),
            _ => Value::I64(0),
        }),
    );
    for (name, reject) in [("strspn", false), ("strcspn", true)] {
        fw.register_host_fn(
            "php:string",
            name,
            Box::new(move |_, args| {
                let [Value::String(subject), Value::String(mask), ..] = args else {
                    return Value::I64(0);
                };
                let len = subject.len() as i64;
                let offset = args.get(2).map_or(0, |v| v.as_f64() as i64);
                let start = if offset < 0 {
                    len.saturating_add(offset).max(0)
                } else {
                    offset.min(len)
                };
                let end = match args.get(3) {
                    None | Some(Value::Null | Value::Undefined) => len,
                    Some(value) => {
                        let length = value.as_f64() as i64;
                        if length < 0 {
                            len.saturating_add(length).max(start)
                        } else {
                            start.saturating_add(length).min(len)
                        }
                    }
                };
                let mut accepted = [false; 256];
                for byte in mask.as_bytes() {
                    accepted[*byte as usize] = true;
                }
                let count = subject.as_bytes()[start as usize..end as usize]
                    .iter()
                    .take_while(|byte| accepted[**byte as usize] != reject)
                    .count();
                Value::I64(count as i64)
            }),
        );
    }
}
