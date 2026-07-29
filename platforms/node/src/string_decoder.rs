//! `node:string_decoder` — Node.js `StringDecoder`.
//!
//! Reference: <https://nodejs.org/api/string_decoder.html>.
//!
//! Stateful UTF-8 decoder that buffers incomplete multibyte sequences
//! across `write()` calls. State is stored as Object properties.

use std::collections::HashMap;
use std::sync::Arc;
use vybe_runtime::VM;
use vybe_runtime::value::{Object, ObjectKind, Value};

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn extract_bytes(buf: &Value) -> Vec<u8> {
    match buf {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            match &obj.kind {
                ObjectKind::Array(elems) => elems
                    .iter()
                    .map(|v| match v {
                        Value::I32(n) => *n as u8,
                        Value::F64(f) => *f as u8,
                        _ => 0,
                    })
                    .collect(),
                _ => vec![],
            }
        }
        _ => vec![],
    }
}

fn get_buf_bytes(decoder: &Value) -> Vec<u8> {
    if let Value::Object(obj) = decoder {
        let obj = obj.lock().unwrap();
        if let Some(Value::Object(buf)) = obj.properties.get("__buf") {
            let buf = buf.lock().unwrap();
            if let ObjectKind::Array(elems) = &buf.kind {
                return elems
                    .iter()
                    .map(|v| match v {
                        Value::I32(n) => *n as u8,
                        Value::F64(f) => *f as u8,
                        _ => 0,
                    })
                    .collect();
            }
        }
    }
    vec![]
}

fn set_buf_bytes(decoder: &Value, bytes: Vec<u8>) {
    if let Value::Object(obj) = decoder {
        let mut obj = obj.lock().unwrap();
        let elems: Vec<Value> = bytes.into_iter().map(|b| Value::I32(b as i32)).collect();
        let buf = Value::Object(vybe_runtime::heap::alloc(Object {
            kind: ObjectKind::Array(elems),
            properties: HashMap::new(),
            type_id: 0,
            fields: Vec::new(),
        }));
        obj.properties.insert("__buf".into(), buf);
    }
}

fn get_encoding(decoder: &Value) -> String {
    if let Value::Object(obj) = decoder {
        let obj = obj.lock().unwrap();
        if let Some(Value::String(enc)) = obj.properties.get("encoding") {
            return enc.to_string();
        }
    }
    "utf-8".to_string()
}

/// How many continuation bytes does a UTF-8 lead byte expect?
fn utf8_continuation_bytes_needed(lead: u8) -> usize {
    if lead & 0x80 == 0 {
        0
    }
    // 0xxxxxxx
    else if lead & 0xE0 == 0xC0 {
        1
    }
    // 110xxxxx
    else if lead & 0xF0 == 0xE0 {
        2
    }
    // 1110xxxx
    else if lead & 0xF8 == 0xF0 {
        3
    }
    // 11110xxx
    else {
        0
    }
}

fn decode_utf8_partial(bytes: &[u8]) -> (String, Vec<u8>) {
    let mut i = 0;
    let mut out = String::new();
    let len = bytes.len();
    while i < len {
        let lead = bytes[i];
        let needed = utf8_continuation_bytes_needed(lead);
        if i + needed < len {
            // Full sequence available
            let slice = &bytes[i..=i + needed];
            match std::str::from_utf8(slice) {
                Ok(s) => out.push_str(s),
                Err(_) => out.push('\u{FFFD}'),
            }
            i += needed + 1;
        } else {
            // Incomplete — buffer remaining bytes
            return (out, bytes[i..].to_vec());
        }
    }
    (out, vec![])
}

fn decode_bytes(encoding: &str, all_bytes: &[u8]) -> (String, Vec<u8>) {
    match encoding {
        "utf-8" | "utf8" => decode_utf8_partial(all_bytes),
        "latin1" | "binary" => {
            let s: String = all_bytes
                .iter()
                .map(|&b| char::from_u32(b as u32).unwrap_or('\u{FFFD}'))
                .collect();
            (s, vec![])
        }
        "hex" => {
            let s: String = all_bytes.iter().map(|b| format!("{b:02x}")).collect();
            (s, vec![])
        }
        "base64" => {
            const TABLE: &[u8] =
                b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
            let mut out = String::new();
            for chunk in all_bytes.chunks(3) {
                let b0 = chunk[0] as usize;
                let b1 = if chunk.len() > 1 {
                    chunk[1] as usize
                } else {
                    0
                };
                let b2 = if chunk.len() > 2 {
                    chunk[2] as usize
                } else {
                    0
                };
                out.push(TABLE[b0 >> 2] as char);
                out.push(TABLE[((b0 & 3) << 4) | (b1 >> 4)] as char);
                if chunk.len() > 1 {
                    out.push(TABLE[((b1 & 0xF) << 2) | (b2 >> 6)] as char);
                } else {
                    out.push('=');
                }
                if chunk.len() > 2 {
                    out.push(TABLE[b2 & 0x3F] as char);
                } else {
                    out.push('=');
                }
            }
            (out, vec![])
        }
        "ascii" => {
            let s: String = all_bytes.iter().map(|&b| (b & 0x7F) as char).collect();
            (s, vec![])
        }
        "utf16le" | "ucs2" => {
            let chars: Vec<u16> = all_bytes
                .chunks(2)
                .filter(|c| c.len() == 2)
                .map(|c| u16::from_le_bytes([c[0], c[1]]))
                .collect();
            (String::from_utf16_lossy(&chars), vec![])
        }
        _ => (String::from_utf8_lossy(all_bytes).to_string(), vec![]),
    }
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "node:string_decoder",
        "StringDecoder",
        Box::new(|_ctx, args| {
            let enc_raw = match args.first() {
                Some(Value::String(e)) => e.to_string(),
                _ => "utf8".to_string(),
            };
            let enc_norm = match enc_raw.to_lowercase().as_str() {
                "utf8" | "utf-8" => "utf-8",
                "latin1" | "binary" | "iso-8859-1" => "latin1",
                "hex" => "hex",
                "base64" | "base64url" => "base64",
                "ascii" => "ascii",
                "utf16le" | "utf-16le" | "ucs2" | "ucs-2" => "utf16le",
                _ => "utf-8",
            };
            let mut obj = Object::new();
            obj.properties.insert("encoding".into(), s(enc_norm));
            let empty: Vec<Value> = vec![];
            obj.properties.insert(
                "__buf".into(),
                Value::Object(vybe_runtime::heap::alloc(Object {
                    kind: ObjectKind::Array(empty),
                    properties: HashMap::new(),
                    type_id: 0,
                    fields: Vec::new(),
                })),
            );
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    vm.register_host_fn(
        "node:string_decoder",
        "write",
        Box::new(|_ctx, args| {
            let decoder = args.first().cloned().unwrap_or(Value::Undefined);
            let buf_val = args.get(1).cloned().unwrap_or(Value::Undefined);
            let encoding = get_encoding(&decoder);
            let mut bytes = get_buf_bytes(&decoder);
            bytes.extend(extract_bytes(&buf_val));
            let (out, remaining) = decode_bytes(&encoding, &bytes);
            set_buf_bytes(&decoder, remaining);
            s(&out)
        }),
    );

    vm.register_host_fn(
        "node:string_decoder",
        "end",
        Box::new(|_ctx, args| {
            let decoder = args.first().cloned().unwrap_or(Value::Undefined);
            let buf_val = args.get(1).cloned();
            let encoding = get_encoding(&decoder);
            let mut bytes = get_buf_bytes(&decoder);
            if let Some(buf) = buf_val {
                bytes.extend(extract_bytes(&buf));
            }
            // Flush: for UTF-8, any remaining incomplete bytes → replacement char
            let out = if encoding == "utf-8" || encoding == "utf8" {
                let (mut decoded, leftover) = decode_bytes(&encoding, &bytes);
                if !leftover.is_empty() {
                    decoded.push('\u{FFFD}');
                }
                decoded
            } else {
                let (decoded, _) = decode_bytes(&encoding, &bytes);
                decoded
            };
            set_buf_bytes(&decoder, vec![]);
            s(&out)
        }),
    );
}

#[allow(dead_code)]
fn _force_use(_: ObjectKind) {}
