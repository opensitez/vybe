//! `node:buffer` — Node.js Buffer module.
//!
//! Reference: <https://nodejs.org/api/buffer.html>.
//!
//! Buffers are represented as Object with ObjectKind::Array of I32 byte values (0-255).

use std::sync::Arc;
use vybe_runtime::VM;
use vybe_runtime::value::{Object, ObjectKind, Value};

fn bytes_from_value(v: &Value) -> Vec<u8> {
    match v {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            match &obj.kind {
                ObjectKind::Array(elems) => elems
                    .iter()
                    .map(|e| match e {
                        Value::I32(n) => *n as u8,
                        Value::F64(f) => *f as u8,
                        _ => 0 })
                    .collect(),
                _ => Vec::new() }
        }
        _ => Vec::new() }
}

fn bytes_to_buf(bytes: Vec<u8>) -> Value {
    let elems: Vec<Value> = bytes.iter().map(|b| Value::I32(*b as i32)).collect();
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(elems)))
}

fn hex_decode(s: &str) -> Vec<u8> {
    let s = s.trim();
    let mut bytes = Vec::new();
    let mut i = 0;
    while i + 1 < s.len() {
        if let Ok(b) = u8::from_str_radix(&s[i..i + 2], 16) {
            bytes.push(b);
        }
        i += 2;
    }
    bytes
}

fn hex_encode(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{:02x}", b)).collect()
}

fn base64_encode(bytes: &[u8]) -> String {
    const CHARS: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::new();
    let mut i = 0;
    while i < bytes.len() {
        let b0 = bytes[i] as u32;
        let b1 = if i + 1 < bytes.len() {
            bytes[i + 1] as u32
        } else {
            0
        };
        let b2 = if i + 2 < bytes.len() {
            bytes[i + 2] as u32
        } else {
            0
        };
        let n = (b0 << 16) | (b1 << 8) | b2;
        out.push(CHARS[((n >> 18) & 63) as usize] as char);
        out.push(CHARS[((n >> 12) & 63) as usize] as char);
        if i + 1 < bytes.len() {
            out.push(CHARS[((n >> 6) & 63) as usize] as char);
        } else {
            out.push('=');
        }
        if i + 2 < bytes.len() {
            out.push(CHARS[(n & 63) as usize] as char);
        } else {
            out.push('=');
        }
        i += 3;
    }
    out
}

fn base64_decode(s: &str) -> Vec<u8> {
    let decode_char = |c: char| -> u8 {
        match c {
            'A'..='Z' => c as u8 - b'A',
            'a'..='z' => c as u8 - b'a' + 26,
            '0'..='9' => c as u8 - b'0' + 52,
            '+' | '-' => 62,
            '/' | '_' => 63,
            _ => 255 }
    };
    let chars: Vec<u8> = s
        .chars()
        .filter(|c| *c != '=' && *c != '\n' && *c != '\r')
        .map(decode_char)
        .filter(|&b| b != 255)
        .collect();
    let mut out = Vec::new();
    let mut i = 0;
    while i + 1 < chars.len() {
        let b0 = chars[i] as u32;
        let b1 = chars[i + 1] as u32;
        out.push(((b0 << 2) | (b1 >> 4)) as u8);
        if i + 2 < chars.len() {
            let b2 = chars[i + 2] as u32;
            out.push(((b1 << 4) | (b2 >> 2)) as u8);
            if i + 3 < chars.len() {
                let b3 = chars[i + 3] as u32;
                out.push(((b2 << 6) | b3) as u8);
            }
        }
        i += 4;
    }
    out
}

fn base64_decode_strict(s: &str) -> Option<Vec<u8>> {
    let filtered: Vec<u8> = s.bytes().filter(|b| !b.is_ascii_whitespace()).collect();
    if filtered.len() % 4 == 1 {
        return None;
    }
    if filtered.is_empty() {
        return Some(Vec::new());
    }
    if filtered.len() % 4 != 0 {
        return None;
    }

    let decode = |b: u8| -> Option<u32> {
        match b {
            b'A'..=b'Z' => Some((b - b'A') as u32),
            b'a'..=b'z' => Some((b - b'a' + 26) as u32),
            b'0'..=b'9' => Some((b - b'0' + 52) as u32),
            b'+' => Some(62),
            b'/' => Some(63),
            _ => None }
    };

    let mut out = Vec::new();
    let chunk_count = filtered.len() / 4;
    for (chunk_index, chunk) in filtered.chunks(4).enumerate() {
        let last = chunk_index + 1 == chunk_count;
        let a = decode(chunk[0])?;
        let b = decode(chunk[1])?;
        let c = chunk[2];
        let d = chunk[3];
        if c == b'=' {
            if d != b'=' || !last {
                return None;
            }
            out.push(((a << 2) | (b >> 4)) as u8);
            continue;
        }
        let c = decode(c)?;
        out.push(((a << 2) | (b >> 4)) as u8);
        out.push((((b & 0x0f) << 4) | (c >> 2)) as u8);
        if d == b'=' {
            if !last {
                return None;
            }
            continue;
        }
        let d = decode(d)?;
        out.push((((c & 0x03) << 6) | d) as u8);
    }
    Some(out)
}

fn decode_with_encoding(s: &str, enc: &str) -> Vec<u8> {
    match enc.to_lowercase().as_str() {
        "hex" => hex_decode(s),
        "base64" | "base64url" => base64_decode(s),
        "latin1" | "binary" | "ascii" => s.chars().map(|c| c as u8).collect(),
        "utf16le" | "ucs2" => {
            let chars: Vec<u8> = s.encode_utf16().flat_map(|c| c.to_le_bytes()).collect();
            chars
        }
        _ => s.as_bytes().to_vec() }
}

fn encode_with_encoding(bytes: &[u8], enc: &str) -> String {
    match enc.to_lowercase().as_str() {
        "hex" => hex_encode(bytes),
        "base64" => base64_encode(bytes),
        "base64url" => base64_encode(bytes)
            .replace('+', "-")
            .replace('/', "_")
            .replace('=', ""),
        "latin1" | "binary" | "ascii" => bytes.iter().map(|&b| b as char).collect(),
        _ => String::from_utf8_lossy(bytes).into_owned() }
}

fn get_encoding(args: &[Value], idx: usize) -> String {
    match args.get(idx) {
        Some(Value::String(s)) => s.to_string(),
        _ => "utf8".to_string() }
}

fn get_i32(args: &[Value], idx: usize, default: i32) -> i32 {
    match args.get(idx) {
        Some(Value::I32(n)) => *n,
        Some(Value::F64(f)) => *f as i32,
        _ => default }
}

pub fn register(vm: &mut VM) {
    // alloc(size[, fill[, encoding]])
    vm.register_host_fn(
        "node:buffer",
        "alloc",
        Box::new(|_ctx, args| {
            let size = get_i32(args, 0, 0).max(0) as usize;
            let fill = match args.get(1) {
                Some(Value::I32(n)) => (*n & 0xff) as u8,
                Some(Value::F64(f)) => (*f as i32 & 0xff) as u8,
                _ => 0 };
            bytes_to_buf(vec![fill; size])
        }),
    );

    // from(string, encoding) | from(array)
    vm.register_host_fn(
        "node:buffer",
        "from",
        Box::new(|_ctx, args| match args.first() {
            Some(Value::String(s)) => {
                let enc = get_encoding(args, 1);
                bytes_to_buf(decode_with_encoding(s, &enc))
            }
            Some(Value::Object(obj)) => {
                let obj = obj.lock().unwrap();
                match &obj.kind {
                    ObjectKind::Array(elems) => {
                        let bytes: Vec<u8> = elems
                            .iter()
                            .map(|e| match e {
                                Value::I32(n) => (*n & 0xff) as u8,
                                Value::F64(f) => (*f as i32 & 0xff) as u8,
                                _ => 0 })
                            .collect();
                        drop(obj);
                        bytes_to_buf(bytes)
                    }
                    _ => bytes_to_buf(Vec::new()) }
            }
            _ => bytes_to_buf(Vec::new()) }),
    );

    // fromBase64Strict(string) -> Buffer | null. Node's `Buffer.from(s,
    // "base64")` is intentionally lenient; .NET Convert.FromBase64String
    // needs strict validation so it can throw FormatException.
    vm.register_host_fn(
        "node:buffer",
        "fromBase64Strict",
        Box::new(|_ctx, args| {
            let Some(Value::String(text)) = args.first() else {
                return Value::Null;
            };
            match base64_decode_strict(text) {
                Some(bytes) => bytes_to_buf(bytes),
                None => Value::Null }
        }),
    );

    // byteLength(string, encoding)
    vm.register_host_fn(
        "node:buffer",
        "byteLength",
        Box::new(|_ctx, args| {
            let s = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                _ => return Value::I32(0) };
            let enc = get_encoding(args, 1);
            let len = decode_with_encoding(&s, &enc).len();
            Value::I32(len as i32)
        }),
    );

    // concat(list[, totalLength])
    vm.register_host_fn(
        "node:buffer",
        "concat",
        Box::new(|_ctx, args| {
            let mut all: Vec<u8> = Vec::new();
            if let Some(Value::Object(arr)) = args.first() {
                let arr = arr.lock().unwrap();
                if let ObjectKind::Array(elems) = &arr.kind {
                    for e in elems {
                        let bytes = bytes_from_value(e);
                        all.extend_from_slice(&bytes);
                    }
                }
            }
            if let Some(limit) = args.get(1) {
                let len = match limit {
                    Value::I32(n) => *n as usize,
                    Value::F64(f) => *f as usize,
                    _ => all.len() };
                all.truncate(len);
            }
            bytes_to_buf(all)
        }),
    );

    // isBuffer(value)
    vm.register_host_fn(
        "node:buffer",
        "isBuffer",
        Box::new(|_ctx, args| {
            let result = match args.first() {
                Some(Value::Object(obj)) => {
                    let obj = obj.lock().unwrap();
                    matches!(obj.kind, ObjectKind::Array(_))
                }
                _ => false };
            Value::Bool(result)
        }),
    );

    // isEncoding(encoding)
    vm.register_host_fn(
        "node:buffer",
        "isEncoding",
        Box::new(|_ctx, args| {
            let enc = match args.first() {
                Some(Value::String(s)) => s.to_lowercase(),
                _ => return Value::Bool(false) };
            let valid = [
                "utf8",
                "utf-8",
                "hex",
                "base64",
                "base64url",
                "ascii",
                "latin1",
                "binary",
                "ucs2",
                "ucs-2",
                "utf16le",
                "utf-16le",
            ];
            Value::Bool(valid.contains(&enc.as_str()))
        }),
    );

    // compare(buf1, buf2) → -1 | 0 | 1
    vm.register_host_fn(
        "node:buffer",
        "compare",
        Box::new(|_ctx, args| {
            let a = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let b = bytes_from_value(args.get(1).unwrap_or(&Value::Undefined));
            Value::I32(match a.cmp(&b) {
                std::cmp::Ordering::Less => -1,
                std::cmp::Ordering::Equal => 0,
                std::cmp::Ordering::Greater => 1 })
        }),
    );

    // toString(buf, encoding[, start, end])
    vm.register_host_fn(
        "node:buffer",
        "toString",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let enc = get_encoding(args, 1);
            let start = get_i32(args, 2, 0).max(0) as usize;
            let end = match args.get(3) {
                Some(Value::I32(n)) => (*n as usize).min(bytes.len()),
                Some(Value::F64(f)) => (*f as usize).min(bytes.len()),
                _ => bytes.len() };
            let slice = if start < bytes.len() {
                &bytes[start..end]
            } else {
                &[]
            };
            Value::String(Arc::from(encode_with_encoding(slice, &enc).as_str()))
        }),
    );

    // slice(buf, start, end)
    vm.register_host_fn(
        "node:buffer",
        "slice",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let len = bytes.len() as i32;
            let start = {
                let s = get_i32(args, 1, 0);
                if s < 0 {
                    (len + s).max(0) as usize
                } else {
                    s.min(len) as usize
                }
            };
            let end = {
                let e = get_i32(args, 2, len);
                if e < 0 {
                    (len + e).max(0) as usize
                } else {
                    e.min(len) as usize
                }
            };
            bytes_to_buf(bytes[start.min(end)..end].to_vec())
        }),
    );

    // subarray — alias of slice
    vm.register_host_fn(
        "node:buffer",
        "subarray",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let len = bytes.len() as i32;
            let start = {
                let s = get_i32(args, 1, 0);
                if s < 0 {
                    (len + s).max(0) as usize
                } else {
                    s.min(len) as usize
                }
            };
            let end = {
                let e = get_i32(args, 2, len);
                if e < 0 {
                    (len + e).max(0) as usize
                } else {
                    e.min(len) as usize
                }
            };
            bytes_to_buf(bytes[start.min(end)..end].to_vec())
        }),
    );

    // copy(src, dst, targetStart, srcStart, srcEnd) → bytes_copied
    vm.register_host_fn(
        "node:buffer",
        "copy",
        Box::new(|_ctx, args| {
            let src = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let src_len = src.len();
            let target_start = get_i32(args, 2, 0).max(0) as usize;
            let src_start = get_i32(args, 3, 0).max(0) as usize;
            let src_end = get_i32(args, 4, src_len as i32).max(0) as usize;
            let src_slice = if src_start < src_len {
                &src[src_start..src_end.min(src_len)]
            } else {
                &[]
            };
            if let Some(Value::Object(dst_obj)) = args.get(1) {
                let mut dst_obj = dst_obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = dst_obj.kind {
                    let mut copied = 0;
                    for (i, &b) in src_slice.iter().enumerate() {
                        let ti = target_start + i;
                        if ti < elems.len() {
                            elems[ti] = Value::I32(b as i32);
                            copied += 1;
                        }
                    }
                    return Value::I32(copied);
                }
            }
            Value::I32(0)
        }),
    );

    // indexOf(buf, val[, offset]) → i32
    vm.register_host_fn(
        "node:buffer",
        "indexOf",
        Box::new(|_ctx, args| {
            let haystack = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let needle: Vec<u8> = match args.get(1) {
                Some(Value::I32(n)) => vec![(*n & 0xff) as u8],
                Some(Value::F64(f)) => vec![(*f as i32 & 0xff) as u8],
                Some(Value::String(s)) => s.as_bytes().to_vec(),
                Some(v) => bytes_from_value(v),
                None => return Value::I32(-1) };
            let offset = get_i32(args, 2, 0).max(0) as usize;
            if needle.is_empty() {
                return Value::I32(offset as i32);
            }
            for i in offset..haystack.len() {
                if haystack[i..].starts_with(&needle) {
                    return Value::I32(i as i32);
                }
            }
            Value::I32(-1)
        }),
    );

    // includes(buf, val) → bool
    vm.register_host_fn(
        "node:buffer",
        "includes",
        Box::new(|_ctx, args| {
            let haystack = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let needle: Vec<u8> = match args.get(1) {
                Some(Value::I32(n)) => vec![(*n & 0xff) as u8],
                Some(Value::F64(f)) => vec![(*f as i32 & 0xff) as u8],
                Some(Value::String(s)) => s.as_bytes().to_vec(),
                Some(v) => bytes_from_value(v),
                None => return Value::Bool(false) };
            if needle.is_empty() {
                return Value::Bool(true);
            }
            for i in 0..haystack.len() {
                if haystack[i..].starts_with(&needle) {
                    return Value::Bool(true);
                }
            }
            Value::Bool(false)
        }),
    );

    // equals(buf1, buf2) → bool
    vm.register_host_fn(
        "node:buffer",
        "equals",
        Box::new(|_ctx, args| {
            let a = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let b = bytes_from_value(args.get(1).unwrap_or(&Value::Undefined));
            Value::Bool(a == b)
        }),
    );

    // fill(buf, value[, start, end]) → buf (modifies in place and returns)
    vm.register_host_fn(
        "node:buffer",
        "fill",
        Box::new(|_ctx, args| {
            let fill_byte = match args.get(1) {
                Some(Value::I32(n)) => (*n & 0xff) as u8,
                Some(Value::F64(f)) => (*f as i32 & 0xff) as u8,
                _ => 0 };
            if let Some(Value::Object(obj)) = args.first() {
                let mut obj = obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = obj.kind {
                    let len = elems.len();
                    let start = get_i32(args, 2, 0).max(0) as usize;
                    let end = get_i32(args, 3, len as i32).max(0) as usize;
                    for i in start..end.min(len) {
                        elems[i] = Value::I32(fill_byte as i32);
                    }
                }
            }
            args.first().cloned().unwrap_or(Value::Undefined)
        }),
    );

    // swap16(buf) — swap byte order in pairs
    vm.register_host_fn(
        "node:buffer",
        "swap16",
        Box::new(|_ctx, args| {
            let mut bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let mut i = 0;
            while i + 1 < bytes.len() {
                bytes.swap(i, i + 1);
                i += 2;
            }
            bytes_to_buf(bytes)
        }),
    );

    // swap32(buf) — swap byte order in quads
    vm.register_host_fn(
        "node:buffer",
        "swap32",
        Box::new(|_ctx, args| {
            let mut bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let mut i = 0;
            while i + 3 < bytes.len() {
                bytes.swap(i, i + 3);
                bytes.swap(i + 1, i + 2);
                i += 4;
            }
            bytes_to_buf(bytes)
        }),
    );

    // readUInt8(buf, offset)
    vm.register_host_fn(
        "node:buffer",
        "readUInt8",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            Value::I32(bytes.get(off).copied().unwrap_or(0) as i32)
        }),
    );

    // readInt8(buf, offset)
    vm.register_host_fn(
        "node:buffer",
        "readInt8",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            Value::I32(bytes.get(off).copied().unwrap_or(0) as i8 as i32)
        }),
    );

    // readUInt16BE
    vm.register_host_fn(
        "node:buffer",
        "readUInt16BE",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            if off + 1 < bytes.len() {
                Value::I32(u16::from_be_bytes([bytes[off], bytes[off + 1]]) as i32)
            } else {
                Value::I32(0)
            }
        }),
    );

    // readUInt16LE
    vm.register_host_fn(
        "node:buffer",
        "readUInt16LE",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            if off + 1 < bytes.len() {
                Value::I32(u16::from_le_bytes([bytes[off], bytes[off + 1]]) as i32)
            } else {
                Value::I32(0)
            }
        }),
    );

    // readUInt32BE
    vm.register_host_fn(
        "node:buffer",
        "readUInt32BE",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            if off + 3 < bytes.len() {
                Value::F64(u32::from_be_bytes([
                    bytes[off],
                    bytes[off + 1],
                    bytes[off + 2],
                    bytes[off + 3],
                ]) as f64)
            } else {
                Value::I32(0)
            }
        }),
    );

    // readUInt32LE
    vm.register_host_fn(
        "node:buffer",
        "readUInt32LE",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            if off + 3 < bytes.len() {
                Value::F64(u32::from_le_bytes([
                    bytes[off],
                    bytes[off + 1],
                    bytes[off + 2],
                    bytes[off + 3],
                ]) as f64)
            } else {
                Value::I32(0)
            }
        }),
    );

    // writeUInt8(buf, value, offset) → offset + 1
    vm.register_host_fn(
        "node:buffer",
        "writeUInt8",
        Box::new(|_ctx, args| {
            let val = (get_i32(args, 1, 0) & 0xff) as u8;
            let off = get_i32(args, 2, 0) as usize;
            if let Some(Value::Object(obj)) = args.first() {
                let mut obj = obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = obj.kind {
                    if off < elems.len() {
                        elems[off] = Value::I32(val as i32);
                    }
                }
            }
            Value::I32((off + 1) as i32)
        }),
    );

    // writeInt8(buf, value, offset) → offset + 1
    vm.register_host_fn(
        "node:buffer",
        "writeInt8",
        Box::new(|_ctx, args| {
            let val = (get_i32(args, 1, 0) as i8) as u8;
            let off = get_i32(args, 2, 0) as usize;
            if let Some(Value::Object(obj)) = args.first() {
                let mut obj = obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = obj.kind {
                    if off < elems.len() {
                        elems[off] = Value::I32(val as i32);
                    }
                }
            }
            Value::I32((off + 1) as i32)
        }),
    );

    // writeUInt16BE(buf, value, offset) → offset + 2
    vm.register_host_fn(
        "node:buffer",
        "writeUInt16BE",
        Box::new(|_ctx, args| {
            let val = (get_i32(args, 1, 0) as u16).to_be_bytes();
            let off = get_i32(args, 2, 0) as usize;
            if let Some(Value::Object(obj)) = args.first() {
                let mut obj = obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = obj.kind {
                    for (i, &b) in val.iter().enumerate() {
                        if off + i < elems.len() {
                            elems[off + i] = Value::I32(b as i32);
                        }
                    }
                }
            }
            Value::I32((off + 2) as i32)
        }),
    );

    // writeUInt16LE(buf, value, offset) → offset + 2
    vm.register_host_fn(
        "node:buffer",
        "writeUInt16LE",
        Box::new(|_ctx, args| {
            let val = (get_i32(args, 1, 0) as u16).to_le_bytes();
            let off = get_i32(args, 2, 0) as usize;
            if let Some(Value::Object(obj)) = args.first() {
                let mut obj = obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = obj.kind {
                    for (i, &b) in val.iter().enumerate() {
                        if off + i < elems.len() {
                            elems[off + i] = Value::I32(b as i32);
                        }
                    }
                }
            }
            Value::I32((off + 2) as i32)
        }),
    );

    // writeUInt32BE(buf, value, offset) → offset + 4
    vm.register_host_fn(
        "node:buffer",
        "writeUInt32BE",
        Box::new(|_ctx, args| {
            let val = match args.get(1) {
                Some(Value::I32(n)) => (*n as u32).to_be_bytes(),
                Some(Value::F64(f)) => (*f as u32).to_be_bytes(),
                _ => 0u32.to_be_bytes() };
            let off = get_i32(args, 2, 0) as usize;
            if let Some(Value::Object(obj)) = args.first() {
                let mut obj = obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = obj.kind {
                    for (i, &b) in val.iter().enumerate() {
                        if off + i < elems.len() {
                            elems[off + i] = Value::I32(b as i32);
                        }
                    }
                }
            }
            Value::I32((off + 4) as i32)
        }),
    );

    // writeUInt32LE(buf, value, offset) → offset + 4
    vm.register_host_fn(
        "node:buffer",
        "writeUInt32LE",
        Box::new(|_ctx, args| {
            let val = match args.get(1) {
                Some(Value::I32(n)) => (*n as u32).to_le_bytes(),
                Some(Value::F64(f)) => (*f as u32).to_le_bytes(),
                _ => 0u32.to_le_bytes() };
            let off = get_i32(args, 2, 0) as usize;
            if let Some(Value::Object(obj)) = args.first() {
                let mut obj = obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = obj.kind {
                    for (i, &b) in val.iter().enumerate() {
                        if off + i < elems.len() {
                            elems[off + i] = Value::I32(b as i32);
                        }
                    }
                }
            }
            Value::I32((off + 4) as i32)
        }),
    );

    // readInt16BE / readInt16LE
    vm.register_host_fn(
        "node:buffer",
        "readInt16BE",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            if off + 1 < bytes.len() {
                Value::I32(i16::from_be_bytes([bytes[off], bytes[off + 1]]) as i32)
            } else {
                Value::I32(0)
            }
        }),
    );
    vm.register_host_fn(
        "node:buffer",
        "readInt16LE",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            if off + 1 < bytes.len() {
                Value::I32(i16::from_le_bytes([bytes[off], bytes[off + 1]]) as i32)
            } else {
                Value::I32(0)
            }
        }),
    );

    // readInt32BE / readInt32LE
    vm.register_host_fn(
        "node:buffer",
        "readInt32BE",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            if off + 3 < bytes.len() {
                Value::I32(i32::from_be_bytes([
                    bytes[off],
                    bytes[off + 1],
                    bytes[off + 2],
                    bytes[off + 3],
                ]))
            } else {
                Value::I32(0)
            }
        }),
    );
    vm.register_host_fn(
        "node:buffer",
        "readInt32LE",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            if off + 3 < bytes.len() {
                Value::I32(i32::from_le_bytes([
                    bytes[off],
                    bytes[off + 1],
                    bytes[off + 2],
                    bytes[off + 3],
                ]))
            } else {
                Value::I32(0)
            }
        }),
    );

    // readFloatBE / readFloatLE
    vm.register_host_fn(
        "node:buffer",
        "readFloatBE",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            if off + 3 < bytes.len() {
                Value::F64(f32::from_be_bytes([
                    bytes[off],
                    bytes[off + 1],
                    bytes[off + 2],
                    bytes[off + 3],
                ]) as f64)
            } else {
                Value::F64(0.0)
            }
        }),
    );
    vm.register_host_fn(
        "node:buffer",
        "readFloatLE",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            if off + 3 < bytes.len() {
                Value::F64(f32::from_le_bytes([
                    bytes[off],
                    bytes[off + 1],
                    bytes[off + 2],
                    bytes[off + 3],
                ]) as f64)
            } else {
                Value::F64(0.0)
            }
        }),
    );

    // readDoubleBE / readDoubleLE
    vm.register_host_fn(
        "node:buffer",
        "readDoubleBE",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            if off + 7 < bytes.len() {
                let arr: [u8; 8] = bytes[off..off + 8].try_into().unwrap_or([0; 8]);
                Value::F64(f64::from_be_bytes(arr))
            } else {
                Value::F64(0.0)
            }
        }),
    );
    vm.register_host_fn(
        "node:buffer",
        "readDoubleLE",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            if off + 7 < bytes.len() {
                let arr: [u8; 8] = bytes[off..off + 8].try_into().unwrap_or([0; 8]);
                Value::F64(f64::from_le_bytes(arr))
            } else {
                Value::F64(0.0)
            }
        }),
    );

    // readBigUInt64BE / readBigUInt64LE
    vm.register_host_fn(
        "node:buffer",
        "readBigUInt64BE",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            if off + 7 < bytes.len() {
                let arr: [u8; 8] = bytes[off..off + 8].try_into().unwrap_or([0; 8]);
                Value::F64(u64::from_be_bytes(arr) as f64)
            } else {
                Value::F64(0.0)
            }
        }),
    );
    vm.register_host_fn(
        "node:buffer",
        "readBigUInt64LE",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            if off + 7 < bytes.len() {
                let arr: [u8; 8] = bytes[off..off + 8].try_into().unwrap_or([0; 8]);
                Value::F64(u64::from_le_bytes(arr) as f64)
            } else {
                Value::F64(0.0)
            }
        }),
    );

    // readBigInt64BE / readBigInt64LE
    vm.register_host_fn(
        "node:buffer",
        "readBigInt64BE",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            if off + 7 < bytes.len() {
                let arr: [u8; 8] = bytes[off..off + 8].try_into().unwrap_or([0; 8]);
                Value::F64(i64::from_be_bytes(arr) as f64)
            } else {
                Value::F64(0.0)
            }
        }),
    );
    vm.register_host_fn(
        "node:buffer",
        "readBigInt64LE",
        Box::new(|_ctx, args| {
            let bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let off = get_i32(args, 1, 0) as usize;
            if off + 7 < bytes.len() {
                let arr: [u8; 8] = bytes[off..off + 8].try_into().unwrap_or([0; 8]);
                Value::F64(i64::from_le_bytes(arr) as f64)
            } else {
                Value::F64(0.0)
            }
        }),
    );

    // writeInt16BE / writeInt16LE
    vm.register_host_fn(
        "node:buffer",
        "writeInt16BE",
        Box::new(|_ctx, args| {
            let val = (get_i32(args, 1, 0) as i16).to_be_bytes();
            let off = get_i32(args, 2, 0) as usize;
            if let Some(Value::Object(obj)) = args.first() {
                let mut obj = obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = obj.kind {
                    for (i, &b) in val.iter().enumerate() {
                        if off + i < elems.len() {
                            elems[off + i] = Value::I32(b as i32);
                        }
                    }
                }
            }
            Value::I32((off + 2) as i32)
        }),
    );
    vm.register_host_fn(
        "node:buffer",
        "writeInt16LE",
        Box::new(|_ctx, args| {
            let val = (get_i32(args, 1, 0) as i16).to_le_bytes();
            let off = get_i32(args, 2, 0) as usize;
            if let Some(Value::Object(obj)) = args.first() {
                let mut obj = obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = obj.kind {
                    for (i, &b) in val.iter().enumerate() {
                        if off + i < elems.len() {
                            elems[off + i] = Value::I32(b as i32);
                        }
                    }
                }
            }
            Value::I32((off + 2) as i32)
        }),
    );

    // writeInt32BE / writeInt32LE
    vm.register_host_fn(
        "node:buffer",
        "writeInt32BE",
        Box::new(|_ctx, args| {
            let val = (get_i32(args, 1, 0)).to_be_bytes();
            let off = get_i32(args, 2, 0) as usize;
            if let Some(Value::Object(obj)) = args.first() {
                let mut obj = obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = obj.kind {
                    for (i, &b) in val.iter().enumerate() {
                        if off + i < elems.len() {
                            elems[off + i] = Value::I32(b as i32);
                        }
                    }
                }
            }
            Value::I32((off + 4) as i32)
        }),
    );
    vm.register_host_fn(
        "node:buffer",
        "writeInt32LE",
        Box::new(|_ctx, args| {
            let val = (get_i32(args, 1, 0)).to_le_bytes();
            let off = get_i32(args, 2, 0) as usize;
            if let Some(Value::Object(obj)) = args.first() {
                let mut obj = obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = obj.kind {
                    for (i, &b) in val.iter().enumerate() {
                        if off + i < elems.len() {
                            elems[off + i] = Value::I32(b as i32);
                        }
                    }
                }
            }
            Value::I32((off + 4) as i32)
        }),
    );

    // writeFloatBE / writeFloatLE
    vm.register_host_fn(
        "node:buffer",
        "writeFloatBE",
        Box::new(|_ctx, args| {
            let f = match args.get(1) {
                Some(Value::F64(f)) => *f as f32,
                Some(Value::I32(n)) => *n as f32,
                _ => 0.0f32 };
            let val = f.to_be_bytes();
            let off = get_i32(args, 2, 0) as usize;
            if let Some(Value::Object(obj)) = args.first() {
                let mut obj = obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = obj.kind {
                    for (i, &b) in val.iter().enumerate() {
                        if off + i < elems.len() {
                            elems[off + i] = Value::I32(b as i32);
                        }
                    }
                }
            }
            Value::I32((off + 4) as i32)
        }),
    );
    vm.register_host_fn(
        "node:buffer",
        "writeFloatLE",
        Box::new(|_ctx, args| {
            let f = match args.get(1) {
                Some(Value::F64(f)) => *f as f32,
                Some(Value::I32(n)) => *n as f32,
                _ => 0.0f32 };
            let val = f.to_le_bytes();
            let off = get_i32(args, 2, 0) as usize;
            if let Some(Value::Object(obj)) = args.first() {
                let mut obj = obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = obj.kind {
                    for (i, &b) in val.iter().enumerate() {
                        if off + i < elems.len() {
                            elems[off + i] = Value::I32(b as i32);
                        }
                    }
                }
            }
            Value::I32((off + 4) as i32)
        }),
    );

    // writeDoubleBE / writeDoubleLE
    vm.register_host_fn(
        "node:buffer",
        "writeDoubleBE",
        Box::new(|_ctx, args| {
            let f = match args.get(1) {
                Some(Value::F64(f)) => *f,
                Some(Value::I32(n)) => *n as f64,
                _ => 0.0f64 };
            let val = f.to_be_bytes();
            let off = get_i32(args, 2, 0) as usize;
            if let Some(Value::Object(obj)) = args.first() {
                let mut obj = obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = obj.kind {
                    for (i, &b) in val.iter().enumerate() {
                        if off + i < elems.len() {
                            elems[off + i] = Value::I32(b as i32);
                        }
                    }
                }
            }
            Value::I32((off + 8) as i32)
        }),
    );
    vm.register_host_fn(
        "node:buffer",
        "writeDoubleLE",
        Box::new(|_ctx, args| {
            let f = match args.get(1) {
                Some(Value::F64(f)) => *f,
                Some(Value::I32(n)) => *n as f64,
                _ => 0.0f64 };
            let val = f.to_le_bytes();
            let off = get_i32(args, 2, 0) as usize;
            if let Some(Value::Object(obj)) = args.first() {
                let mut obj = obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = obj.kind {
                    for (i, &b) in val.iter().enumerate() {
                        if off + i < elems.len() {
                            elems[off + i] = Value::I32(b as i32);
                        }
                    }
                }
            }
            Value::I32((off + 8) as i32)
        }),
    );

    // swap64 — swap byte order in 8-byte groups
    vm.register_host_fn(
        "node:buffer",
        "swap64",
        Box::new(|_ctx, args| {
            let mut bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let mut i = 0;
            while i + 7 < bytes.len() {
                bytes.swap(i, i + 7);
                bytes.swap(i + 1, i + 6);
                bytes.swap(i + 2, i + 5);
                bytes.swap(i + 3, i + 4);
                i += 8;
            }
            bytes_to_buf(bytes)
        }),
    );

    // reverse — reverse byte order
    vm.register_host_fn(
        "node:buffer",
        "reverse",
        Box::new(|_ctx, args| {
            let mut bytes = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            bytes.reverse();
            bytes_to_buf(bytes)
        }),
    );
}
