//! PHP string/hash helpers that do not match ECMA string semantics exactly.

use hmac::{Hmac, Mac};
use sha2::{Digest, Sha256};
use std::sync::Arc;
use vybe_bytecode::{HostContext, VM, Value};

type HmacSha256 = Hmac<Sha256>;

const BASE64_CHARS: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

fn s_arg(args: &[Value], idx: usize) -> String {
    args.get(idx).map(|v| format!("{v}")).unwrap_or_default()
}

fn b_arg(args: &[Value], idx: usize) -> bool {
    args.get(idx).map(|v| v.as_bool()).unwrap_or(false)
}

fn s_val(text: impl AsRef<str>) -> Value {
    Value::String(Arc::from(text.as_ref()))
}

fn hex(bytes: &[u8]) -> String {
    const HEX: &[u8] = b"0123456789abcdef";
    let mut out = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        out.push(HEX[(byte >> 4) as usize] as char);
        out.push(HEX[(byte & 0x0f) as usize] as char);
    }
    out
}

fn base64_decode(s: &str, strict: bool) -> Option<Vec<u8>> {
    let mut filtered = Vec::new();
    for byte in s.bytes() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        if byte == b'=' || BASE64_CHARS.contains(&byte) {
            filtered.push(byte);
        } else if strict {
            return None;
        }
    }
    if filtered.len() % 4 == 1 {
        return None;
    }
    while filtered.len() % 4 != 0 {
        filtered.push(b'=');
    }

    let value = |byte: u8| -> Option<u32> {
        match byte {
            b'A'..=b'Z' => Some((byte - b'A') as u32),
            b'a'..=b'z' => Some((26 + byte - b'a') as u32),
            b'0'..=b'9' => Some((52 + byte - b'0') as u32),
            b'+' => Some(62),
            b'/' => Some(63),
            _ => None,
        }
    };

    let mut out = Vec::new();
    let chunk_count = filtered.len() / 4;
    for (idx, chunk) in filtered.chunks(4).enumerate() {
        let last = idx + 1 == chunk_count;
        if (chunk[0] == b'=') || (chunk[1] == b'=') {
            return None;
        }
        let a = value(chunk[0])?;
        let b = value(chunk[1])?;
        out.push(((a << 2) | (b >> 4)) as u8);

        if chunk[2] == b'=' {
            if chunk[3] != b'=' || !last {
                return None;
            }
            continue;
        }
        let c = value(chunk[2])?;
        out.push((((b & 0x0f) << 4) | (c >> 2)) as u8);

        if chunk[3] == b'=' {
            if !last {
                return None;
            }
            continue;
        }
        let d = value(chunk[3])?;
        out.push((((c & 0x03) << 6) | d) as u8);
    }
    Some(out)
}

fn latin1(bytes: &[u8]) -> String {
    bytes.iter().map(|byte| char::from(*byte)).collect()
}

fn char_to_byte_index(s: &str, char_index: usize) -> usize {
    s.char_indices()
        .nth(char_index)
        .map(|(idx, _)| idx)
        .unwrap_or_else(|| s.len())
}

fn byte_to_char_index(s: &str, byte_index: usize) -> usize {
    s[..byte_index].chars().count()
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "php:string",
        "strlen",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| Value::F64(s_arg(args, 0).len() as f64)),
    );

    vm.register_host_fn(
        "php:string",
        "mb_strpos",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let haystack = s_arg(args, 0);
            let needle = s_arg(args, 1);
            let offset = args.get(2).map(|v| v.as_f64() as isize).unwrap_or(0);
            let start_char = if offset < 0 {
                haystack.chars().count().saturating_sub((-offset) as usize)
            } else {
                offset as usize
            };
            let start_byte = char_to_byte_index(&haystack, start_char);
            match haystack[start_byte..].find(&needle) {
                Some(rel) => Value::F64(byte_to_char_index(&haystack, start_byte + rel) as f64),
                None => Value::Bool(false),
            }
        }),
    );

    vm.register_host_fn(
        "php:string",
        "mb_strrpos",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let haystack = s_arg(args, 0);
            let needle = s_arg(args, 1);
            match haystack.rfind(&needle) {
                Some(idx) => Value::F64(byte_to_char_index(&haystack, idx) as f64),
                None => Value::Bool(false),
            }
        }),
    );

    vm.register_host_fn(
        "php:string",
        "base64_decode",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            match base64_decode(&s_arg(args, 0), b_arg(args, 1)) {
                Some(bytes) => s_val(latin1(&bytes)),
                None => Value::Bool(false),
            }
        }),
    );

    vm.register_host_fn(
        "php:string",
        "hash",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if s_arg(args, 0).eq_ignore_ascii_case("sha256") {
                s_val(hex(&Sha256::digest(s_arg(args, 1).as_bytes())))
            } else {
                Value::Bool(false)
            }
        }),
    );

    vm.register_host_fn(
        "php:string",
        "hash_hmac",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if !s_arg(args, 0).eq_ignore_ascii_case("sha256") {
                return Value::Bool(false);
            }
            let mut mac = HmacSha256::new_from_slice(s_arg(args, 2).as_bytes()).unwrap();
            mac.update(s_arg(args, 1).as_bytes());
            s_val(hex(&mac.finalize().into_bytes()))
        }),
    );
}
