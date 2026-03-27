use vybe_bytecode::{VM, Value};

pub fn register(vm: &mut VM) {
    vm.register_host_fn("vybe:convert", "parseInt", Box::new(|args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        match s.trim().parse::<i64>() {
            Ok(n) => Value::F64(n as f64),
            Err(_) => match s.trim().parse::<f64>() {
                Ok(n) => Value::F64(n.trunc()),
                Err(_) => Value::F64(f64::NAN),
            }
        }
    }));
    vm.register_host_fn("vybe:convert", "parseFloat", Box::new(|args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        Value::F64(s.trim().parse::<f64>().unwrap_or(f64::NAN))
    }));
    vm.register_host_fn("vybe:convert", "toString", Box::new(|args: &[Value]| {
        Value::String(std::rc::Rc::from(format!("{}", args.first().unwrap_or(&Value::Null)).as_str()))
    }));
    vm.register_host_fn("vybe:convert", "isNaN", Box::new(|args: &[Value]| {
        Value::Bool(args.first().map(|v| v.as_f64().is_nan()).unwrap_or(true))
    }));
    vm.register_host_fn("vybe:convert", "isFinite", Box::new(|args: &[Value]| {
        Value::Bool(args.first().map(|v| v.as_f64().is_finite()).unwrap_or(false))
    }));

    // Number.isInteger(n)
    vm.register_host_fn("vybe:convert", "isInteger", Box::new(|args: &[Value]| {
        let n = args.first().map(|v| v.as_f64()).unwrap_or(f64::NAN);
        Value::Bool(!n.is_nan() && !n.is_infinite() && n == n.trunc())
    }));

    // btoa(string) — base64 encode
    vm.register_host_fn("vybe:convert", "btoa", Box::new(|args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        use std::rc::Rc;
        Value::String(Rc::from(base64_encode(s.as_bytes()).as_str()))
    }));

    // atob(base64) — base64 decode
    vm.register_host_fn("vybe:convert", "atob", Box::new(|args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        use std::rc::Rc;
        match base64_decode(&s) {
            Some(bytes) => Value::String(Rc::from(String::from_utf8_lossy(&bytes).as_ref())),
            None => Value::Null,
        }
    }));

    // encodeURIComponent
    vm.register_host_fn("vybe:convert", "encodeURIComponent", Box::new(|args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        use std::rc::Rc;
        let encoded: String = s.chars().map(|c| {
            if c.is_ascii_alphanumeric() || "-_.!~*'()".contains(c) {
                c.to_string()
            } else {
                let mut buf = [0u8; 4];
                c.encode_utf8(&mut buf);
                buf[..c.len_utf8()].iter().map(|b| format!("%{:02X}", b)).collect()
            }
        }).collect();
        Value::String(Rc::from(encoded.as_str()))
    }));

    // decodeURIComponent
    vm.register_host_fn("vybe:convert", "decodeURIComponent", Box::new(|args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        use std::rc::Rc;
        let mut result = String::new();
        let bytes = s.as_bytes();
        let mut i = 0;
        while i < bytes.len() {
            if bytes[i] == b'%' && i + 2 < bytes.len() {
                if let Ok(byte) = u8::from_str_radix(&s[i+1..i+3], 16) {
                    result.push(byte as char);
                    i += 3;
                    continue;
                }
            }
            result.push(bytes[i] as char);
            i += 1;
        }
        Value::String(Rc::from(result.as_str()))
    }));

    // --- VB-compatible conversion functions ---

    // val(str) → parse as number, 0 on failure
    vm.register_host_fn("vybe:convert", "val", Box::new(|args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        Value::F64(s.trim().parse::<f64>().unwrap_or(0.0))
    }));

    // isNothing(value) → true if null
    vm.register_host_fn("vybe:convert", "isNothing", Box::new(|args: &[Value]| {
        Value::Bool(matches!(args.first().unwrap_or(&Value::Null), Value::Null))
    }));

    // isNumeric(value) → true if can be parsed as number
    vm.register_host_fn("vybe:convert", "isNumeric", Box::new(|args: &[Value]| {
        match args.first().unwrap_or(&Value::Null) {
            Value::F64(n) => Value::Bool(!n.is_nan()),
            Value::I32(_) | Value::I64(_) => Value::Bool(true),
            Value::String(s) => Value::Bool(s.trim().parse::<f64>().is_ok()),
            _ => Value::Bool(false),
        }
    }));

    // cint(value) → floor to integer
    vm.register_host_fn("vybe:convert", "cint", Box::new(|args: &[Value]| {
        Value::F64(args.first().map(|v| v.as_f64().floor()).unwrap_or(0.0))
    }));

    // cdbl(value) → to double (identity for numbers, parse for strings)
    vm.register_host_fn("vybe:convert", "cdbl", Box::new(|args: &[Value]| {
        match args.first().unwrap_or(&Value::Null) {
            Value::String(s) => Value::F64(s.trim().parse::<f64>().unwrap_or(0.0)),
            v => Value::F64(v.as_f64()),
        }
    }));

    // cbool(value) → to boolean
    vm.register_host_fn("vybe:convert", "cbool", Box::new(|args: &[Value]| {
        match args.first().unwrap_or(&Value::Null) {
            Value::Null => Value::Bool(false),
            Value::Bool(b) => Value::Bool(*b),
            Value::F64(n) => Value::Bool(*n != 0.0),
            Value::I32(n) => Value::Bool(*n != 0),
            Value::I64(n) => Value::Bool(*n != 0),
            Value::String(s) => Value::Bool(!s.is_empty() && s.to_lowercase() != "false"),
            Value::Object(_) => Value::Bool(true),
        }
    }));
}

fn base64_encode(data: &[u8]) -> String {
    const CHARS: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut result = String::new();
    for chunk in data.chunks(3) {
        let b0 = chunk[0] as u32;
        let b1 = if chunk.len() > 1 { chunk[1] as u32 } else { 0 };
        let b2 = if chunk.len() > 2 { chunk[2] as u32 } else { 0 };
        let n = (b0 << 16) | (b1 << 8) | b2;
        result.push(CHARS[((n >> 18) & 63) as usize] as char);
        result.push(CHARS[((n >> 12) & 63) as usize] as char);
        if chunk.len() > 1 { result.push(CHARS[((n >> 6) & 63) as usize] as char); } else { result.push('='); }
        if chunk.len() > 2 { result.push(CHARS[(n & 63) as usize] as char); } else { result.push('='); }
    }
    result
}

fn base64_decode(s: &str) -> Option<Vec<u8>> {
    const DECODE: [u8; 128] = {
        let mut t = [255u8; 128];
        let mut i = 0u8;
        while i < 26 { t[(b'A' + i) as usize] = i; i += 1; }
        i = 0;
        while i < 26 { t[(b'a' + i) as usize] = 26 + i; i += 1; }
        i = 0;
        while i < 10 { t[(b'0' + i) as usize] = 52 + i; i += 1; }
        t[b'+' as usize] = 62;
        t[b'/' as usize] = 63;
        t
    };
    let bytes: Vec<u8> = s.bytes().filter(|&b| b != b'=' && b != b'\n' && b != b'\r').collect();
    let mut result = Vec::new();
    for chunk in bytes.chunks(4) {
        if chunk.len() < 2 { break; }
        let a = *DECODE.get(chunk[0] as usize)? as u32;
        let b = *DECODE.get(chunk[1] as usize)? as u32;
        result.push(((a << 2) | (b >> 4)) as u8);
        if chunk.len() > 2 {
            let c = *DECODE.get(chunk[2] as usize)? as u32;
            result.push((((b & 0xF) << 4) | (c >> 2)) as u8);
            if chunk.len() > 3 {
                let d = *DECODE.get(chunk[3] as usize)? as u32;
                result.push((((c & 0x3) << 6) | d) as u8);
            }
        }
    }
    Some(result)
}
