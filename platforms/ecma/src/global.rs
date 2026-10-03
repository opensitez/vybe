use std::borrow::Cow;
use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

const HEX_UPPER: &[u8; 16] = b"0123456789ABCDEF";

#[inline]
fn push_percent_byte(out: &mut String, byte: u8) {
    out.push('%');
    out.push(HEX_UPPER[(byte >> 4) as usize] as char);
    out.push(HEX_UPPER[(byte & 0x0f) as usize] as char);
}

#[inline]
fn hex_nibble(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

#[inline]
fn arg_string(value: Option<&Value>) -> Option<Cow<'_, str>> {
    match value {
        Some(Value::String(text)) => Some(Cow::Borrowed(text.as_ref())),
        Some(value) => Some(crate::keys::value_display_cow(value)),
        None => None,
    }
}

#[inline]
fn owned_string_value(text: String) -> Value {
    crate::keys::owned_string_value(text)
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:global",
        "isNaN",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = to_number(args.first().unwrap_or(&Value::Undefined));
            Value::Bool(n.is_nan())
        }),
    );

    vm.register_host_fn(
        "ecma:global",
        "isFinite",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = to_number(args.first().unwrap_or(&Value::Undefined));
            Value::Bool(n.is_finite())
        }),
    );

    vm.register_host_fn(
        "ecma:global",
        "parseInt",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(s) = arg_string(args.first()) else {
                return Value::F64(f64::NAN);
            };
            let radix = args.get(1).map(|v| v.as_i32()).unwrap_or(10).max(2).min(36) as u32;
            let s = s.trim_start();
            let (neg, s) = if s.starts_with('-') {
                (true, &s[1..])
            } else if s.starts_with('+') {
                (false, &s[1..])
            } else {
                (false, s)
            };
            let s = if radix == 16 && (s.starts_with("0x") || s.starts_with("0X")) {
                &s[2..]
            } else {
                s
            };
            let mut result: i64 = 0;
            let mut any = false;
            for c in s.chars() {
                let d = c.to_digit(radix);
                match d {
                    Some(d) => {
                        result = result.wrapping_mul(radix as i64).wrapping_add(d as i64);
                        any = true;
                    }
                    None => break,
                }
            }
            if !any {
                return Value::F64(f64::NAN);
            }
            Value::F64(if neg { -(result as f64) } else { result as f64 })
        }),
    );

    vm.register_host_fn(
        "ecma:global",
        "parseFloat",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let s = match args.first() {
                Some(Value::String(s)) => s.trim(),
                Some(Value::F64(n)) => return Value::F64(*n),
                Some(Value::I32(n)) => return Value::F64(*n as f64),
                _ => return Value::F64(f64::NAN),
            };
            match s.parse::<f64>() {
                Ok(n) => Value::F64(n),
                Err(_) => Value::F64(f64::NAN),
            }
        }),
    );

    vm.register_free_fn(
        "ecma:global",
        "eval",
        Box::new(
            |_ctx: &mut HostContext, args: &[Value]| match _ctx.user_args(args, 0).first() {
                // ⛔ `user_args`: INDIRECT eval (`const g = eval; g("3+4")`)
                // reaches this host fn as a VALUE through a dynamic call, and
                // under `ReceiverAbi::Parameter` that call puts a receiver at
                // argument 0. Reading the source at a fixed index picked up the
                // receiver and returned `undefined`. Direct `eval(...)` is
                // compiled specially and was unaffected, which is why only the
                // indirect form failed.
                Some(Value::String(s)) => {
                    if let Ok(n) = s.trim().parse::<f64>() {
                        Value::F64(n)
                    } else {
                        Value::Undefined
                    }
                }
                Some(v) => v.clone(),
                None => Value::Undefined,
            },
        ),
    );

    vm.register_host_fn(
        "ecma:global",
        "globalThis",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let obj = Object::new();
            Value::Object(std::sync::Arc::new(std::sync::Mutex::new(obj)))
        }),
    );

    vm.register_host_fn(
        "ecma:global",
        "Infinity",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::F64(f64::INFINITY)),
    );

    vm.register_host_fn(
        "ecma:global",
        "NaN",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::F64(f64::NAN)),
    );

    vm.register_host_fn(
        "ecma:global",
        "undefined",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::Undefined),
    );

    vm.register_host_fn(
        "ecma:global",
        "encodeURIComponent",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let s = match arg_string(args.first()) {
                Some(s) => s,
                None => return Value::Undefined,
            };
            let mut encoded = String::with_capacity(s.len());
            for c in s.chars() {
                if c.is_alphanumeric()
                    || matches!(c, '-' | '_' | '.' | '!' | '~' | '*' | '\'' | '(' | ')')
                {
                    encoded.push(c);
                } else {
                    let mut buf = [0u8; 4];
                    let len = c.encode_utf8(&mut buf).len();
                    for &byte in &buf[..len] {
                        push_percent_byte(&mut encoded, byte);
                    }
                }
            }
            owned_string_value(encoded)
        }),
    );

    vm.register_host_fn(
        "ecma:global",
        "decodeURIComponent",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let s = match arg_string(args.first()) {
                Some(s) => s,
                None => return Value::Undefined,
            };
            owned_string_value(decode_uri(&s))
        }),
    );

    vm.register_host_fn(
        "ecma:global",
        "encodeURI",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let s = match arg_string(args.first()) {
                Some(s) => s,
                None => return Value::Undefined,
            };
            let mut encoded = String::with_capacity(s.len());
            for c in s.chars() {
                if c.is_alphanumeric()
                    || matches!(
                        c,
                        '-' | '_'
                            | '.'
                            | '!'
                            | '~'
                            | '*'
                            | '\''
                            | '('
                            | ')'
                            | ';'
                            | ','
                            | '/'
                            | '?'
                            | ':'
                            | '@'
                            | '&'
                            | '='
                            | '+'
                            | '$'
                            | '#'
                    )
                {
                    encoded.push(c);
                } else {
                    let mut buf = [0u8; 4];
                    let len = c.encode_utf8(&mut buf).len();
                    for &byte in &buf[..len] {
                        push_percent_byte(&mut encoded, byte);
                    }
                }
            }
            owned_string_value(encoded)
        }),
    );

    vm.register_host_fn(
        "ecma:global",
        "decodeURI",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let s = match arg_string(args.first()) {
                Some(s) => s,
                None => return Value::Undefined,
            };
            owned_string_value(decode_uri(&s))
        }),
    );
}

fn to_number(v: &Value) -> f64 {
    match v {
        Value::F64(n) => *n,
        Value::I32(n) => *n as f64,
        Value::I64(n) => *n as f64,
        Value::Bool(b) => {
            if *b {
                1.0
            } else {
                0.0
            }
        }
        Value::Null => 0.0,
        Value::Undefined => f64::NAN,
        Value::String(s) => s.trim().parse::<f64>().unwrap_or(f64::NAN),
        _ => f64::NAN,
    }
}

fn decode_uri(s: &str) -> String {
    let bytes = s.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            if let (Some(h), Some(l)) = (hex_nibble(bytes[i + 1]), hex_nibble(bytes[i + 2])) {
                out.push((h << 4) | l);
                i += 3;
                continue;
            }
        }
        out.push(bytes[i]);
        i += 1;
    }
    String::from_utf8_lossy(&out).into_owned()
}
