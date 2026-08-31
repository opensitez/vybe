use std::sync::Arc;
use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

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
            let s = match args.first() {
                Some(Value::String(s)) => s.trim().to_string(),
                Some(v) => format!("{}", v),
                None => return Value::F64(f64::NAN),
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
                Some(Value::String(s)) => s.trim().to_string(),
                Some(Value::F64(n)) => return Value::F64(*n),
                Some(Value::I32(n)) => return Value::F64(*n as f64),
                _ => return Value::F64(f64::NAN),
            };
            match s.trim().parse::<f64>() {
                Ok(n) => Value::F64(n),
                Err(_) => Value::F64(f64::NAN),
            }
        }),
    );

    vm.register_free_fn(
        "ecma:global",
        "eval",
        Box::new(
            |_ctx: &mut HostContext, args: &[Value]| match _ctx
                .user_args(args, 0)
                .first()
            {
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
            let s = match args.first() {
                Some(v) => format!("{}", v),
                None => return Value::Undefined,
            };
            let encoded: String = s
                .chars()
                .map(|c| {
                    if c.is_alphanumeric()
                        || matches!(c, '-' | '_' | '.' | '!' | '~' | '*' | '\'' | '(' | ')')
                    {
                        c.to_string()
                    } else {
                        c.to_string()
                            .bytes()
                            .map(|b| format!("%{:02X}", b))
                            .collect()
                    }
                })
                .collect();
            Value::String(Arc::from(encoded.as_str()))
        }),
    );

    vm.register_host_fn(
        "ecma:global",
        "decodeURIComponent",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let s = match args.first() {
                Some(v) => format!("{}", v),
                None => return Value::Undefined,
            };
            Value::String(Arc::from(decode_uri(&s).as_str()))
        }),
    );

    vm.register_host_fn(
        "ecma:global",
        "encodeURI",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let s = match args.first() {
                Some(v) => format!("{}", v),
                None => return Value::Undefined,
            };
            let encoded: String = s
                .chars()
                .map(|c| {
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
                        c.to_string()
                    } else {
                        c.to_string()
                            .bytes()
                            .map(|b| format!("%{:02X}", b))
                            .collect()
                    }
                })
                .collect();
            Value::String(Arc::from(encoded.as_str()))
        }),
    );

    vm.register_host_fn(
        "ecma:global",
        "decodeURI",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let s = match args.first() {
                Some(v) => format!("{}", v),
                None => return Value::Undefined,
            };
            Value::String(Arc::from(decode_uri(&s).as_str()))
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
    let bytes: Vec<u8> = s.as_bytes().to_vec();
    let mut out = Vec::new();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            if let (Some(h), Some(l)) = (
                (bytes[i + 1] as char).to_digit(16),
                (bytes[i + 2] as char).to_digit(16),
            ) {
                out.push((h * 16 + l) as u8);
                i += 3;
                continue;
            }
        }
        out.push(bytes[i]);
        i += 1;
    }
    String::from_utf8_lossy(&out).into_owned()
}
