//! `ecma:number` — ECMA-262 §21.1 Number + global numeric helpers.
//!
//! Exposes the JS-runtime numeric surface that Vybe-emitted .wasm
//! calls into: `Number.{isFinite,isNaN,isInteger,isSafeInteger,
//! parseInt,parseFloat}`, `Number.MAX_SAFE_INTEGER` etc., and
//! `Number.prototype.{toFixed,toString,valueOf}`.
//!
//! The merged `wasm:js-number` proposal (`proposals/
//! js-primitive-builtins/proposals/js-primitive-builtins/Overview.md`)
//! covers primitive type tests / boxed-primitive conversions
//! (`test`, `testI32`, `testU32`, `fromF64`, `fromI32`, `fromU32`,
//! `toF64`, `toI32`, `toU32`). Anything in the spec beyond that
//! lives here. The two layers are complementary: a runtime that
//! ships `wasm:js-number` natively + provides `ecma:number` shims
//! has the full ECMA-262 numeric surface.

use std::sync::{Arc, Mutex, OnceLock};
use vybe_bytecode::value::Object;
use vybe_bytecode::{HostContext, VM, Value};

static NUMBER_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();

pub(crate) fn shared_number_prototype() -> Value {
    Value::Object(
        NUMBER_PROTOTYPE
            .get_or_init(|| Arc::new(Mutex::new(Object::new())))
            .clone(),
    )
}

pub(crate) fn boxed_number(value: Value) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Number")));
    obj.properties.insert("__primitive".into(), value);
    obj.properties
        .insert("__proto__".into(), shared_number_prototype());
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn f_arg(args: &[Value], idx: usize) -> Option<f64> {
    match args.get(idx) {
        Some(Value::F64(n)) => Some(*n),
        Some(Value::I32(n)) => Some(*n as f64),
        Some(Value::I64(n)) => Some(*n as f64),
        Some(Value::Object(obj)) => {
            let primitive = {
                let locked = obj.lock().unwrap();
                if matches!(locked.properties.get("__type"), Some(Value::String(tag)) if tag.as_ref() == "Number")
                {
                    locked.properties.get("__primitive").cloned()
                } else {
                    None
                }
            };
            primitive.as_ref().map(coerce_to_number)
        }
        _ => None,
    }
}

fn s_arg(args: &[Value], idx: usize) -> String {
    match args.get(idx) {
        Some(Value::String(text)) => text.to_string(),
        Some(other) => format!("{}", other),
        None => String::new(),
    }
}

fn s_val(text: &str) -> Value {
    Value::String(Arc::from(text))
}

pub fn register(vm: &mut VM) {
    register_constants(vm);
    register_predicates(vm);
    register_parsers(vm);
    register_prototype(vm);
    register_constructor(vm);
}

// `Number(v)` — §21.1.1.1: ToNumber(v).
//
//   undefined → NaN, null → +0, true/false → 1/0, number → identity,
//   string → StringToNumber (parse with whitespace trim, "" → 0).
//   Other types fall through to NaN — boxed wrappers aren't supported.
fn register_constructor(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:number",
        "Number",
        Box::new(|ctx, args| {
            match coerce_to_number_with_context(ctx, args.first().unwrap_or(&Value::Undefined)) {
                Ok(n) => Value::F64(n),
                Err(error) => {
                    ctx.throw_value(error);
                    Value::Null
                }
            }
        }),
    );
    vm.register_host_fn(
        "ecma:number",
        "new",
        Box::new(|ctx, args| {
            match coerce_to_number_with_context(ctx, args.first().unwrap_or(&Value::Undefined)) {
                Ok(n) => boxed_number(Value::F64(n)),
                Err(error) => {
                    ctx.throw_value(error);
                    Value::Null
                }
            }
        }),
    );
}

fn coerce_to_number_with_context(ctx: &mut HostContext, value: &Value) -> Result<f64, Value> {
    let primitive = crate::ecma::value::to_primitive(ctx, value, "number");
    match primitive {
        // ECMA-262 §21.1.1.1: Number(bigint) converts to the nearest f64.
        // Throwing TypeError here is wrong — that only applies to mixed arithmetic.
        Value::BigInt(n) => Ok(n as f64),
        Value::Symbol(_) => Err(crate::ecma::error::new_error(
            "TypeError",
            "Cannot convert a Symbol value to a number",
        )),
        other => Ok(coerce_to_number(&other)),
    }
}

fn coerce_to_number(value: &Value) -> f64 {
    match value {
        Value::Null => 0.0,
        Value::Undefined => f64::NAN,
        Value::Bool(b) => {
            if *b {
                1.0
            } else {
                0.0
            }
        }
        Value::F64(n) => *n,
        Value::I32(n) => *n as f64,
        Value::I64(n) => *n as f64,
        Value::String(s) => parse_to_number(s),
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            match &o.kind {
                vybe_bytecode::value::ObjectKind::Array(elems) => {
                    let joined: Vec<String> = elems
                        .iter()
                        .map(|v| match v {
                            Value::Null | Value::Undefined => String::new(),
                            other => format!("{}", other),
                        })
                        .collect();
                    parse_to_number(&joined.join(","))
                }
                _ if matches!(o.properties.get("__type"), Some(Value::String(s)) if s.as_ref() == "Date") => {
                    o.properties
                        .get("__time")
                        .map(|v| v.as_f64())
                        .unwrap_or(f64::NAN)
                }
                _ => f64::NAN,
            }
        }
        _ => f64::NAN,
    }
}

/// ECMA-262 §7.1.4.1.1 StringToNumber — same trim + parse the
/// constructor uses for both String and Object→toString fallbacks.
fn parse_to_number(s: &str) -> f64 {
    let trimmed = s.trim();
    if trimmed.is_empty() {
        return 0.0;
    }

    let (sign, body) = if let Some(rest) = trimmed.strip_prefix('+') {
        (1.0, rest)
    } else if let Some(rest) = trimmed.strip_prefix('-') {
        (-1.0, rest)
    } else {
        (1.0, trimmed)
    };

    let parsed_prefixed =
        if let Some(rest) = body.strip_prefix("0x").or_else(|| body.strip_prefix("0X")) {
            i64::from_str_radix(rest, 16)
                .ok()
                .map(|value| sign * value as f64)
        } else if let Some(rest) = body.strip_prefix("0o").or_else(|| body.strip_prefix("0O")) {
            i64::from_str_radix(rest, 8)
                .ok()
                .map(|value| sign * value as f64)
        } else if let Some(rest) = body.strip_prefix("0b").or_else(|| body.strip_prefix("0B")) {
            i64::from_str_radix(rest, 2)
                .ok()
                .map(|value| sign * value as f64)
        } else {
            None
        };

    parsed_prefixed.unwrap_or_else(|| trimmed.parse::<f64>().unwrap_or(f64::NAN))
}

// ── Constants ─────────────────────────────────────────────────────
//
// ESM named imports and namespace reads should observe these as
// immutable value bindings, not zero-arg callables.

fn register_constants(vm: &mut VM) {
    // Duplicate as 0-arg host fns first (so CALL_IMPORT via function index works),
    // then register_host_value to overwrite the module record with ExportEntry::Value
    // so flatten_module_value_exports sees them and the compiler can inline them.
    vm.register_host_fn(
        "ecma:number",
        "MAX_SAFE_INTEGER",
        Box::new(|_ctx, _args| Value::F64(9007199254740991.0)),
    );
    vm.register_host_fn(
        "ecma:number",
        "MIN_SAFE_INTEGER",
        Box::new(|_ctx, _args| Value::F64(-9007199254740991.0)),
    );
    vm.register_host_fn(
        "ecma:number",
        "MAX_VALUE",
        Box::new(|_ctx, _args| Value::F64(f64::MAX)),
    );
    vm.register_host_fn(
        "ecma:number",
        "MIN_VALUE",
        Box::new(|_ctx, _args| Value::F64(f64::from_bits(1))),
    );
    vm.register_host_fn(
        "ecma:number",
        "EPSILON",
        Box::new(|_ctx, _args| Value::F64(f64::EPSILON)),
    );
    vm.register_host_fn(
        "ecma:number",
        "POSITIVE_INFINITY",
        Box::new(|_ctx, _args| Value::F64(f64::INFINITY)),
    );
    vm.register_host_fn(
        "ecma:number",
        "NEGATIVE_INFINITY",
        Box::new(|_ctx, _args| Value::F64(f64::NEG_INFINITY)),
    );
    vm.register_host_fn(
        "ecma:number",
        "NaN",
        Box::new(|_ctx, _args| Value::F64(f64::NAN)),
    );
    // Number.MAX_SAFE_INTEGER = 2^53 − 1.
    vm.register_host_value(
        "ecma:number",
        "MAX_SAFE_INTEGER",
        Value::F64(9007199254740991.0),
    );
    vm.register_host_value(
        "ecma:number",
        "MIN_SAFE_INTEGER",
        Value::F64(-9007199254740991.0),
    );
    // Number.MAX_VALUE / MIN_VALUE — largest / smallest representable
    // positive normal f64. MIN_VALUE is the smallest *positive* > 0,
    // not the most negative.
    vm.register_host_value("ecma:number", "MAX_VALUE", Value::F64(f64::MAX));
    vm.register_host_value("ecma:number", "MIN_VALUE", Value::F64(f64::from_bits(1)));
    vm.register_host_value("ecma:number", "EPSILON", Value::F64(f64::EPSILON));
    vm.register_host_value(
        "ecma:number",
        "POSITIVE_INFINITY",
        Value::F64(f64::INFINITY),
    );
    vm.register_host_value(
        "ecma:number",
        "NEGATIVE_INFINITY",
        Value::F64(f64::NEG_INFINITY),
    );
    vm.register_host_value("ecma:number", "NaN", Value::F64(f64::NAN));
}

// ── Predicates ────────────────────────────────────────────────────
//
// `Number.isFinite` / `Number.isNaN` are STRICT — they don't coerce.
// Non-Number arguments always return false. The global (unprefixed)
// `isFinite` / `isNaN` coerce first; those live separately.

fn to_f64_coerce(v: &Value) -> f64 {
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
        Value::String(s) => {
            let trimmed = s.trim();
            if trimmed.is_empty() {
                0.0
            } else {
                trimmed.parse::<f64>().unwrap_or(f64::NAN)
            }
        }
        _ => f64::NAN,
    }
}

fn register_predicates(vm: &mut VM) {
    // ecma:number:isFinite — STRICT, no coercion (Number.isFinite).
    vm.register_host_fn(
        "ecma:number",
        "isFinite",
        Box::new(|_ctx, args| match args.first() {
            Some(Value::F64(n)) => Value::Bool(n.is_finite()),
            Some(Value::I32(_)) => Value::Bool(true),
            _ => Value::Bool(false),
        }),
    );

    // ecma:number:isNaN — STRICT, no coercion (Number.isNaN).
    vm.register_host_fn(
        "ecma:number",
        "isNaN",
        Box::new(|_ctx, args| match args.first() {
            Some(Value::F64(n)) => Value::Bool(n.is_nan()),
            _ => Value::Bool(false),
        }),
    );

    // ecma:number:globalIsFinite — coerces (global isFinite).
    vm.register_host_fn(
        "ecma:number",
        "globalIsFinite",
        Box::new(|_ctx, args| {
            let n = args.first().map(to_f64_coerce).unwrap_or(f64::NAN);
            Value::Bool(n.is_finite())
        }),
    );

    // ecma:number:globalIsNaN — coerces (global isNaN).
    vm.register_host_fn(
        "ecma:number",
        "globalIsNaN",
        Box::new(|_ctx, args| {
            let n = args.first().map(to_f64_coerce).unwrap_or(f64::NAN);
            Value::Bool(n.is_nan())
        }),
    );

    vm.register_host_fn(
        "ecma:number",
        "isInteger",
        Box::new(|_ctx, args| match args.first() {
            Some(Value::F64(n)) => Value::Bool(n.is_finite() && n.fract() == 0.0),
            Some(Value::I32(_)) => Value::Bool(true),
            _ => Value::Bool(false),
        }),
    );

    vm.register_host_fn(
        "ecma:number",
        "isSafeInteger",
        Box::new(|_ctx, args| {
            let n = match args.first() {
                Some(Value::F64(n)) => *n,
                Some(Value::I32(n)) => *n as f64,
                _ => return Value::Bool(false),
            };
            Value::Bool(n.is_finite() && n.fract() == 0.0 && n.abs() <= 9007199254740991.0)
        }),
    );
}

// ── Parsers ───────────────────────────────────────────────────────
//
// `Number.parseInt` / `Number.parseFloat` — same behaviour as the
// global `parseInt` / `parseFloat` (ECMA-262 §21.1.2.{12,13}).

fn register_parsers(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:number",
        "parseInt",
        Box::new(|_ctx, args| {
            let input = s_arg(args, 0);
            let radix = match args.get(1) {
                Some(Value::F64(n)) if *n != 0.0 => *n as u32,
                Some(Value::I32(n)) if *n != 0 => *n as u32,
                _ => 0,
            };
            Value::F64(parse_int_ecma(&input, radix))
        }),
    );

    vm.register_host_fn(
        "ecma:number",
        "parseFloat",
        Box::new(|_ctx, args| {
            let input = s_arg(args, 0);
            Value::F64(parse_float_ecma(&input))
        }),
    );
}

/// ECMA-262 §19.2.5 ParseInt: skip leading whitespace, optional
/// sign, consume the longest radix-valid prefix, return its parsed
/// value (or NaN if the prefix is empty).
fn parse_int_ecma(input: &str, radix: u32) -> f64 {
    let trimmed = input.trim_start();
    if trimmed.is_empty() {
        return f64::NAN;
    }
    let (sign, rest) = match trimmed.as_bytes()[0] {
        b'+' => (1.0_f64, &trimmed[1..]),
        b'-' => (-1.0_f64, &trimmed[1..]),
        _ => (1.0_f64, trimmed),
    };
    if rest.is_empty() {
        return f64::NAN;
    }

    // 0x / 0X auto-detects hex when radix is unspecified or 16.
    let (effective_radix, body) = if (radix == 0 || radix == 16)
        && rest.len() >= 2
        && rest.as_bytes()[0] == b'0'
        && (rest.as_bytes()[1] == b'x' || rest.as_bytes()[1] == b'X')
    {
        (16u32, &rest[2..])
    } else if radix == 0 {
        (10u32, rest)
    } else {
        (radix, rest)
    };

    let mut acc: u128 = 0;
    let mut consumed = 0usize;
    for (i, ch) in body.chars().enumerate() {
        let digit = match ch.to_digit(effective_radix) {
            Some(d) => d,
            None => break,
        };
        acc = acc
            .saturating_mul(effective_radix as u128)
            .saturating_add(digit as u128);
        consumed = i + 1;
    }
    if consumed == 0 {
        return f64::NAN;
    }
    sign * (acc as f64)
}

/// ECMA-262 §19.2.5 ParseFloat: skip leading whitespace, parse the
/// longest substring that looks like a float; return NaN if the
/// prefix isn't a valid number start.
fn parse_float_ecma(input: &str) -> f64 {
    let trimmed = input.trim_start();
    if trimmed.is_empty() {
        return f64::NAN;
    }

    // Try progressively shorter prefixes until one parses. Slow but
    // correct; ECMA's algorithm is "longest valid prefix".
    let bytes = trimmed.as_bytes();
    let mut end = bytes.len();
    while end > 0 {
        if let Ok(n) = trimmed[..end].parse::<f64>() {
            return n;
        }
        end -= 1;
    }
    f64::NAN
}

// ── Number.prototype methods ──────────────────────────────────────
//
// Method receivers are the first arg per Component-Model `[method]`
// convention. JS callers route `(42).toFixed(2)` through Vybe's
// compiler as `ecma:number.toFixed(42, 2)`.

fn register_prototype(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:number",
        "toFixed",
        Box::new(|_ctx, args| {
            let n = f_arg(args, 0).unwrap_or(0.0);
            let digits = match args.get(1) {
                Some(Value::F64(d)) => *d as usize,
                Some(Value::I32(d)) => *d as usize,
                _ => 0,
            };
            s_val(&format!("{:.1$}", n, digits))
        }),
    );

    vm.register_host_fn(
        "ecma:number",
        "toString",
        Box::new(|_ctx, args| {
            let n = f_arg(args, 0).unwrap_or(0.0);
            let radix = match args.get(1) {
                Some(Value::F64(r)) => *r as u32,
                Some(Value::I32(r)) => *r as u32,
                _ => 10,
            };
            if radix == 10 {
                // Integer-valued floats print without trailing ".0" per JS.
                if n.is_finite() && n.fract() == 0.0 {
                    return s_val(&format!("{}", n as i64));
                }
                return s_val(&format!("{}", n));
            }
            if !(2..=36).contains(&radix) || !n.is_finite() {
                return s_val(&format!("{}", n));
            }
            // Integer-only radix conversion (ECMA's algorithm for
            // fractional values is quite involved; this covers the
            // common case.)
            let int_value = n as i64;
            let negative = int_value < 0;
            let mut value = (int_value as i128).unsigned_abs();
            if value == 0 {
                return s_val("0");
            }
            let mut out = String::new();
            while value > 0 {
                let digit = (value % radix as u128) as u32;
                let ch = char::from_digit(digit, radix).unwrap_or('?');
                out.insert(0, ch);
                value /= radix as u128;
            }
            if negative {
                out.insert(0, '-');
            }
            s_val(&out)
        }),
    );

    vm.register_host_fn(
        "ecma:number",
        "valueOf",
        Box::new(|_ctx, args| Value::F64(f_arg(args, 0).unwrap_or(0.0))),
    );
    vm.register_host_fn(
        "ecma:number",
        "toLocaleString",
        Box::new(|_ctx, args| {
            let n = match args.first() {
                Some(Value::F64(f)) => *f,
                Some(Value::I32(i)) => *i as f64,
                _ => return Value::String(Arc::from("0")),
            };
            if n.is_finite() && n.fract() == 0.0 {
                Value::String(Arc::from(format!("{}", n as i64).as_str()))
            } else {
                Value::String(Arc::from(format!("{}", n).as_str()))
            }
        }),
    );

    vm.register_host_fn(
        "ecma:number",
        "toExponential",
        Box::new(|_ctx, args| {
            let n = f_arg(args, 0).unwrap_or(0.0);
            let raw = match args.get(1) {
                Some(Value::F64(d)) => format!("{:.1$e}", n, *d as usize),
                Some(Value::I32(d)) => format!("{:.1$e}", n, *d as usize),
                _ => format!("{:e}", n),
            };
            let parts: Vec<&str> = raw.splitn(2, 'e').collect();
            if parts.len() == 2 {
                let exp: i32 = parts[1].parse().unwrap_or(0);
                let sign = if exp >= 0 { "+" } else { "" };
                s_val(&format!("{}e{}{}", parts[0], sign, exp))
            } else {
                s_val(&raw)
            }
        }),
    );

    vm.register_host_fn(
        "ecma:number",
        "toPrecision",
        Box::new(|_ctx, args| {
            let n = f_arg(args, 0).unwrap_or(0.0);
            let prec = match args.get(1) {
                Some(Value::F64(p)) => *p as usize,
                Some(Value::I32(p)) => *p as usize,
                _ => return s_val(&format!("{}", n)),
            };
            if prec == 0 {
                return s_val(&format!("{}", n));
            }
            if n == 0.0 {
                if prec <= 1 {
                    return s_val("0");
                }
                return s_val(&format!("0.{:0<1$}", "", prec - 1));
            }
            let abs_n = n.abs();
            let e = abs_n.log10().floor() as i32;
            if e >= prec as i32 || e < -6 {
                // exponential notation
                let raw = format!("{:.1$e}", n, prec - 1);
                let parts: Vec<&str> = raw.splitn(2, 'e').collect();
                if parts.len() == 2 {
                    let exp: i32 = parts[1].parse().unwrap_or(0);
                    let sign = if exp >= 0 { "+" } else { "" };
                    return s_val(&format!("{}e{}{}", parts[0], sign, exp));
                }
                return s_val(&raw);
            }
            let decimal_places = ((prec as i32 - 1 - e).max(0)) as usize;
            s_val(&format!("{:.1$}", n, decimal_places))
        }),
    );
}
