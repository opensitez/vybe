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

use std::borrow::Cow;
use std::sync::{Arc, Mutex, OnceLock};
use vybe_runtime::value::Object;
use vybe_runtime::vm::HostFnDecl;
use vybe_runtime::{FuncSig, HostContext, Param, VM, ValType, Value};

/// Declare an `ecma:number` function — same closure, plus the signature.
fn number_fn(
    vm: &mut VM,
    name: &str,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(
        HostFnDecl::new("ecma:number", name, call).with_sig(FuncSig {
            name: name.to_string(),
            params: Param::unnamed_list(params),
            results,
        }),
    );
}
/// Register a FREE FUNCTION — one whose type has no receiver parameter.
///
/// `encodeURIComponent(s)` takes a string and nothing else; it is not a method
/// being spelled as a call the way `padStart(s, 4)` is. Declaring that keeps
/// the shape identical whether it is called directly or handed to `map`.
fn number_fn_free(
    vm: &mut VM,
    name: &str,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(
        HostFnDecl::new("ecma:number", name, call)
            .with_sig(FuncSig {
                name: name.to_string(),
                params: Param::unnamed_list(params),
                results,
            })
            .without_receiver(),
    );
}

/// `Number.MAX_VALUE`, `Number.NaN`, … — the callable form of a constant,
/// registered alongside the `register_host_value` form below. Takes nothing.
fn number_const(
    vm: &mut VM,
    name: &str,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    number_fn(vm, name, vec![], vec![ValType::F64], call);
}

static NUMBER_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();

pub fn shared_number_prototype() -> Value {
    Value::Object(
        NUMBER_PROTOTYPE
            .get_or_init(|| vybe_runtime::heap::alloc(Object::new()))
            .clone(),
    )
}

pub fn boxed_number(value: Value) -> Value {
    let mut obj = Object::new();
    obj.properties.reserve(3);
    obj.properties
        .insert("__type".into(), crate::keys::string_value("Number"));
    obj.properties.insert("__primitive".into(), value);
    obj.properties
        .insert("__proto__".into(), shared_number_prototype());
    Value::Object(vybe_runtime::heap::alloc(obj))
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

fn s_arg(args: &[Value], idx: usize) -> Cow<'_, str> {
    match args.get(idx) {
        Some(Value::String(text)) => Cow::Borrowed(text.as_ref()),
        Some(other) => crate::keys::value_display_cow(other),
        None => Cow::Borrowed(""),
    }
}

fn s_val(text: &str) -> Value {
    crate::keys::string_value(text)
}

#[inline]
fn s_owned(text: String) -> Value {
    crate::keys::owned_string_value(text)
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
    vm.register_free_fn(
        "ecma:number",
        "Number",
        Box::new(|ctx, args| {
            let value = args.first().unwrap_or(&Value::Undefined);
            match coerce_to_number_with_context(ctx, value) {
                Ok(n) => Value::F64(n),
                Err(error) => {
                    ctx.throw_value(error);
                    Value::Null
                }
            }
        }),
    );
    vm.register_free_fn(
        "ecma:number",
        "new",
        Box::new(|ctx, args| {
            let value = args.first().unwrap_or(&Value::Undefined);
            match coerce_to_number_with_context(ctx, value) {
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
    let primitive = crate::value::to_primitive(ctx, value, "number");
    match primitive {
        // ECMA-262 §21.1.1.1: Number(bigint) converts to the nearest f64.
        // Throwing TypeError here is wrong — that only applies to mixed arithmetic.
        Value::BigInt(n) => Ok(n.to_f64()),
        Value::Symbol(_) => Err(crate::error::new_error(
            ctx,
            "TypeError",
            "Cannot convert a Symbol value to a number",
        )),
        other => Ok(coerce_to_number(&other)),
    }
}

/// ECMA ToLength uses ToNumber, whose BigInt rule differs from Number(bigint).
pub(crate) fn try_to_length(ctx: &mut HostContext, value: &Value) -> Result<u64, Value> {
    let primitive = crate::value::try_to_primitive(ctx, value, "number")?;
    if matches!(primitive, Value::BigInt(_) | Value::Symbol(_)) {
        return Err(crate::error::new_error(
            ctx,
            "TypeError",
            "Cannot convert length to a Number",
        ));
    }
    let number = coerce_to_number(&primitive);
    if number.is_nan() || number <= 0.0 {
        return Ok(0);
    }
    Ok(number.trunc().min(9_007_199_254_740_991.0) as u64)
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
                vybe_runtime::value::ObjectKind::Array(elems) => {
                    let mut joined = String::with_capacity(elems.len().saturating_mul(4));
                    for (index, value) in elems.iter().enumerate() {
                        if index != 0 {
                            joined.push(',');
                        }
                        if !matches!(value, Value::Null | Value::Undefined) {
                            match value {
                                Value::String(text) => joined.push_str(text),
                                other => joined.push_str(&crate::keys::value_display_string(other)),
                            }
                        }
                    }
                    parse_to_number(&joined)
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
    number_const(
        vm,
        "MAX_SAFE_INTEGER",
        Box::new(|_ctx, _args| Value::F64(9007199254740991.0)),
    );
    number_const(
        vm,
        "MIN_SAFE_INTEGER",
        Box::new(|_ctx, _args| Value::F64(-9007199254740991.0)),
    );
    number_const(
        vm,
        "MAX_VALUE",
        Box::new(|_ctx, _args| Value::F64(f64::MAX)),
    );
    number_const(
        vm,
        "MIN_VALUE",
        Box::new(|_ctx, _args| Value::F64(f64::from_bits(1))),
    );
    number_const(
        vm,
        "EPSILON",
        Box::new(|_ctx, _args| Value::F64(f64::EPSILON)),
    );
    number_const(
        vm,
        "POSITIVE_INFINITY",
        Box::new(|_ctx, _args| Value::F64(f64::INFINITY)),
    );
    number_const(
        vm,
        "NEGATIVE_INFINITY",
        Box::new(|_ctx, _args| Value::F64(f64::NEG_INFINITY)),
    );
    number_const(vm, "NaN", Box::new(|_ctx, _args| Value::F64(f64::NAN)));
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
    number_fn_free(
        vm,
        "isFinite",
        vec![ValType::Any],
        vec![ValType::Bool],
        Box::new(|_ctx, args| match args.first() {
            Some(Value::F64(n)) => Value::Bool(n.is_finite()),
            Some(Value::I32(_)) => Value::Bool(true),
            _ => Value::Bool(false),
        }),
    );

    // ecma:number:isNaN — STRICT, no coercion (Number.isNaN).
    number_fn_free(
        vm,
        "isNaN",
        vec![ValType::Any],
        vec![ValType::Bool],
        Box::new(|_ctx, args| match args.first() {
            Some(Value::F64(n)) => Value::Bool(n.is_nan()),
            _ => Value::Bool(false),
        }),
    );

    // ecma:number:globalIsFinite — coerces (global isFinite).
    number_fn(
        vm,
        "globalIsFinite",
        vec![ValType::Any],
        vec![ValType::Bool],
        Box::new(|_ctx, args| {
            let n = args.first().map(to_f64_coerce).unwrap_or(f64::NAN);
            Value::Bool(n.is_finite())
        }),
    );

    // ecma:number:globalIsNaN — coerces (global isNaN).
    number_fn(
        vm,
        "globalIsNaN",
        vec![ValType::Any],
        vec![ValType::Bool],
        Box::new(|_ctx, args| {
            let n = args.first().map(to_f64_coerce).unwrap_or(f64::NAN);
            Value::Bool(n.is_nan())
        }),
    );

    number_fn(
        vm,
        "isInteger",
        vec![ValType::Any],
        vec![ValType::Bool],
        Box::new(|_ctx, args| match args.first() {
            Some(Value::F64(n)) => Value::Bool(n.is_finite() && n.fract() == 0.0),
            Some(Value::I32(_)) => Value::Bool(true),
            _ => Value::Bool(false),
        }),
    );

    number_fn(
        vm,
        "isSafeInteger",
        vec![ValType::Any],
        vec![ValType::Bool],
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
    vm.register_free_fn(
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

    number_fn_free(
        vm,
        "parseFloat",
        vec![ValType::String],
        vec![ValType::F64],
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
    // §19.2.5 step 8: a radix outside 2..=36 (0 = auto-detect handled
    // below) returns NaN — and must never reach char::to_digit, which
    // PANICS on out-of-range radices.
    if radix != 0 && !(2..=36).contains(&radix) {
        return f64::NAN;
    }
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

/// ECMA-262 §19.2.4 ParseFloat: skip leading whitespace, parse the
/// longest substring that looks like a float; return NaN if the
/// prefix isn't a valid number start.
fn parse_float_ecma(input: &str) -> f64 {
    let trimmed = input.trim_start_matches(|ch| {
        matches!(ch,
            '\t' | '\n' | '\x0b' | '\x0c' | '\r' | ' '
            | '\u{a0}' | '\u{1680}' | '\u{2000}'..='\u{200a}'
            | '\u{2028}' | '\u{2029}' | '\u{202f}' | '\u{205f}'
            | '\u{3000}' | '\u{feff}'
        )
    });
    let bytes = trimmed.as_bytes();
    let mut end = usize::from(matches!(bytes.first(), Some(b'+' | b'-')));
    if trimmed[end..].starts_with("Infinity") {
        return if bytes.first() == Some(&b'-') {
            f64::NEG_INFINITY
        } else {
            f64::INFINITY
        };
    }
    // Most calls contain a complete numeric string. Parse those directly,
    // without first walking the same digits in the prefix scanner. Restrict
    // this path to decimal starts so Rust's `inf`/`NaN` extensions cannot
    // bypass the ECMAScript grammar.
    if bytes.get(end).is_some_and(|byte| byte.is_ascii_digit() || *byte == b'.') {
        if let Ok(number) = trimmed.parse::<f64>() {
            return number;
        }
    }
    let integer_start = end;
    while bytes.get(end).is_some_and(|byte| byte.is_ascii_digit()) {
        end += 1;
    }
    let mut digits = end - integer_start;
    if bytes.get(end) == Some(&b'.') {
        end += 1;
        let fraction_start = end;
        while bytes.get(end).is_some_and(|byte| byte.is_ascii_digit()) {
            end += 1;
        }
        digits += end - fraction_start;
    }
    if digits == 0 {
        return f64::NAN;
    }
    if matches!(bytes.get(end), Some(b'e' | b'E')) {
        let exponent_start = end;
        end += 1;
        if matches!(bytes.get(end), Some(b'+' | b'-')) {
            end += 1;
        }
        let exponent_digits = end;
        while bytes.get(end).is_some_and(|byte| byte.is_ascii_digit()) {
            end += 1;
        }
        if end == exponent_digits {
            end = exponent_start;
        }
    }
    // Only ASCII grammar bytes were consumed, so this is a UTF-8 boundary.
    trimmed[..end].parse::<f64>().unwrap_or(f64::NAN)
}

// ── Number.prototype methods ──────────────────────────────────────
//
// Method receivers are the first arg per Component-Model `[method]`
// convention. JS callers route `(42).toFixed(2)` through Vybe's
// compiler as `ecma:number.toFixed(42, 2)`.

/// §6.1.6.1.20 Number::toString for non-finite values — "Infinity",
/// "-Infinity", "NaN" (Rust's Display prints "inf"/"NaN").
fn js_nonfinite_str(n: f64) -> String {
    if n.is_nan() {
        "NaN".to_string()
    } else if n > 0.0 {
        "Infinity".to_string()
    } else {
        "-Infinity".to_string()
    }
}

fn register_prototype(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:number",
        "toFixed",
        Box::new(|ctx, args| {
            let n = f_arg(args, 0).unwrap_or(0.0);
            let digits_f = match args.get(1) {
                Some(Value::F64(d)) => *d,
                Some(Value::I32(d)) => *d as f64,
                _ => 0.0,
            };
            // §21.1.3.3 step 2: RangeError unless 0 ≤ digits ≤ 100.
            if !(0.0..=100.0).contains(&digits_f) || digits_f.is_nan() {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "toFixed() digits argument must be between 0 and 100",
                ));
                return Value::Undefined;
            }
            let digits = digits_f as usize;
            // §21.1.3.3 step 6: non-finite → ToString ("Infinity", not
            // Rust's "inf"); -0 formats as "0" (§6.1.6.1.20).
            if !n.is_finite() {
                return s_val(&js_nonfinite_str(n));
            }
            let n = if n == 0.0 { 0.0 } else { n };
            s_owned(format!("{:.1$}", n, digits))
        }),
    );

    vm.register_host_fn(
        "ecma:number",
        "toString",
        Box::new(|ctx, args| {
            let n = f_arg(args, 0).unwrap_or(0.0);
            let radix = match args.get(1) {
                Some(Value::F64(r)) => *r as u32,
                Some(Value::I32(r)) => *r as u32,
                _ => 10,
            };
            // §21.1.3.6 step 2: RangeError unless 2 ≤ radix ≤ 36.
            if !(2..=36).contains(&radix) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "toString() radix must be between 2 and 36",
                ));
                return Value::Undefined;
            }
            if radix == 10 {
                // Integer-valued floats print without trailing ".0" per JS.
                if n.is_finite() && n.fract() == 0.0 {
                    return s_owned((n as i64).to_string());
                }
                return s_owned(n.to_string());
            }
            if !n.is_finite() {
                return s_owned(n.to_string());
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
            s_owned(out)
        }),
    );

    // Receiver only — `Number.prototype.valueOf()` takes no argument.
    number_fn(
        vm,
        "valueOf",
        vec![ValType::F64],
        vec![ValType::F64],
        Box::new(|_ctx, args| Value::F64(f_arg(args, 0).unwrap_or(0.0))),
    );
    vm.register_host_fn(
        "ecma:number",
        "toLocaleString",
        Box::new(|_ctx, args| {
            let n = match args.first() {
                Some(Value::F64(f)) => *f,
                Some(Value::I32(i)) => *i as f64,
                _ => return crate::keys::string_value("0"),
            };
            if !n.is_finite() {
                return s_owned(n.to_string());
            }
            // ECMA-402 §16.2 default formatting: grouped integer part,
            // up to 3 fraction digits.
            let rounded = (n * 1000.0).round() / 1000.0;
            let neg = rounded < 0.0;
            let abs = rounded.abs();
            let int_part = abs.trunc();
            let int_str = (int_part as u64).to_string();
            let mut grouped = String::new();
            for (i, c) in int_str.chars().enumerate() {
                if i > 0 && (int_str.len() - i) % 3 == 0 {
                    grouped.push(',');
                }
                grouped.push(c);
            }
            let frac = abs - int_part;
            if frac > 0.0 {
                let frac_str = format!("{:.3}", frac);
                let frac_str = frac_str[1..].trim_end_matches('0');
                if frac_str != "." {
                    grouped.push_str(frac_str);
                }
            }
            if neg {
                grouped.insert(0, '-');
            }
            s_owned(grouped)
        }),
    );

    vm.register_host_fn(
        "ecma:number",
        "toExponential",
        Box::new(|ctx, args| {
            let n = f_arg(args, 0).unwrap_or(0.0);
            // §21.1.3.2 step 8: RangeError unless 0 ≤ fractionDigits ≤ 100.
            let frac = match args.get(1) {
                Some(Value::F64(d)) => Some(*d),
                Some(Value::I32(d)) => Some(*d as f64),
                _ => None,
            };
            if let Some(d) = frac {
                if !(0.0..=100.0).contains(&d) || d.is_nan() {
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "RangeError",
                        "toExponential() argument must be between 0 and 100",
                    ));
                    return Value::Undefined;
                }
            }
            // §21.1.3.2 step 6: non-finite → ToString.
            if !n.is_finite() {
                return s_val(&js_nonfinite_str(n));
            }
            let raw = match args.get(1) {
                Some(Value::F64(d)) => format!("{:.1$e}", n, *d as usize),
                Some(Value::I32(d)) => format!("{:.1$e}", n, *d as usize),
                _ => format!("{:e}", n),
            };
            if let Some((mantissa, exponent)) = raw.split_once('e') {
                let exp: i32 = exponent.parse().unwrap_or(0);
                let sign = if exp >= 0 { "+" } else { "" };
                s_owned(format!("{}e{}{}", mantissa, sign, exp))
            } else {
                s_owned(raw)
            }
        }),
    );

    vm.register_host_fn(
        "ecma:number",
        "toPrecision",
        Box::new(|ctx, args| {
            let n = f_arg(args, 0).unwrap_or(0.0);
            let prec_f = match args.get(1) {
                Some(Value::F64(p)) => *p,
                Some(Value::I32(p)) => *p as f64,
                _ => return s_owned(n.to_string()),
            };
            // §21.1.3.5 step 8: RangeError unless 1 ≤ precision ≤ 100.
            if !(1.0..=100.0).contains(&prec_f) || prec_f.is_nan() {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "toPrecision() argument must be between 1 and 100",
                ));
                return Value::Undefined;
            }
            let prec = prec_f as usize;
            // §21.1.3.5 step 9: non-finite → ToString.
            if !n.is_finite() {
                return s_val(&js_nonfinite_str(n));
            }
            if n == 0.0 {
                if prec <= 1 {
                    return s_val("0");
                }
                return s_owned(format!("0.{:0<1$}", "", prec - 1));
            }
            // The exponent must come from the value AFTER rounding to
            // `prec` significant digits ((9.99).toPrecision(1) rounds to
            // 10 → e=1 ≥ 1 → "1e+1") — so derive it from the rounded
            // exponential form, not log10 of the raw value.
            let raw = format!("{:.1$e}", n, prec - 1);
            let (mantissa, exp_str) = raw.split_once('e').unwrap_or((raw.as_str(), "0"));
            let e: i32 = exp_str.parse().unwrap_or(0);
            if e >= prec as i32 || e < -6 {
                let sign = if e >= 0 { "+" } else { "" };
                return s_owned(format!("{}e{}{}", mantissa, sign, e));
            }
            let decimal_places = ((prec as i32 - 1 - e).max(0)) as usize;
            s_owned(format!("{:.1$}", n, decimal_places))
        }),
    );
}

#[cfg(test)]
mod numeric_parser_speedup_tests {
    use super::{parse_float_ecma, s_arg};
    use std::borrow::Cow;
    use vybe_runtime::Value;

    #[test]
    #[ignore = "explicit native microbenchmark; timing is not a conformance gate"]
    fn native_parse_float_microbenchmark() {
        use std::hint::black_box;
        use std::time::Instant;

        fn old_prefix_retry(input: &str) -> f64 {
            let trimmed = input.trim_start();
            let mut end = trimmed.len();
            while end > 0 {
                if let Ok(number) = trimmed[..end].parse::<f64>() {
                    return number;
                }
                end -= 1;
            }
            f64::NAN
        }

        // ASCII only: the old implementation cannot safely handle UTF-8 suffixes.
        // Include both a normal valid input and long trailing invalid text.
        let suffix_input = format!("12345678901234567890.125{}", "x".repeat(1024));
        for input in ["12345.125e-2", suffix_input.as_str()] {
            assert_eq!(old_prefix_retry(input).to_bits(), parse_float_ecma(input).to_bits());
            let iterations = 1000;
            let started = Instant::now();
            for _ in 0..iterations {
                black_box(old_prefix_retry(black_box(input)));
            }
            let retry = started.elapsed();
            let started = Instant::now();
            for _ in 0..iterations {
                black_box(parse_float_ecma(black_box(input)));
            }
            let scanned = started.elapsed();
            eprintln!(
                "native parseFloat, n={iterations}, input_bytes={}: prefix-retry={retry:?}, single-scan={scanned:?}",
                input.len()
            );
        }
    }

    #[test]
    fn decimal_prefixes_and_incomplete_exponents() {
        for (input, expected) in [
            ("12345.125e-2", 123.45125),
            ("12.5px", 12.5), ("+.5rest", 0.5), ("1.tail", 1.0),
            ("1e+2tail", 100.0), ("1e+tail", 1.0), ("1e", 1.0),
            ("0x10", 0.0), ("1_000", 1.0), ("12\u{1f600}", 12.0),
            ("\u{feff}\u{2028}-2.5", -2.5),
        ] {
            assert_eq!(parse_float_ecma(input), expected, "{input:?}");
        }
        assert_eq!(parse_float_ecma("-0tail").to_bits(), (-0.0f64).to_bits());
        assert_eq!(parse_float_ecma("Infinitytail"), f64::INFINITY);
        assert_eq!(parse_float_ecma("-Infinitytail"), f64::NEG_INFINITY);
        assert_eq!(parse_float_ecma("1e999"), f64::INFINITY);
        for input in ["", "+", ".", ".e1", "inf", "infinity", "NaN", "\u{85}1", "\u{1f600}12"] {
            assert!(parse_float_ecma(input).is_nan(), "{input:?}");
        }
    }

    #[test]
    fn long_invalid_suffix_and_borrowed_input() {
        let input = format!("3.25{}", "x".repeat(100_000));
        assert_eq!(parse_float_ecma(&input), 3.25);
        let args = [crate::keys::string_value(&input)];
        assert!(matches!(s_arg(&args, 0), Cow::Borrowed(_)));
        assert_eq!(s_arg(&[Value::I32(42)], 0), "42");
    }
}
