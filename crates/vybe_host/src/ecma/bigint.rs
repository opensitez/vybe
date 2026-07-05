//! ECMA-262 §21.2 — BigInt + arithmetic adapters.
//!
//!   §21.2.1   BigInt(value) — arbitrary-precision integer constructor
//!   §21.2.2.1 BigInt.asIntN(bits, bigint) — wrap to signed N-bit
//!   §21.2.2.2 BigInt.asUintN(bits, bigint) — wrap to unsigned N-bit
//!   §21.2.3   BigInt.prototype.{toString, valueOf, toLocaleString}
//!
//! Plus arithmetic adapters that mirror `i64.*` WASM opcodes — when the
//! compiler statically knows both operands are BigInt it emits the
//! opcode directly; this module is the dynamic-dispatch fallback.
//!
//! Vybe's `Value::BigInt(i64)` is bounded to i64 range. Truly arbitrary
//! precision (`BigInt("9999999999999999999999")`) saturates / wraps —
//! the spec-conformant 256+ bit path is a future upgrade.

use vybe_bytecode::{HostContext, VM, Value};

/// §7.1.14 StringToBigInt: optional sign for decimal, 0x/0o/0b radix
/// prefixes (no sign), empty/whitespace → 0. None = invalid → the caller
/// decides (BigInt() throws SyntaxError; coercions treat as 0).
fn parse_bigint_str(s: &str) -> Option<i64> {
    let t = s.trim();
    if t.is_empty() {
        return Some(0);
    }
    let radix = |p: &str, r: u32| i64::from_str_radix(p, r).ok();
    if let Some(hex) = t.strip_prefix("0x").or_else(|| t.strip_prefix("0X")) {
        return radix(hex, 16);
    }
    if let Some(bin) = t.strip_prefix("0b").or_else(|| t.strip_prefix("0B")) {
        return radix(bin, 2);
    }
    if let Some(oct) = t.strip_prefix("0o").or_else(|| t.strip_prefix("0O")) {
        return radix(oct, 8);
    }
    t.parse::<i64>().ok()
}

fn to_bigint(v: &Value) -> i64 {
    match v {
        Value::BigInt(n) => *n,
        Value::I32(n) => *n as i64,
        Value::I64(n) => *n,
        Value::F64(n) => *n as i64,
        Value::Bool(b) => {
            if *b {
                1
            } else {
                0
            }
        }
        Value::String(s) => parse_bigint_str(s).unwrap_or(0),
        _ => 0,
    }
}

pub fn register(vm: &mut VM) {
    // BigInt(value) — §21.2.1.1.
    vm.register_host_fn(
        "ecma:bigint",
        "BigInt",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // §21.2.1.1: return Value::BigInt so strict equality (===) against
            // BigInt literals works — I64 and BigInt are different types.
            // §7.1.14: an unparsable string is a SyntaxError.
            if let Some(Value::String(text)) = args.first() {
                return match parse_bigint_str(text) {
                    Some(n) => Value::BigInt(n),
                    None => {
                        ctx.throw_value(crate::ecma::error::new_error(
                            "SyntaxError",
                            &format!("Cannot convert {} to a BigInt", text),
                        ));
                        Value::Undefined
                    }
                };
            }
            Value::BigInt(to_bigint(args.first().unwrap_or(&Value::Null)))
        }),
    );

    // BigInt.asIntN(bits, bigint) — §21.2.2.1. Sign-extend low `bits`.
    vm.register_host_fn(
        "ecma:bigint",
        "asIntN",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let bits = args.first().map(|v| v.as_f64() as u32).unwrap_or(0).min(64);
            let n = to_bigint(args.get(1).unwrap_or(&Value::Null));
            if bits == 0 {
                return Value::BigInt(0);
            }
            if bits >= 64 {
                return Value::BigInt(n);
            }
            let mask = (1i64 << bits) - 1;
            let truncated = n & mask;
            let sign_bit = 1i64 << (bits - 1);
            let result = if truncated & sign_bit != 0 {
                truncated | !mask
            } else {
                truncated
            };
            Value::BigInt(result)
        }),
    );

    // BigInt.asUintN(bits, bigint) — §21.2.2.2. Truncate low `bits`.
    vm.register_host_fn(
        "ecma:bigint",
        "asUintN",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let bits = args.first().map(|v| v.as_f64() as u32).unwrap_or(0).min(64);
            let n = to_bigint(args.get(1).unwrap_or(&Value::Null));
            if bits == 0 {
                return Value::BigInt(0);
            }
            if bits >= 64 {
                return Value::BigInt(n);
            }
            let mask = (1i64 << bits) - 1;
            Value::BigInt(n & mask)
        }),
    );

    // BigInt.prototype.toLocaleString — §21.2.3.3.
    vm.register_host_fn(
        "ecma:bigint",
        "toLocaleString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = to_bigint(args.first().unwrap_or(&Value::Null));
            Value::String(std::sync::Arc::from(n.to_string().as_str()))
        }),
    );

    // BigInt.prototype.toString(radix?) — §21.2.3.4.
    vm.register_host_fn(
        "ecma:bigint",
        "toString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = to_bigint(args.first().unwrap_or(&Value::Null));
            let radix = args.get(1).map(|v| v.as_f64() as u32).unwrap_or(10);
            if !(2..=36).contains(&radix) {
                return Value::String(std::sync::Arc::from(n.to_string().as_str()));
            }
            let s = if radix == 10 {
                n.to_string()
            } else {
                // Manual radix conversion preserving sign.
                let (neg, mut abs_n) = if n < 0 {
                    (true, n.unsigned_abs())
                } else {
                    (false, n as u64)
                };
                if abs_n == 0 {
                    return Value::String(std::sync::Arc::from("0"));
                }
                let mut digits = Vec::new();
                while abs_n > 0 {
                    let d = (abs_n % radix as u64) as u32;
                    digits.push(std::char::from_digit(d, radix).unwrap_or('?'));
                    abs_n /= radix as u64;
                }
                digits.reverse();
                let mut s: String = digits.into_iter().collect();
                if neg {
                    s.insert(0, '-');
                }
                s
            };
            Value::String(std::sync::Arc::from(s.as_str()))
        }),
    );

    // Alias: toStringRadix(bigint, radix) — same as toString(radix).
    vm.register_host_fn(
        "ecma:bigint",
        "toStringRadix",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = to_bigint(args.first().unwrap_or(&Value::Null));
            let radix = args.get(1).map(|v| v.as_f64() as u32).unwrap_or(10);
            let (neg, mut abs_n) = if n < 0 {
                (true, n.unsigned_abs())
            } else {
                (false, n as u64)
            };
            if abs_n == 0 {
                return Value::String(std::sync::Arc::from("0"));
            }
            let mut digits = Vec::new();
            while abs_n > 0 {
                let d = (abs_n % radix as u64) as u32;
                digits.push(std::char::from_digit(d, radix).unwrap_or('?'));
                abs_n /= radix as u64;
            }
            digits.reverse();
            let mut s: String = digits.into_iter().collect();
            if neg {
                s.insert(0, '-');
            }
            Value::String(std::sync::Arc::from(s.as_str()))
        }),
    );

    // BigInt.prototype.valueOf — §21.2.3.5. Returns the primitive BigInt.
    vm.register_host_fn(
        "ecma:bigint",
        "valueOf",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::BigInt(to_bigint(args.first().unwrap_or(&Value::Null)))
        }),
    );

    // ── Arithmetic adapters mirroring i64.* WASM opcodes ────────────
    //
    // When the compiler statically knows both operands are BigInt it
    // emits I64_ADD/SUB/etc. directly; these handlers cover the
    // dynamic-dispatch case (e.g. `Reflect.apply(BigInt.add, ...)`).
    // Semantics use Rust's wrapping_* to match WASM i64.* trap-free
    // arithmetic. div/rem/pow are registered separately below: the spec
    // requires them to throw RangeError (§6.1.6.2.3/5/6) rather than
    // produce a value.

    macro_rules! binop {
        ($name:expr, $op:expr) => {
            vm.register_host_fn(
                "ecma:bigint",
                $name,
                Box::new(|_ctx: &mut HostContext, args: &[Value]| {
                    let a = to_bigint(args.first().unwrap_or(&Value::Null));
                    let b = to_bigint(args.get(1).unwrap_or(&Value::Null));
                    Value::BigInt($op(a, b))
                }),
            );
        };
    }
    binop!("add", |a: i64, b: i64| a.wrapping_add(b));
    binop!("sub", |a: i64, b: i64| a.wrapping_sub(b));
    binop!("mul", |a: i64, b: i64| a.wrapping_mul(b));
    binop!("and", |a: i64, b: i64| a & b);
    binop!("or", |a: i64, b: i64| a | b);
    binop!("xor", |a: i64, b: i64| a ^ b);
    binop!("shl", |a: i64, b: i64| a.wrapping_shl((b & 63) as u32));
    binop!("shr", |a: i64, b: i64| a.wrapping_shr((b & 63) as u32));

    // §6.1.6.2.5 BigInt::divide / §6.1.6.2.6 BigInt::remainder — a zero
    // divisor throws RangeError (unlike Number's Infinity/NaN).
    macro_rules! divlike {
        ($name:expr, $op:expr) => {
            vm.register_host_fn(
                "ecma:bigint",
                $name,
                Box::new(|ctx: &mut HostContext, args: &[Value]| {
                    let a = to_bigint(args.first().unwrap_or(&Value::Null));
                    let b = to_bigint(args.get(1).unwrap_or(&Value::Null));
                    if b == 0 {
                        ctx.throw_value(crate::ecma::error::new_error(
                            "RangeError",
                            "Division by zero",
                        ));
                        return Value::Undefined;
                    }
                    Value::BigInt($op(a, b))
                }),
            );
        };
    }
    divlike!("div", |a: i64, b: i64| a.wrapping_div(b));
    divlike!("rem", |a: i64, b: i64| a.wrapping_rem(b));

    // §6.1.6.2.3 BigInt::exponentiate — negative exponent throws RangeError.
    vm.register_host_fn(
        "ecma:bigint",
        "pow",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let a = to_bigint(args.first().unwrap_or(&Value::Null));
            let b = to_bigint(args.get(1).unwrap_or(&Value::Null));
            if b < 0 {
                ctx.throw_value(crate::ecma::error::new_error(
                    "RangeError",
                    "Exponent must be non-negative",
                ));
                return Value::Undefined;
            }
            Value::BigInt(if b == 0 { 1 } else { a.wrapping_pow(b as u32) })
        }),
    );

    vm.register_host_fn(
        "ecma:bigint",
        "neg",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::BigInt(to_bigint(args.first().unwrap_or(&Value::Null)).wrapping_neg())
        }),
    );
    vm.register_host_fn(
        "ecma:bigint",
        "not",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::BigInt(!to_bigint(args.first().unwrap_or(&Value::Null)))
        }),
    );

    // Comparison adapters.
    macro_rules! cmpop {
        ($name:expr, $op:expr) => {
            vm.register_host_fn(
                "ecma:bigint",
                $name,
                Box::new(|_ctx: &mut HostContext, args: &[Value]| {
                    let a = to_bigint(args.first().unwrap_or(&Value::Null));
                    let b = to_bigint(args.get(1).unwrap_or(&Value::Null));
                    Value::Bool($op(a, b))
                }),
            );
        };
    }
    cmpop!("eq", |a: i64, b: i64| a == b);
    cmpop!("ne", |a: i64, b: i64| a != b);
    cmpop!("lt", |a: i64, b: i64| a < b);
    cmpop!("le", |a: i64, b: i64| a <= b);
    cmpop!("gt", |a: i64, b: i64| a > b);
    cmpop!("ge", |a: i64, b: i64| a >= b);
}
