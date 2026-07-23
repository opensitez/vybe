//! ECMA-262 §21.2 — BigInt + arithmetic adapters.
//!
//!   §21.2.1   BigInt(value) — arbitrary-precision integer constructor
//!   §21.2.2.1 BigInt.asIntN(bits, bigint) — wrap to signed N-bit
//!   §21.2.2.2 BigInt.asUintN(bits, bigint) — wrap to unsigned N-bit
//!   §21.2.3   BigInt.prototype.{toString, valueOf, toLocaleString}
//!
//! Plus the §6.1.6.2 operation adapters the compiler's dynamic dispatch
//! calls (`ecma:bigint.add` etc.). All EXACT at arbitrary width —
//! `Value::BigInt` is backed by `vybe_bytecode::bigint::BigIntVal`
//! (sign + limbs). The only wraps anywhere are the js-types JS-API
//! ToBigInt64/ToBigUint64 conversions at wasm boundaries, and the
//! spec's own asIntN/asUintN operators.
//!
//! §6.1.6.2 permits an implementation-defined size limit surfaced as
//! RangeError ("Maximum BigInt size exceeded") — enforced for the
//! explosive operators (pow, shl) via `BigIntVal::exceeds_cap`.

use std::sync::Arc;
use vybe_bytecode::bigint::BigIntVal;
use vybe_bytecode::{HostContext, VM, Value};

/// §7.1.14 StringToBigInt — exact at any length. None = invalid → the
/// caller decides (BigInt() throws SyntaxError; coercions treat as 0).
pub(crate) fn parse_bigint_str(s: &str) -> Option<BigIntVal> {
    BigIntVal::parse(s)
}

/// ToBigInt-style coercion for adapter arguments (§7.1.13 shape; the
/// compiler already guards types at the call sites that must throw).
fn to_bigint(v: &Value) -> BigIntVal {
    match v {
        Value::BigInt(n) => (**n).clone(),
        Value::I32(n) => BigIntVal::from_i64(*n as i64),
        Value::I64(n) => BigIntVal::from_i64(*n),
        Value::F64(n) => BigIntVal::from_f64(*n),
        Value::Bool(b) => BigIntVal::from_i64(*b as i64),
        Value::String(s) => BigIntVal::parse(s).unwrap_or_else(BigIntVal::zero),
        _ => BigIntVal::zero(),
    }
}

fn big(v: BigIntVal) -> Value {
    Value::BigInt(Arc::new(v))
}

fn throw_range(ctx: &mut HostContext, msg: &str) -> Value {
    ctx.throw_value(crate::ecma::error::new_error(ctx, "RangeError", msg));
    Value::Undefined
}

pub fn register(vm: &mut VM) {
    // BigInt(value) — §21.2.1.1.
    vm.register_host_fn(
        "ecma:bigint",
        "BigInt",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // §7.1.14: an unparsable string is a SyntaxError.
            if let Some(Value::String(text)) = args.first() {
                return match BigIntVal::parse(text) {
                    Some(n) => big(n),
                    None => {
                        ctx.throw_value(crate::ecma::error::new_error(
                            ctx,
                            "SyntaxError",
                            &format!("Cannot convert {} to a BigInt", text),
                        ));
                        Value::Undefined
                    }
                };
            }
            big(to_bigint(args.first().unwrap_or(&Value::Null)))
        }),
    );

    // BigInt.asIntN(bits, bigint) — §21.2.2.1.
    vm.register_host_fn(
        "ecma:bigint",
        "asIntN",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let bits = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
            let n = to_bigint(args.get(1).unwrap_or(&Value::Null));
            big(n.as_int_n(bits))
        }),
    );

    // BigInt.asUintN(bits, bigint) — §21.2.2.2.
    vm.register_host_fn(
        "ecma:bigint",
        "asUintN",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let bits = args.first().map(|v| v.as_f64() as u64).unwrap_or(0);
            let n = to_bigint(args.get(1).unwrap_or(&Value::Null));
            big(n.as_uint_n(bits))
        }),
    );

    // BigInt.prototype.toLocaleString — §21.2.3.3.
    vm.register_host_fn(
        "ecma:bigint",
        "toLocaleString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = to_bigint(args.first().unwrap_or(&Value::Null));
            Value::String(Arc::from(n.to_string().as_str()))
        }),
    );

    // BigInt.prototype.toString(radix?) — §21.2.3.4 / §6.1.6.2.23.
    vm.register_host_fn(
        "ecma:bigint",
        "toString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = to_bigint(args.first().unwrap_or(&Value::Null));
            let radix = args.get(1).map(|v| v.as_f64() as u32).unwrap_or(10);
            let radix = if (2..=36).contains(&radix) { radix } else { 10 };
            Value::String(Arc::from(n.to_string_radix(radix).as_str()))
        }),
    );

    // Alias: toStringRadix(bigint, radix) — same as toString(radix).
    vm.register_host_fn(
        "ecma:bigint",
        "toStringRadix",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = to_bigint(args.first().unwrap_or(&Value::Null));
            let radix = args.get(1).map(|v| v.as_f64() as u32).unwrap_or(10);
            let radix = if (2..=36).contains(&radix) { radix } else { 10 };
            Value::String(Arc::from(n.to_string_radix(radix).as_str()))
        }),
    );

    // BigInt.prototype.valueOf — §21.2.3.5. Returns the primitive BigInt.
    vm.register_host_fn(
        "ecma:bigint",
        "valueOf",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            big(to_bigint(args.first().unwrap_or(&Value::Null)))
        }),
    );

    // ── §6.1.6.2 operation adapters — EXACT arbitrary precision ──────
    macro_rules! binop {
        ($name:expr, $op:expr) => {
            vm.register_host_fn(
                "ecma:bigint",
                $name,
                Box::new(|_ctx: &mut HostContext, args: &[Value]| {
                    let a = to_bigint(args.first().unwrap_or(&Value::Null));
                    let b = to_bigint(args.get(1).unwrap_or(&Value::Null));
                    let f: fn(&BigIntVal, &BigIntVal) -> BigIntVal = $op;
                    big(f(&a, &b))
                }),
            );
        };
    }
    binop!("add", BigIntVal::add);
    binop!("sub", BigIntVal::sub);
    binop!("mul", BigIntVal::mul);
    binop!("and", BigIntVal::bit_and);
    binop!("or", BigIntVal::bit_or);
    binop!("xor", BigIntVal::bit_xor);

    // §6.1.6.2.9/10 shifts — arbitrary shift counts; oversize left
    // shifts hit the implementation-defined RangeError cap.
    vm.register_host_fn(
        "ecma:bigint",
        "shl",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let a = to_bigint(args.first().unwrap_or(&Value::Null));
            let b = to_bigint(args.get(1).unwrap_or(&Value::Null));
            let (a, b, left) = if b.is_negative() {
                (a, b.neg(), false)
            } else {
                (a, b, true)
            };
            if !b.fits_i64() {
                return if left {
                    throw_range(ctx, "Maximum BigInt size exceeded")
                } else {
                    // Shifting everything out: 0 or -1 by sign.
                    big(BigIntVal::from_i64(if a.is_negative() { -1 } else { 0 }))
                };
            }
            let bits = b.to_i64_wrapping() as u64;
            if left {
                if BigIntVal::exceeds_cap(a.bit_len() + bits as usize) {
                    return throw_range(ctx, "Maximum BigInt size exceeded");
                }
                big(a.shl(bits))
            } else {
                big(a.shr(bits))
            }
        }),
    );
    vm.register_host_fn(
        "ecma:bigint",
        "shr",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let a = to_bigint(args.first().unwrap_or(&Value::Null));
            let b = to_bigint(args.get(1).unwrap_or(&Value::Null));
            let (b, right) = if b.is_negative() {
                (b.neg(), false)
            } else {
                (b, true)
            };
            if !b.fits_i64() {
                return if right {
                    big(BigIntVal::from_i64(if a.is_negative() { -1 } else { 0 }))
                } else {
                    throw_range(ctx, "Maximum BigInt size exceeded")
                };
            }
            let bits = b.to_i64_wrapping() as u64;
            if right {
                big(a.shr(bits))
            } else {
                if BigIntVal::exceeds_cap(a.bit_len() + bits as usize) {
                    return throw_range(ctx, "Maximum BigInt size exceeded");
                }
                big(a.shl(bits))
            }
        }),
    );

    // §6.1.6.2.5 BigInt::divide / §6.1.6.2.6 BigInt::remainder — a zero
    // divisor throws RangeError (unlike Number's Infinity/NaN).
    macro_rules! divlike {
        ($name:expr, $pick:expr) => {
            vm.register_host_fn(
                "ecma:bigint",
                $name,
                Box::new(|ctx: &mut HostContext, args: &[Value]| {
                    let a = to_bigint(args.first().unwrap_or(&Value::Null));
                    let b = to_bigint(args.get(1).unwrap_or(&Value::Null));
                    if b.is_zero() {
                        return throw_range(ctx, "Division by zero");
                    }
                    let (q, r) = a.divrem(&b);
                    let pick: fn(BigIntVal, BigIntVal) -> BigIntVal = $pick;
                    big(pick(q, r))
                }),
            );
        };
    }
    divlike!("div", |q, _r| q);
    divlike!("rem", |_q, r| r);

    // §6.1.6.2.3 BigInt::exponentiate — negative exponent throws
    // RangeError; oversize results hit the size cap (also RangeError).
    vm.register_host_fn(
        "ecma:bigint",
        "pow",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let a = to_bigint(args.first().unwrap_or(&Value::Null));
            let b = to_bigint(args.get(1).unwrap_or(&Value::Null));
            if b.is_negative() {
                return throw_range(ctx, "Exponent must be non-negative");
            }
            if b.is_zero() {
                return big(BigIntVal::from_i64(1));
            }
            if !b.fits_i64() {
                return throw_range(ctx, "Maximum BigInt size exceeded");
            }
            let exp = b.to_i64_wrapping() as u64;
            // Result needs ~bit_len(a) * exp bits — pre-check the cap.
            if BigIntVal::exceeds_cap(a.bit_len().saturating_mul(exp as usize)) {
                return throw_range(ctx, "Maximum BigInt size exceeded");
            }
            big(a.pow(exp))
        }),
    );

    vm.register_host_fn(
        "ecma:bigint",
        "neg",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            big(to_bigint(args.first().unwrap_or(&Value::Null)).neg())
        }),
    );
    vm.register_host_fn(
        "ecma:bigint",
        "not",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            big(to_bigint(args.first().unwrap_or(&Value::Null)).not())
        }),
    );

    // Comparison adapters — §7.2.13: BigInt/Number comparisons are on
    // MATHEMATICAL values (a Number operand is not truncated).
    macro_rules! cmpop {
        ($name:expr, $accept:expr, $nan:expr) => {
            vm.register_host_fn(
                "ecma:bigint",
                $name,
                Box::new(|_ctx: &mut HostContext, args: &[Value]| {
                    use std::cmp::Ordering;
                    let a = args.first().cloned().unwrap_or(Value::Null);
                    let b = args.get(1).cloned().unwrap_or(Value::Null);
                    // None = a NaN operand: relational ops are false,
                    // != is true (§7.2.13 / IsLessThan undefined).
                    let ord: Option<Ordering> = match (&a, &b) {
                        (Value::BigInt(x), Value::F64(f)) => x.cmp_f64(*f),
                        (Value::F64(f), Value::BigInt(y)) => y.cmp_f64(*f).map(Ordering::reverse),
                        _ => Some(to_bigint(&a).cmp_big(&to_bigint(&b))),
                    };
                    let accept: fn(Ordering) -> bool = $accept;
                    Value::Bool(ord.map(accept).unwrap_or($nan))
                }),
            );
        };
    }
    cmpop!("eq", |o| o == std::cmp::Ordering::Equal, false);
    cmpop!("ne", |o| o != std::cmp::Ordering::Equal, true);
    cmpop!("lt", |o| o == std::cmp::Ordering::Less, false);
    cmpop!("le", |o| o != std::cmp::Ordering::Greater, false);
    cmpop!("gt", |o| o == std::cmp::Ordering::Greater, false);
    cmpop!("ge", |o| o != std::cmp::Ordering::Less, false);
}
