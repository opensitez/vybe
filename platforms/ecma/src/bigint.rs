//! ECMA-262 §21.2 — BigInt + arithmetic adapters.
//!
//!   §21.2.1   BigInt(value) — arbitrary-precision integer constructor
//!   §21.2.2.1 BigInt.asIntN(bits, bigint) — wrap to signed N-bit
//!   §21.2.2.2 BigInt.asUintN(bits, bigint) — wrap to unsigned N-bit
//!   §21.2.3   BigInt.prototype.{toString, valueOf, toLocaleString}
//!
//! Plus the §6.1.6.2 operation adapters the compiler's dynamic dispatch
//! calls (`ecma:bigint.add` etc.). All EXACT at arbitrary width —
//! `Value::BigInt` is backed by `vybe_runtime::bigint::BigIntVal`
//! (sign + limbs). The only wraps anywhere are the js-types JS-API
//! ToBigInt64/ToBigUint64 conversions at wasm boundaries, and the
//! spec's own asIntN/asUintN operators.
//!
//! §6.1.6.2 permits an implementation-defined size limit surfaced as
//! RangeError ("Maximum BigInt size exceeded") — enforced for the
//! explosive operators (pow, shl) via `BigIntVal::exceeds_cap`.

use std::borrow::Cow;
use std::cmp::Ordering;
use std::sync::Arc;
use vybe_runtime::bigint::BigIntVal;
use vybe_runtime::{HostContext, VM, Value};

#[inline]
fn owned_string_value(text: String) -> Value {
    crate::keys::owned_string_value(text)
}

/// §7.1.14 StringToBigInt — exact at any length. None = invalid → the
/// caller decides (BigInt() throws SyntaxError; coercions treat as 0).
pub fn parse_bigint_str(s: &str) -> Option<BigIntVal> {
    BigIntVal::parse(s)
}

/// ToBigInt-style coercion for adapter arguments (§7.1.13 shape; the
/// compiler already guards types at the call sites that must throw).
fn to_bigint(v: &Value) -> Cow<'_, BigIntVal> {
    match v {
        Value::BigInt(n) => Cow::Borrowed(n.as_ref()),
        Value::I32(n) => Cow::Owned(BigIntVal::from_i64(*n as i64)),
        Value::I64(n) => Cow::Owned(BigIntVal::from_i64(*n)),
        Value::F64(n) => Cow::Owned(BigIntVal::from_f64(*n)),
        Value::Bool(b) => Cow::Owned(BigIntVal::from_i64(*b as i64)),
        Value::String(s) => Cow::Owned(BigIntVal::parse(s).unwrap_or_else(BigIntVal::zero)),
        _ => Cow::Owned(BigIntVal::zero()),
    }
}

fn to_bigint_value(v: &Value) -> Value {
    match v {
        Value::BigInt(n) => Value::BigInt(n.clone()),
        _ => big(to_bigint(v).into_owned()),
    }
}

fn bigint_comparison(a: &Value, b: &Value) -> Option<Ordering> {
    match (a, b) {
        (Value::BigInt(x), Value::BigInt(y)) => {
            Some(if Arc::ptr_eq(x, y) { Ordering::Equal } else { x.cmp_big(y) })
        }
        (Value::BigInt(x), Value::F64(f)) => x.cmp_f64(*f),
        (Value::F64(f), Value::BigInt(y)) => y.cmp_f64(*f).map(Ordering::reverse),
        _ => Some(to_bigint(a).cmp_big(&to_bigint(b))),
    }
}

fn big(v: BigIntVal) -> Value {
    Value::BigInt(Arc::new(v))
}

fn throw_range(ctx: &mut HostContext, msg: &str) -> Value {
    ctx.throw_value(crate::error::new_error(ctx, "RangeError", msg));
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
                        ctx.throw_value(crate::error::new_error(
                            ctx,
                            "SyntaxError",
                            &format!("Cannot convert {} to a BigInt", text),
                        ));
                        Value::Undefined
                    }
                };
            }
            to_bigint_value(args.first().unwrap_or(&Value::Null))
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
            owned_string_value(n.to_string())
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
            owned_string_value(n.to_string_radix(radix))
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
            owned_string_value(n.to_string_radix(radix))
        }),
    );

    // BigInt.prototype.valueOf — §21.2.3.5. Returns the primitive BigInt.
    vm.register_host_fn(
        "ecma:bigint",
        "valueOf",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            to_bigint_value(args.first().unwrap_or(&Value::Null))
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
                (a, Cow::Owned(b.neg()), false)
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
                (Cow::Owned(b.neg()), false)
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
                    let a = args.first().unwrap_or(&Value::Null);
                    let b = args.get(1).unwrap_or(&Value::Null);
                    // None = a NaN operand: relational ops are false,
                    // != is true (§7.2.13 / IsLessThan undefined).
                    let ord = bigint_comparison(a, b);
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

#[cfg(test)]
mod operand_borrow_speedup_tests {
    use super::*;
    use std::hint::black_box;
    use std::time::{Duration, Instant};

    #[test]
    fn immutable_bigint_operands_and_passthrough_values_share_storage() {
        let value = big(BigIntVal::from_i64(1).shl(65536));
        let Value::BigInt(storage) = &value else { panic!("expected BigInt") };
        let operand = to_bigint(&value);
        assert!(matches!(&operand, Cow::Borrowed(_)));
        assert!(std::ptr::eq(operand.as_ref(), storage.as_ref()));
        let Value::BigInt(passthrough) = to_bigint_value(&value) else { panic!("expected BigInt") };
        assert!(Arc::ptr_eq(storage, &passthrough));
        let plus_one = operand.add(&BigIntVal::from_i64(1));
        assert_eq!(plus_one.sub(operand.as_ref()).to_string(), "1");
        assert_eq!(storage.bit_len(), 65537);
    }

    #[test]
    fn owned_conversion_and_arithmetic_results_remain_exact() {
        for (value, expected) in [
            (Value::I32(-7), "-7"), (Value::I64(i64::MAX), "9223372036854775807"),
            (Value::F64(3.0), "3"), (Value::Bool(true), "1"),
            (crate::keys::string_value("18446744073709551617"), "18446744073709551617"),
            (Value::Null, "0"), (Value::Undefined, "0"),
        ] {
            let operand = to_bigint(&value);
            assert!(matches!(&operand, Cow::Owned(_)));
            assert_eq!(operand.to_string(), expected);
        }
        let a = big(BigIntVal::parse("18446744073709551617").unwrap());
        let b = big(BigIntVal::from_i64(-3));
        let a = to_bigint(&a);
        let b = to_bigint(&b);
        assert_eq!(a.add(&b).to_string(), "18446744073709551614");
        assert_eq!(a.mul(&b).to_string(), "-55340232221128654851");
        let (q, r) = a.divrem(&BigIntVal::from_i64(3));
        assert_eq!(q.to_string(), "6148914691236517205");
        assert_eq!(r.to_string(), "2");
        assert_eq!(a.as_uint_n(8).to_string(), "1");
        assert_eq!(b.as_int_n(8).to_string(), "-3");
        assert_eq!(b.neg().to_string(), "3");
    }

    #[test]
    fn comparisons_preserve_fractional_numbers_nan_infinity_and_identity() {
        let three = big(BigIntVal::from_i64(3));
        assert_eq!(bigint_comparison(&three, &three), Some(Ordering::Equal));
        assert_eq!(bigint_comparison(&three, &big(BigIntVal::from_i64(3))), Some(Ordering::Equal));
        assert_eq!(bigint_comparison(&three, &Value::F64(3.5)), Some(Ordering::Less));
        assert_eq!(bigint_comparison(&Value::F64(3.5), &three), Some(Ordering::Greater));
        assert_eq!(bigint_comparison(&three, &Value::F64(f64::NAN)), None);
        assert_eq!(bigint_comparison(&Value::F64(f64::NAN), &three), None);
        assert_eq!(bigint_comparison(&three, &Value::F64(f64::INFINITY)), Some(Ordering::Less));
        assert_eq!(bigint_comparison(&three, &Value::F64(f64::NEG_INFINITY)), Some(Ordering::Greater));
        let exact = big(BigIntVal::parse("9007199254740993").unwrap());
        assert_eq!(bigint_comparison(&exact, &Value::F64(9007199254740992.0)), Some(Ordering::Greater));
    }

    fn original_comparison(a: &Value, b: &Value) -> Option<Ordering> {
        let a = a.clone();
        let b = b.clone();
        match (&a, &b) {
            (Value::BigInt(x), Value::F64(f)) => x.cmp_f64(*f),
            (Value::F64(f), Value::BigInt(y)) => y.cmp_f64(*f).map(Ordering::reverse),
            _ => Some(to_bigint(&a).into_owned().cmp_big(&to_bigint(&b).into_owned())),
        }
    }

    #[test]
    #[ignore = "native operand/comparison microbenchmark, not complete BigInt dispatch timing"]
    fn native_borrowed_comparison_microbenchmark() {
        let large = BigIntVal::from_i64(1).shl(65536);
        let separate = big(large.clone());
        let large = big(large);
        let small = big(BigIntVal::from_i64(1));
        for (name, other) in [("different-width", &small), ("same-handle", &large), ("equal-distinct", &separate)] {
            assert_eq!(original_comparison(&large, other), bigint_comparison(&large, other));
            let mut old = Duration::ZERO;
            let mut new = Duration::ZERO;
            for batch in 0..10 {
                for old_path in if batch % 2 == 0 { [true, false] } else { [false, true] } {
                    let start = Instant::now();
                    for _ in 0..1000 {
                        black_box(if old_path { original_comparison(black_box(&large), black_box(other)) }
                            else { bigint_comparison(black_box(&large), black_box(other)) });
                    }
                    if old_path { old += start.elapsed(); } else { new += start.elapsed(); }
                }
            }
            eprintln!("native BigInt comparison, bits=65537, calls=10000, case={name}: cloned={old:?}, borrowed={new:?}");
        }
    }
}
