use std::sync::Arc;

use vybe_runtime::vm::HostFnDecl;
use vybe_runtime::{FuncSig, HostContext, VM, ValType, Value};

/// Declare an `ecma:math` function — same closure, plus the signature.
///
/// §21.3 is the one ECMA namespace that is almost entirely fixed-arity: the
/// operations take one or two numbers and answer one. The exceptions are the
/// variadic four — `max`, `min`, `hypot`, and the `maxOf`/`minOf`/`sumPrecise`
/// trio built on `collect_nums` — plus `log`, which reads an optional base.
/// Those stay undeclared, since a Component Model signature cannot say "rest".
fn math_fn(
    vm: &mut VM,
    name: &str,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(HostFnDecl::new("ecma:math", name, call).with_sig(FuncSig {
        name: name.to_string(),
        params,
        results,
    }));
}

/// A double. Every §21.3 operand and result is one — the spec's `Number`.
fn num() -> ValType {
    ValType::F64
}

/// A one-number operation: `sqrt`, `floor`, `sin`, …
fn unary(
    vm: &mut VM,
    name: &str,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    math_fn(vm, name, vec![num()], vec![num()], call);
}

/// A two-number operation: `pow`, `atan2`, `imul`.
fn binary(
    vm: &mut VM,
    name: &str,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    math_fn(vm, name, vec![num(), num()], vec![num()], call);
}

/// A spec CONSTANT (`Math.PI`). Registered as a host fn taking nothing —
/// `func() -> f64` is the honest signature, and it is why a caller that pushes
/// an operand is worth hearing about.
fn constant(
    vm: &mut VM,
    name: &str,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    math_fn(vm, name, vec![], vec![num()], call);
}

pub fn register(vm: &mut VM) {
    // Core math — callable ops stay as host functions, while spec
    // constants are registered as immutable value exports.
    unary(
        vm,
        "floor",
        Box::new(|_ctx, a| Value::F64(f(a, 0).floor())),
    );
    unary(
        vm,
        "ceil",
        Box::new(|_ctx, a| Value::F64(f(a, 0).ceil())),
    );
    unary(
        vm,
        "abs",
        Box::new(|_ctx, a| Value::F64(f(a, 0).abs())),
    );
    unary(
        vm,
        "sqrt",
        Box::new(|_ctx, a| Value::F64(f(a, 0).sqrt())),
    );
    unary(
        vm,
        "trunc",
        Box::new(|ctx, a| {
            let n = match number_arg(ctx, a, 0) {
                Ok(n) => n,
                Err(e) => return e,
            };
            let truncated = n.trunc();
            if truncated == 0.0 && truncated.is_sign_negative() {
                Value::String(Arc::from("-0"))
            } else {
                Value::F64(truncated)
            }
        }),
    );
    // ECMA-262 §21.3.2.28: Math.round ties toward +Infinity (not symmetric).
    // Rust's f64::round() uses round-half-away-from-zero, which breaks for
    // negative halves: Math.round(-0.5) must be 0, not -1.
    unary(
        vm,
        "round",
        Box::new(|_ctx, a| {
            let x = f(a, 0);
            Value::F64((x + 0.5).floor())
        }),
    );
    // ECMA-262 §21.3.2.24/25: max()/min() with no args return -Infinity/+Infinity.
    vm.register_host_fn(
        "ecma:math",
        "min",
        Box::new(|_ctx, a| {
            if a.is_empty() {
                return Value::F64(f64::INFINITY);
            }
            Value::F64(a.iter().map(|v| v.as_f64()).fold(f64::INFINITY, f64::min))
        }),
    );
    vm.register_host_fn(
        "ecma:math",
        "max",
        Box::new(|_ctx, a| {
            if a.is_empty() {
                return Value::F64(f64::NEG_INFINITY);
            }
            Value::F64(
                a.iter()
                    .map(|v| v.as_f64())
                    .fold(f64::NEG_INFINITY, f64::max),
            )
        }),
    );
    binary(
        vm,
        "pow",
        Box::new(|_ctx, a| Value::F64(f(a, 0).powf(f(a, 1)))),
    );
    math_fn(
        vm,
        "random",
        vec![],
        vec![num()],
        Box::new(|_ctx, _| {
            let t = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap_or_default()
                .subsec_nanos();
            Value::F64((t as f64 % 1_000_000.0) / 1_000_000.0)
        }),
    );
    unary(
        vm,
        "sin",
        Box::new(|_ctx, a| Value::F64(f(a, 0).sin())),
    );
    unary(
        vm,
        "cos",
        Box::new(|_ctx, a| Value::F64(f(a, 0).cos())),
    );
    vm.register_host_fn(
        "ecma:math",
        "log",
        Box::new(|_ctx, a| {
            let x = f(a, 0);
            if a.len() > 1 {
                Value::F64(x.ln() / f(a, 1).ln())
            } else {
                Value::F64(x.ln())
            }
        }),
    );
    // Math constants: register as 0-arg host fns first (so CALL_IMPORT via the
    // function index works at runtime), then register_host_value to overwrite
    // the module record ExportEntry with Value — so flatten_module_value_exports
    // sees them as constants and the compiler can inline them directly.
    constant(
        vm,
        "PI",
        Box::new(|_ctx, _| Value::F64(std::f64::consts::PI)),
    );
    constant(
        vm,
        "E",
        Box::new(|_ctx, _| Value::F64(std::f64::consts::E)),
    );
    constant(
        vm,
        "LN2",
        Box::new(|_ctx, _| Value::F64(std::f64::consts::LN_2)),
    );
    constant(
        vm,
        "LN10",
        Box::new(|_ctx, _| Value::F64(std::f64::consts::LN_10)),
    );
    constant(
        vm,
        "LOG2E",
        Box::new(|_ctx, _| Value::F64(std::f64::consts::LOG2_E)),
    );
    constant(
        vm,
        "LOG10E",
        Box::new(|_ctx, _| Value::F64(std::f64::consts::LOG10_E)),
    );
    constant(
        vm,
        "SQRT2",
        Box::new(|_ctx, _| Value::F64(std::f64::consts::SQRT_2)),
    );
    constant(
        vm,
        "SQRT1_2",
        Box::new(|_ctx, _| Value::F64(std::f64::consts::FRAC_1_SQRT_2)),
    );
    // Value registration after fn registration: overwrites ExportEntry::Function
    // with ExportEntry::Value so the compiler inlines these as constants.
    vm.register_host_value("ecma:math", "PI", Value::F64(std::f64::consts::PI));
    vm.register_host_value("ecma:math", "E", Value::F64(std::f64::consts::E));

    unary(
        vm,
        "sign",
        Box::new(|_ctx, a| {
            let n = f(a, 0);
            if n.is_nan() {
                Value::F64(f64::NAN)
            } else if n > 0.0 {
                Value::F64(1.0)
            } else if n < 0.0 {
                Value::F64(-1.0)
            } else {
                Value::F64(0.0)
            }
        }),
    );
    unary(
        vm,
        "log2",
        Box::new(|_ctx, a| Value::F64(f(a, 0).log2())),
    );
    unary(
        vm,
        "log10",
        Box::new(|_ctx, a| Value::F64(f(a, 0).log10())),
    );
    unary(
        vm,
        "cbrt",
        Box::new(|_ctx, a| Value::F64(f(a, 0).cbrt())),
    );
    vm.register_host_fn(
        "ecma:math",
        "hypot",
        Box::new(|_ctx, a| {
            // Variadic: Math.hypot(3, 4) = 5
            let sum: f64 = a
                .iter()
                .map(|v| {
                    let x = v.as_f64();
                    x * x
                })
                .sum();
            Value::F64(sum.sqrt())
        }),
    );
    binary(
        vm,
        "atan2",
        Box::new(|_ctx, a| Value::F64(f(a, 0).atan2(f(a, 1)))),
    );
    unary(
        vm,
        "tan",
        Box::new(|_ctx, a| Value::F64(f(a, 0).tan())),
    );
    unary(
        vm,
        "asin",
        Box::new(|_ctx, a| Value::F64(f(a, 0).asin())),
    );
    unary(
        vm,
        "acos",
        Box::new(|_ctx, a| Value::F64(f(a, 0).acos())),
    );
    unary(
        vm,
        "atan",
        Box::new(|_ctx, a| Value::F64(f(a, 0).atan())),
    );
    unary(
        vm,
        "asinh",
        Box::new(|_ctx, a| Value::F64(f(a, 0).asinh())),
    );
    unary(
        vm,
        "acosh",
        Box::new(|_ctx, a| Value::F64(f(a, 0).acosh())),
    );
    unary(
        vm,
        "atanh",
        Box::new(|_ctx, a| Value::F64(f(a, 0).atanh())),
    );
    unary(
        vm,
        "exp",
        Box::new(|_ctx, a| Value::F64(f(a, 0).exp())),
    );
    unary(
        vm,
        "sinh",
        Box::new(|_ctx, a| Value::F64(f(a, 0).sinh())),
    );
    unary(
        vm,
        "cosh",
        Box::new(|_ctx, a| Value::F64(f(a, 0).cosh())),
    );
    unary(
        vm,
        "tanh",
        Box::new(|_ctx, a| Value::F64(f(a, 0).tanh())),
    );
    // clamp(x, min, max) → emit_clamp (pure WASM F64_MAX + F64_MIN, no host fn).
    unary(
        vm,
        "clz32",
        Box::new(|ctx, a| {
            let n = match number_arg(ctx, a, 0) {
                Ok(n) => n,
                Err(e) => return e,
            };
            Value::F64(to_uint32(n).leading_zeros() as f64)
        }),
    );
    unary(
        vm,
        "fround",
        Box::new(|ctx, a| {
            let n = match number_arg(ctx, a, 0) {
                Ok(n) => n,
                Err(e) => return e,
            };
            let rounded = (n as f32) as f64;
            if rounded == 0.0 && rounded.is_sign_negative() {
                Value::String(Arc::from("-0"))
            } else {
                Value::F64(rounded)
            }
        }),
    );
    // Math.f16round — ES2025: round to nearest IEEE 754 float16 value.
    unary(
        vm,
        "f16round",
        Box::new(|_ctx, a| {
            let x = f(a, 0) as f32;
            // Approximate f16 by truncating f32 mantissa to 10 bits.
            let bits = x.to_bits();
            let exp = (bits >> 23) & 0xFF;
            if exp == 0xFF {
                return Value::F64(x as f64);
            } // inf/nan pass through
            let mantissa = (bits & 0x7FFFFF) >> 13; // keep 10 bits
            let f16_approx_bits = (bits & 0x80000000) | ((exp) << 23) | (mantissa << 13);
            Value::F64(f32::from_bits(f16_approx_bits) as f64)
        }),
    );
    binary(
        vm,
        "imul",
        Box::new(|ctx, a| {
            let x = match number_arg(ctx, a, 0) {
                Ok(n) => to_uint32(n) as i32,
                Err(e) => return e,
            };
            let y = match number_arg(ctx, a, 1) {
                Ok(n) => to_uint32(n) as i32,
                Err(e) => return e,
            };
            Value::F64(x.wrapping_mul(y) as f64)
        }),
    );
    unary(
        vm,
        "expm1",
        Box::new(|_ctx, a| Value::F64(f(a, 0).exp_m1())),
    );
    unary(
        vm,
        "log1p",
        Box::new(|_ctx, a| Value::F64(f(a, 0).ln_1p())),
    );

    // VB `Fix(x)` → walker rewrites to System.Math.Truncate → ecma:math.trunc (§21.3.2.34).
    // VB `Int(x)` → walker rewrites to System.Math.Floor   → ecma:math.floor (§21.3.2.16).
    // VB `Rnd()`  → walker rewrites to System.Math.Random  → ecma:math.random (§21.3.2.27).
    // VB `Sgn(x)` → walker rewrites to System.Math.Sign   → ecma:math.sign  (§21.3.2.29).
    // VB `Randomize` → walker rewrites to Null noop — no host fn needed.
    // C `fmod(a,b)` / PHP `fmod` / Fortran `mod` → emit_c_fmod (pure WASM opcodes — no host fn).

    // ── Stage-3 Math iterator accumulators ──────────────────────────
    //
    // Math.{minOf, maxOf, sumPrecise} accept iterables/arrays directly,
    // letting `min(arr)` / `max(arr)` / `sum(arr)` from Python/Ruby
    // compile to a single host call instead of needing apply/spread.
    //
    // sumPrecise uses the Neumaier-Kahan summation algorithm per the
    // proposal — preserves precision for large sums.

    vm.register_host_fn(
        "ecma:math",
        "minOf",
        Box::new(|_ctx, args| {
            let nums = collect_nums(args);
            // ECMA-262 §21.3.2.25: if any value is NaN, return NaN.
            if nums.iter().any(|x| x.is_nan()) {
                return Value::F64(f64::NAN);
            }
            Value::F64(nums.iter().cloned().fold(f64::INFINITY, f64::min))
        }),
    );
    vm.register_host_fn(
        "ecma:math",
        "maxOf",
        Box::new(|_ctx, args| {
            let nums = collect_nums(args);
            // ECMA-262 §21.3.2.24: if any value is NaN, return NaN.
            if nums.iter().any(|x| x.is_nan()) {
                return Value::F64(f64::NAN);
            }
            Value::F64(nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max))
        }),
    );
    vm.register_host_fn(
        "ecma:math",
        "sumPrecise",
        Box::new(|_ctx, args| {
            let nums = collect_nums(args);
            // Neumaier compensated summation.
            let mut sum = 0.0_f64;
            let mut c = 0.0_f64;
            for x in nums {
                let t = sum + x;
                if sum.abs() >= x.abs() {
                    c += (sum - t) + x;
                } else {
                    c += (x - t) + sum;
                }
                sum = t;
            }
            Value::F64(sum + c)
        }),
    );
}

// Coerce host-fn args into a flat Vec<f64>: accepts a single Array
// argument (the Iterable case from the proposal) or N scalar args.
fn collect_nums(args: &[Value]) -> Vec<f64> {
    if args.len() == 1 {
        if let Some(Value::Object(arr)) = args.first() {
            let o = arr.lock().unwrap();
            if let vybe_runtime::value::ObjectKind::Array(ref v) = o.kind {
                return v.iter().map(|e| e.as_f64()).collect();
            }
        }
    }
    args.iter().map(|v| v.as_f64()).collect()
}

fn f(args: &[Value], idx: usize) -> f64 {
    args.get(idx).map(|v| v.as_f64()).unwrap_or(0.0)
}

fn number_arg(ctx: &mut HostContext, args: &[Value], idx: usize) -> Result<f64, Value> {
    let value = args.get(idx).unwrap_or(&Value::Undefined);
    let primitive = crate::value::to_primitive(ctx, value, "number");
    match primitive {
        Value::Symbol(_) => {
            ctx.throw_value(crate::error::new_error(
                ctx,
                "TypeError",
                "Cannot convert a Symbol value to a number",
            ));
            Err(Value::Undefined)
        }
        Value::String(s) => Ok(string_to_number(&s)),
        other => Ok(other.as_f64()),
    }
}

fn string_to_number(s: &str) -> f64 {
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
    if let Some(rest) = body.strip_prefix("0x").or_else(|| body.strip_prefix("0X")) {
        return i64::from_str_radix(rest, 16)
            .map(|n| sign * n as f64)
            .unwrap_or(f64::NAN);
    }
    trimmed.parse::<f64>().unwrap_or(f64::NAN)
}

fn to_uint32(n: f64) -> u32 {
    if n.is_nan() || n.is_infinite() || n == 0.0 {
        return 0;
    }
    n.trunc().rem_euclid(4_294_967_296.0) as u64 as u32
}
