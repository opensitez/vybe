use vybe_bytecode::{VM, Value};

pub fn register(vm: &mut VM) {
    // Core math — callable ops stay as host functions, while spec
    // constants are registered as immutable value exports.
    vm.register_host_fn("ecma:math", "floor",  Box::new(|_ctx, a| Value::F64(f(a, 0).floor())));
    vm.register_host_fn("ecma:math", "ceil",   Box::new(|_ctx, a| Value::F64(f(a, 0).ceil())));
    vm.register_host_fn("ecma:math", "abs",    Box::new(|_ctx, a| Value::F64(f(a, 0).abs())));
    vm.register_host_fn("ecma:math", "sqrt",   Box::new(|_ctx, a| Value::F64(f(a, 0).sqrt())));
    vm.register_host_fn("ecma:math", "trunc",  Box::new(|_ctx, a| Value::F64(f(a, 0).trunc())));
    vm.register_host_fn("ecma:math", "round",  Box::new(|_ctx, a| Value::F64(f(a, 0).round())));
    vm.register_host_fn("ecma:math", "min",    Box::new(|_ctx, a| Value::F64(f(a, 0).min(f(a, 1)))));
    vm.register_host_fn("ecma:math", "max",    Box::new(|_ctx, a| Value::F64(f(a, 0).max(f(a, 1)))));
    vm.register_host_fn("ecma:math", "pow",    Box::new(|_ctx, a| Value::F64(f(a, 0).powf(f(a, 1)))));
    vm.register_host_fn("ecma:math", "fmod",   Box::new(|_ctx, a| Value::F64(f(a, 0) % f(a, 1))));
    vm.register_host_fn("ecma:math", "random", Box::new(|_ctx, _| {
        let t = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .subsec_nanos();
        Value::F64((t as f64 % 1_000_000.0) / 1_000_000.0)
    }));
    vm.register_host_fn("ecma:math", "sin",  Box::new(|_ctx, a| Value::F64(f(a, 0).sin())));
    vm.register_host_fn("ecma:math", "cos",  Box::new(|_ctx, a| Value::F64(f(a, 0).cos())));
    vm.register_host_fn("ecma:math", "log",  Box::new(|_ctx, a| {
        let x = f(a, 0);
        if a.len() > 1 { Value::F64(x.ln() / f(a, 1).ln()) } else { Value::F64(x.ln()) }
    }));
    vm.register_host_value("ecma:math", "PI", Value::F64(std::f64::consts::PI));
    vm.register_host_value("ecma:math", "E", Value::F64(std::f64::consts::E));
    vm.register_host_fn("ecma:math", "sign",  Box::new(|_ctx, a| {
        let n = f(a, 0);
        if n > 0.0 { Value::F64(1.0) } else if n < 0.0 { Value::F64(-1.0) } else { Value::F64(0.0) }
    }));
    vm.register_host_fn("ecma:math", "log2",  Box::new(|_ctx, a| Value::F64(f(a, 0).log2())));
    vm.register_host_fn("ecma:math", "log10", Box::new(|_ctx, a| Value::F64(f(a, 0).log10())));
    vm.register_host_fn("ecma:math", "cbrt",  Box::new(|_ctx, a| Value::F64(f(a, 0).cbrt())));
    vm.register_host_fn("ecma:math", "hypot", Box::new(|_ctx, a| Value::F64(f(a, 0).hypot(f(a, 1)))));
    vm.register_host_fn("ecma:math", "atan2", Box::new(|_ctx, a| Value::F64(f(a, 0).atan2(f(a, 1)))));
    vm.register_host_fn("ecma:math", "tan",   Box::new(|_ctx, a| Value::F64(f(a, 0).tan())));
    vm.register_host_fn("ecma:math", "asin",  Box::new(|_ctx, a| Value::F64(f(a, 0).asin())));
    vm.register_host_fn("ecma:math", "acos",  Box::new(|_ctx, a| Value::F64(f(a, 0).acos())));
    vm.register_host_fn("ecma:math", "atan",  Box::new(|_ctx, a| Value::F64(f(a, 0).atan())));
    vm.register_host_fn("ecma:math", "exp",   Box::new(|_ctx, a| Value::F64(f(a, 0).exp())));
    vm.register_host_fn("ecma:math", "sinh",  Box::new(|_ctx, a| Value::F64(f(a, 0).sinh())));
    vm.register_host_fn("ecma:math", "cosh",  Box::new(|_ctx, a| Value::F64(f(a, 0).cosh())));
    vm.register_host_fn("ecma:math", "tanh",  Box::new(|_ctx, a| Value::F64(f(a, 0).tanh())));
    vm.register_host_fn("ecma:math", "clamp", Box::new(|_ctx, a| Value::F64(f(a, 0).clamp(f(a, 1), f(a, 2)))));
    vm.register_host_fn("ecma:math", "clz32", Box::new(|_ctx, a| Value::F64((f(a, 0) as u32).leading_zeros() as f64)));
    vm.register_host_fn("ecma:math", "fround", Box::new(|_ctx, a| Value::F64((f(a, 0) as f32) as f64)));
    vm.register_host_fn("ecma:math", "imul",   Box::new(|_ctx, a| {
        let x = f(a, 0) as i64 as i32;
        let y = f(a, 1) as i64 as i32;
        Value::F64(x.wrapping_mul(y) as f64)
    }));
    vm.register_host_fn("ecma:math", "expm1",  Box::new(|_ctx, a| Value::F64(f(a, 0).exp_m1())));
    vm.register_host_fn("ecma:math", "log1p",  Box::new(|_ctx, a| Value::F64(f(a, 0).ln_1p())));

    // VB-specific math functions
    vm.register_host_fn("ecma:math", "fix",   Box::new(|_ctx, a| Value::F64(f(a, 0).trunc()))); // truncate toward zero
    vm.register_host_fn("ecma:math", "int",   Box::new(|_ctx, a| Value::F64(f(a, 0).floor()))); // VB Int = floor
    vm.register_host_fn("ecma:math", "sgn",   Box::new(|_ctx, a| {
        let n = f(a, 0);
        Value::F64(if n > 0.0 { 1.0 } else if n < 0.0 { -1.0 } else { 0.0 })
    }));
    vm.register_host_fn("ecma:math", "rnd",   Box::new(|_ctx, _a| {
        let t = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .subsec_nanos();
        Value::F64((t as f64 % 1_000_000.0) / 1_000_000.0)
    }));
    vm.register_host_fn("ecma:math", "randomize", Box::new(|_ctx, _a| Value::Null));

    // ── Stage-3 Math iterator accumulators ──────────────────────────
    //
    // Math.{minOf, maxOf, sumPrecise} accept iterables/arrays directly,
    // letting `min(arr)` / `max(arr)` / `sum(arr)` from Python/Ruby
    // compile to a single host call instead of needing apply/spread.
    //
    // sumPrecise uses the Neumaier-Kahan summation algorithm per the
    // proposal — preserves precision for large sums.

    vm.register_host_fn("ecma:math", "minOf", Box::new(|_ctx, args| {
        let nums = collect_nums(args);
        Value::F64(nums.iter().cloned().fold(f64::INFINITY, f64::min))
    }));
    vm.register_host_fn("ecma:math", "maxOf", Box::new(|_ctx, args| {
        let nums = collect_nums(args);
        Value::F64(nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max))
    }));
    vm.register_host_fn("ecma:math", "sumPrecise", Box::new(|_ctx, args| {
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
    }));
}

// Coerce host-fn args into a flat Vec<f64>: accepts a single Array
// argument (the Iterable case from the proposal) or N scalar args.
fn collect_nums(args: &[Value]) -> Vec<f64> {
    if args.len() == 1 {
        if let Some(Value::Object(arr)) = args.first() {
            let o = arr.lock().unwrap();
            if let vybe_bytecode::value::ObjectKind::Array(ref v) = o.kind {
                return v.iter().map(|e| e.as_f64()).collect();
            }
        }
    }
    args.iter().map(|v| v.as_f64()).collect()
}


fn f(args: &[Value], idx: usize) -> f64 {
    args.get(idx).map(|v| v.as_f64()).unwrap_or(0.0)
}
