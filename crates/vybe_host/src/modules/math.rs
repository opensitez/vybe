use vybe_bytecode::{VM, Value};

pub fn register(vm: &mut VM) {
    // Core math — also available as opcodes, but registered as host fns
    // so namespace objects can reference them
    vm.register_host_fn("vybe:math", "floor",  Box::new(|_ctx, a| Value::F64(f(a, 0).floor())));
    vm.register_host_fn("vybe:math", "ceil",   Box::new(|_ctx, a| Value::F64(f(a, 0).ceil())));
    vm.register_host_fn("vybe:math", "abs",    Box::new(|_ctx, a| Value::F64(f(a, 0).abs())));
    vm.register_host_fn("vybe:math", "sqrt",   Box::new(|_ctx, a| Value::F64(f(a, 0).sqrt())));
    vm.register_host_fn("vybe:math", "trunc",  Box::new(|_ctx, a| Value::F64(f(a, 0).trunc())));
    vm.register_host_fn("vybe:math", "round",  Box::new(|_ctx, a| Value::F64(f(a, 0).round())));
    vm.register_host_fn("vybe:math", "min",    Box::new(|_ctx, a| Value::F64(f(a, 0).min(f(a, 1)))));
    vm.register_host_fn("vybe:math", "max",    Box::new(|_ctx, a| Value::F64(f(a, 0).max(f(a, 1)))));
    vm.register_host_fn("vybe:math", "pow",    Box::new(|_ctx, a| Value::F64(f(a, 0).powf(f(a, 1)))));
    vm.register_host_fn("vybe:math", "fmod",   Box::new(|_ctx, a| Value::F64(f(a, 0) % f(a, 1))));
    vm.register_host_fn("vybe:math", "random", Box::new(|_ctx, _| {
        let t = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .subsec_nanos();
        Value::F64((t as f64 % 1_000_000.0) / 1_000_000.0)
    }));
    vm.register_host_fn("vybe:math", "sin",  Box::new(|_ctx, a| Value::F64(f(a, 0).sin())));
    vm.register_host_fn("vybe:math", "cos",  Box::new(|_ctx, a| Value::F64(f(a, 0).cos())));
    vm.register_host_fn("vybe:math", "log",  Box::new(|_ctx, a| Value::F64(f(a, 0).ln())));
    vm.register_host_fn("vybe:math", "PI",   Box::new(|_ctx, _| Value::F64(std::f64::consts::PI)));
    vm.register_host_fn("vybe:math", "E",    Box::new(|_ctx, _| Value::F64(std::f64::consts::E)));
    vm.register_host_fn("vybe:math", "sign",  Box::new(|_ctx, a| {
        let n = f(a, 0);
        if n > 0.0 { Value::F64(1.0) } else if n < 0.0 { Value::F64(-1.0) } else { Value::F64(0.0) }
    }));
    vm.register_host_fn("vybe:math", "log2",  Box::new(|_ctx, a| Value::F64(f(a, 0).log2())));
    vm.register_host_fn("vybe:math", "log10", Box::new(|_ctx, a| Value::F64(f(a, 0).log10())));
    vm.register_host_fn("vybe:math", "cbrt",  Box::new(|_ctx, a| Value::F64(f(a, 0).cbrt())));
    vm.register_host_fn("vybe:math", "hypot", Box::new(|_ctx, a| Value::F64(f(a, 0).hypot(f(a, 1)))));
    vm.register_host_fn("vybe:math", "atan2", Box::new(|_ctx, a| Value::F64(f(a, 0).atan2(f(a, 1)))));
    vm.register_host_fn("vybe:math", "tan",   Box::new(|_ctx, a| Value::F64(f(a, 0).tan())));
    vm.register_host_fn("vybe:math", "asin",  Box::new(|_ctx, a| Value::F64(f(a, 0).asin())));
    vm.register_host_fn("vybe:math", "acos",  Box::new(|_ctx, a| Value::F64(f(a, 0).acos())));
    vm.register_host_fn("vybe:math", "atan",  Box::new(|_ctx, a| Value::F64(f(a, 0).atan())));
    vm.register_host_fn("vybe:math", "exp",   Box::new(|_ctx, a| Value::F64(f(a, 0).exp())));
    vm.register_host_fn("vybe:math", "clz32", Box::new(|_ctx, a| Value::F64((f(a, 0) as u32).leading_zeros() as f64)));

    // VB-specific math functions
    vm.register_host_fn("vybe:math", "fix",   Box::new(|_ctx, a| Value::F64(f(a, 0).trunc()))); // truncate toward zero
    vm.register_host_fn("vybe:math", "int",   Box::new(|_ctx, a| Value::F64(f(a, 0).floor()))); // VB Int = floor
    vm.register_host_fn("vybe:math", "sgn",   Box::new(|_ctx, a| {
        let n = f(a, 0);
        Value::F64(if n > 0.0 { 1.0 } else if n < 0.0 { -1.0 } else { 0.0 })
    }));
    vm.register_host_fn("vybe:math", "rnd",   Box::new(|_ctx, _a| {
        let t = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .subsec_nanos();
        Value::F64((t as f64 % 1_000_000.0) / 1_000_000.0)
    }));
    vm.register_host_fn("vybe:math", "randomize", Box::new(|_ctx, _a| Value::Null));
}

fn f(args: &[Value], idx: usize) -> f64 {
    args.get(idx).map(|v| v.as_f64()).unwrap_or(0.0)
}
