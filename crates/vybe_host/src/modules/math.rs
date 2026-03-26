use vybe_bytecode::{VM, Value};

pub fn register(vm: &mut VM) {
    vm.register_host_fn("vybe:math", "floor",  Box::new(|a| Value::F64(f(a, 0).floor())));
    vm.register_host_fn("vybe:math", "ceil",   Box::new(|a| Value::F64(f(a, 0).ceil())));
    vm.register_host_fn("vybe:math", "round",  Box::new(|a| Value::F64(f(a, 0).round())));
    vm.register_host_fn("vybe:math", "abs",    Box::new(|a| Value::F64(f(a, 0).abs())));
    vm.register_host_fn("vybe:math", "sqrt",   Box::new(|a| Value::F64(f(a, 0).sqrt())));
    vm.register_host_fn("vybe:math", "pow",    Box::new(|a| Value::F64(f(a, 0).powf(f(a, 1)))));
    vm.register_host_fn("vybe:math", "min",    Box::new(|a| Value::F64(f(a, 0).min(f(a, 1)))));
    vm.register_host_fn("vybe:math", "max",    Box::new(|a| Value::F64(f(a, 0).max(f(a, 1)))));
    vm.register_host_fn("vybe:math", "random", Box::new(|_| {
        let t = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .subsec_nanos();
        Value::F64((t as f64 % 1_000_000.0) / 1_000_000.0)
    }));
    vm.register_host_fn("vybe:math", "sin",  Box::new(|a| Value::F64(f(a, 0).sin())));
    vm.register_host_fn("vybe:math", "cos",  Box::new(|a| Value::F64(f(a, 0).cos())));
    vm.register_host_fn("vybe:math", "log",  Box::new(|a| Value::F64(f(a, 0).ln())));
    vm.register_host_fn("vybe:math", "PI",   Box::new(|_| Value::F64(std::f64::consts::PI)));
    vm.register_host_fn("vybe:math", "E",    Box::new(|_| Value::F64(std::f64::consts::E)));
    vm.register_host_fn("vybe:math", "trunc", Box::new(|a| Value::F64(f(a, 0).trunc())));
    vm.register_host_fn("vybe:math", "sign",  Box::new(|a| {
        let n = f(a, 0);
        if n > 0.0 { Value::F64(1.0) } else if n < 0.0 { Value::F64(-1.0) } else { Value::F64(0.0) }
    }));
    vm.register_host_fn("vybe:math", "log2",  Box::new(|a| Value::F64(f(a, 0).log2())));
    vm.register_host_fn("vybe:math", "log10", Box::new(|a| Value::F64(f(a, 0).log10())));
    vm.register_host_fn("vybe:math", "cbrt",  Box::new(|a| Value::F64(f(a, 0).cbrt())));
    vm.register_host_fn("vybe:math", "hypot", Box::new(|a| Value::F64(f(a, 0).hypot(f(a, 1)))));
    vm.register_host_fn("vybe:math", "atan2", Box::new(|a| Value::F64(f(a, 0).atan2(f(a, 1)))));
    vm.register_host_fn("vybe:math", "tan",   Box::new(|a| Value::F64(f(a, 0).tan())));
    vm.register_host_fn("vybe:math", "asin",  Box::new(|a| Value::F64(f(a, 0).asin())));
    vm.register_host_fn("vybe:math", "acos",  Box::new(|a| Value::F64(f(a, 0).acos())));
    vm.register_host_fn("vybe:math", "atan",  Box::new(|a| Value::F64(f(a, 0).atan())));
    vm.register_host_fn("vybe:math", "exp",   Box::new(|a| Value::F64(f(a, 0).exp())));
    vm.register_host_fn("vybe:math", "clz32", Box::new(|a| Value::F64((f(a, 0) as u32).leading_zeros() as f64)));
}

fn f(args: &[Value], idx: usize) -> f64 {
    args.get(idx).map(|v| v.as_f64()).unwrap_or(0.0)
}
