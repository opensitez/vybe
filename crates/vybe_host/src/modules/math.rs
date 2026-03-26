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
}

fn f(args: &[Value], idx: usize) -> f64 {
    args.get(idx).map(|v| v.as_f64()).unwrap_or(0.0)
}
