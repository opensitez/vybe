use vybe_bytecode::{VM, Value};

pub fn register(vm: &mut VM) {
    vm.register_host_fn("vybe:convert", "parseInt", Box::new(|args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        match s.trim().parse::<i64>() {
            Ok(n) => Value::F64(n as f64),
            Err(_) => match s.trim().parse::<f64>() {
                Ok(n) => Value::F64(n.trunc()),
                Err(_) => Value::F64(f64::NAN),
            }
        }
    }));
    vm.register_host_fn("vybe:convert", "parseFloat", Box::new(|args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        Value::F64(s.trim().parse::<f64>().unwrap_or(f64::NAN))
    }));
    vm.register_host_fn("vybe:convert", "toString", Box::new(|args: &[Value]| {
        Value::String(std::rc::Rc::from(format!("{}", args.first().unwrap_or(&Value::Null)).as_str()))
    }));
    vm.register_host_fn("vybe:convert", "isNaN", Box::new(|args: &[Value]| {
        Value::Bool(args.first().map(|v| v.as_f64().is_nan()).unwrap_or(true))
    }));
    vm.register_host_fn("vybe:convert", "isFinite", Box::new(|args: &[Value]| {
        Value::Bool(args.first().map(|v| v.as_f64().is_finite()).unwrap_or(false))
    }));
}
