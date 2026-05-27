use super::*;

pub fn register(vm: &mut VM) {
    // Math (direct)
    let math = ensure_namespace(vm, &["Math"]);
    for name in &[
        "floor", "ceil", "round", "abs", "sqrt", "pow", "min", "max",
        "sin", "cos", "tan", "log", "sign", "trunc", "log2", "log10",
        "cbrt", "hypot", "atan2", "asin", "acos", "atan", "exp", "clz32",
    ] {
        set_prop(&math, name, host_fn_ref(vm, "ecma:math", name));
    }
    set_prop(&math, "pi", Value::F64(std::f64::consts::PI));
    set_prop(&math, "e", Value::F64(std::f64::consts::E));
}
