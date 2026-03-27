use super::*;

pub fn register(vm: &mut VM) {
    // Math (direct)
    let math = ensure_namespace(vm, &["Math"]);
    for name in &[
        "floor", "ceil", "round", "abs", "sqrt", "pow", "min", "max",
        "sin", "cos", "tan", "log", "sign", "trunc", "log2", "log10",
        "cbrt", "hypot", "atan2", "asin", "acos", "atan", "exp", "clz32",
    ] {
        set_prop(&math, name, host_fn_ref(vm, "vybe:math", name));
    }
    set_prop(&math, "pi", Value::F64(std::f64::consts::PI));
    set_prop(&math, "e", Value::F64(std::f64::consts::E));

    // System.Math (alias — shares the same object)
    let sys_math = ensure_namespace(vm, &["System", "Math"]);
    for name in &[
        "floor", "ceil", "round", "abs", "sqrt", "pow", "min", "max",
        "sin", "cos", "tan", "log", "sign", "trunc", "log2", "log10",
        "cbrt", "hypot", "atan2", "asin", "acos", "atan", "exp",
    ] {
        set_prop(&sys_math, name, host_fn_ref(vm, "vybe:math", name));
    }
    set_prop(&sys_math, "pi", Value::F64(std::f64::consts::PI));
    set_prop(&sys_math, "e", Value::F64(std::f64::consts::E));
}
