use super::*;

pub fn register(vm: &mut VM) {
    // Convert (direct shortcut)
    let convert = ensure_namespace(vm, &["Convert"]);
    register_convert_methods(vm, &convert);

    // System.Convert
    let sys = ensure_namespace(vm, &["System", "Convert"]);
    register_convert_methods(vm, &sys);
}

/// VB/.NET `Convert.ToXxx(v)` — bound to ECMA-262 §21 Number / §22.1
/// String / §20.3 Boolean primitives. Argument-shape and edge-case
/// semantics that diverge from .NET (e.g. `Convert.ToInt32(null)` →
/// 0 in .NET, `parseInt(null)` → NaN in JS) are caller-side concerns:
/// language adapters in the emitter wrap these calls when exact .NET
/// semantics are required.
fn register_convert_methods(vm: &VM, ns: &Value) {
    set_prop(ns, "todatetime", host_fn_ref(vm, "ecma:date", "now"));
    set_prop(ns, "toint32", host_fn_ref(vm, "ecma:number", "parseInt"));
    set_prop(ns, "toint16", host_fn_ref(vm, "ecma:number", "parseInt"));
    set_prop(ns, "toint64", host_fn_ref(vm, "ecma:number", "parseInt"));
    set_prop(ns, "todouble", host_fn_ref(vm, "ecma:number", "parseFloat"));
    set_prop(ns, "tosingle", host_fn_ref(vm, "ecma:number", "parseFloat"));
    set_prop(ns, "tostring", host_fn_ref(vm, "ecma:string", "String"));
    set_prop(ns, "toboolean", host_fn_ref(vm, "ecma:boolean", "Boolean"));
    set_prop(ns, "tobyte", host_fn_ref(vm, "ecma:number", "parseInt"));
    set_prop(ns, "tochar", host_fn_ref(vm, "ecma:string", "fromCharCode"));
}
