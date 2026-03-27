use super::*;

pub fn register(vm: &mut VM) {
    // Convert (direct shortcut)
    let convert = ensure_namespace(vm, &["Convert"]);
    register_convert_methods(vm, &convert);

    // System.Convert
    let sys = ensure_namespace(vm, &["System", "Convert"]);
    register_convert_methods(vm, &sys);
}

fn register_convert_methods(vm: &VM, ns: &Value) {
    set_prop(ns, "toint32", host_fn_ref(vm, "vybe:convert", "cint"));
    set_prop(ns, "toint16", host_fn_ref(vm, "vybe:convert", "cint"));
    set_prop(ns, "toint64", host_fn_ref(vm, "vybe:convert", "clng"));
    set_prop(ns, "todouble", host_fn_ref(vm, "vybe:convert", "cdbl"));
    set_prop(ns, "tosingle", host_fn_ref(vm, "vybe:convert", "csng"));
    set_prop(ns, "tostring", host_fn_ref(vm, "vybe:convert", "toString"));
    set_prop(ns, "toboolean", host_fn_ref(vm, "vybe:convert", "cbool"));
    set_prop(ns, "tobyte", host_fn_ref(vm, "vybe:convert", "cbyte"));
    set_prop(ns, "tochar", host_fn_ref(vm, "vybe:convert", "cchar"));
}
