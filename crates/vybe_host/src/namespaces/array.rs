use super::*;

pub fn register(vm: &mut VM) {
    let array = ensure_namespace(vm, &["Array"]);
    set_prop(&array, "isarray", host_fn_ref(vm, "vybe:array", "isArray"));
    set_prop(&array, "from", host_fn_ref(vm, "vybe:array", "from"));

    // System.Array
    let sys = ensure_namespace(vm, &["System", "Array"]);
    set_prop(&sys, "reverse", host_fn_ref(vm, "vybe:array", "reverse"));
    set_prop(&sys, "indexof", host_fn_ref(vm, "vybe:array", "indexOf"));

    // System.Collections.Generic
    let coll = ensure_namespace(vm, &["System", "Collections", "Generic"]);
    set_prop(&coll, "list", host_fn_ref(vm, "vybe:array", "from"));
    set_prop(&coll, "dictionary", host_fn_ref(vm, "vybe:collections", "Map"));
}
