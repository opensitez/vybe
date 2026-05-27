use super::*;

pub fn register(vm: &mut VM) {
    let array = ensure_namespace(vm, &["Array"]);
    set_prop(&array, "isarray", host_fn_ref(vm, "ecma:array", "isArray"));
    set_prop(&array, "from", host_fn_ref(vm, "ecma:array", "from"));
}
