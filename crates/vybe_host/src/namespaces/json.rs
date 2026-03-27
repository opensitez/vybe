use super::*;

pub fn register(vm: &mut VM) {
    let json = ensure_namespace(vm, &["JSON"]);
    set_prop(&json, "stringify", host_fn_ref(vm, "vybe:json", "stringify"));
    set_prop(&json, "parse", host_fn_ref(vm, "vybe:json", "parse"));
}
