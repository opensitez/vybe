use super::*;

pub fn register(vm: &mut VM) {
    // System.Net.Http
    let http = ensure_namespace(vm, &["System", "Net", "Http"]);
    set_prop(&http, "get", host_fn_ref(vm, "wasi:http", "get"));
    set_prop(&http, "post", host_fn_ref(vm, "wasi:http", "post"));
    set_prop(&http, "fetch", host_fn_ref(vm, "wasi:http", "fetch"));
}
