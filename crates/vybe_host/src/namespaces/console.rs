use super::*;

pub fn register(vm: &mut VM) {
    // console (JS-style, lowercase)
    let js_console = ensure_namespace(vm, &["console"]);
    set_prop(&js_console, "log", host_fn_ref(vm, "wasi:cli", "log"));
    set_prop(&js_console, "error", host_fn_ref(vm, "wasi:cli", "error"));
    set_prop(&js_console, "warn", host_fn_ref(vm, "wasi:cli", "warn"));
}
