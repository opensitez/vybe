use super::*;

pub fn register(vm: &mut VM) {
    // Console (VB-style, capitalized)
    let console = ensure_namespace(vm, &["Console"]);
    set_prop(&console, "writeline", host_fn_ref(vm, "wasi:cli", "log"));
    set_prop(&console, "write", host_fn_ref(vm, "wasi:cli", "log"));
    set_prop(&console, "readline", host_fn_ref(vm, "wasi:cli", "readLine"));

    // console (JS-style, lowercase)
    let js_console = ensure_namespace(vm, &["console"]);
    set_prop(&js_console, "log", host_fn_ref(vm, "wasi:cli", "log"));
    set_prop(&js_console, "error", host_fn_ref(vm, "wasi:cli", "error"));
    set_prop(&js_console, "warn", host_fn_ref(vm, "wasi:cli", "warn"));

    // System.Console
    let sys = ensure_namespace(vm, &["System", "Console"]);
    set_prop(&sys, "writeline", host_fn_ref(vm, "wasi:cli", "log"));
    set_prop(&sys, "write", host_fn_ref(vm, "wasi:cli", "log"));
    set_prop(&sys, "error", host_fn_ref(vm, "wasi:cli", "error"));
    set_prop(&sys, "readline", host_fn_ref(vm, "wasi:cli", "readLine"));

    // System.Diagnostics.Debug
    let debug = ensure_namespace(vm, &["System", "Diagnostics", "Debug"]);
    set_prop(&debug, "writeline", host_fn_ref(vm, "wasi:cli", "log"));
    set_prop(&debug, "print", host_fn_ref(vm, "wasi:cli", "log"));
}
