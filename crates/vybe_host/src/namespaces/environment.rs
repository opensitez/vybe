use super::*;

pub fn register(vm: &mut VM) {
    // Environment (direct shortcut)
    let env = ensure_namespace(vm, &["Environment"]);
    register_env_methods(vm, &env);

    // System.Environment
    let sys_env = ensure_namespace(vm, &["System", "Environment"]);
    register_env_methods(vm, &sys_env);

    // Thread (direct shortcut)
    let thread = ensure_namespace(vm, &["Thread"]);
    set_prop(&thread, "sleep", host_fn_ref(vm, "wasi:clocks", "sleep"));

    // System.Threading.Thread
    let sys_thread = ensure_namespace(vm, &["System", "Threading", "Thread"]);
    set_prop(&sys_thread, "sleep", host_fn_ref(vm, "wasi:clocks", "sleep"));
}

fn register_env_methods(vm: &VM, ns: &Value) {
    set_prop(ns, "getenvironmentvariable", host_fn_ref(vm, "wasi:cli", "getEnv"));
    set_prop(ns, "currentdirectory", host_fn_ref(vm, "wasi:cli", "cwd"));
    set_prop(ns, "machinename", host_fn_ref(vm, "wasi:cli", "machineName"));
    set_prop(ns, "username", host_fn_ref(vm, "wasi:cli", "userName"));
    set_prop(ns, "osversion", host_fn_ref(vm, "wasi:cli", "platform"));
    set_prop(ns, "newline", Value::String(Arc::from("\n")));
    set_prop(ns, "tickcount", host_fn_ref(vm, "wasi:cli", "tickCount"));
    set_prop(ns, "commandline", host_fn_ref(vm, "wasi:cli", "args"));
    set_prop(ns, "exit", host_fn_ref(vm, "wasi:cli", "exit"));
}
