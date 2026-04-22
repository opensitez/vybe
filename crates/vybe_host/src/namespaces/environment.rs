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
    // 0-arg properties — computed once at startup and stored as values
    let cwd = std::env::current_dir()
        .map(|p| Value::String(Arc::from(p.to_string_lossy().as_ref())))
        .unwrap_or(Value::Null);
    set_prop(ns, "currentdirectory", cwd);

    let machine = std::env::var("HOSTNAME")
        .or_else(|_| std::env::var("HOST"))
        .or_else(|_| std::env::var("COMPUTERNAME"))
        .unwrap_or_else(|_| {
            std::env::var("USER").map(|u| format!("{}-mac", u)).unwrap_or_else(|_| "unknown".into())
        });
    set_prop(ns, "machinename", Value::String(Arc::from(machine.as_str())));

    let user = std::env::var("USER")
        .or_else(|_| std::env::var("USERNAME"))
        .unwrap_or_else(|_| "unknown".into());
    set_prop(ns, "username", Value::String(Arc::from(user.as_str())));

    set_prop(ns, "osversion", Value::String(Arc::from(std::env::consts::OS)));
    set_prop(ns, "newline", Value::String(Arc::from("\n")));

    let nprocs = std::thread::available_parallelism().map(|n| n.get()).unwrap_or(1);
    set_prop(ns, "processorcount", Value::F64(nprocs as f64));

    set_prop(ns, "is64bitoperatingsystem", Value::Bool(cfg!(target_pointer_width = "64")));
    set_prop(ns, "is64bitprocess", Value::Bool(cfg!(target_pointer_width = "64")));

    let ms = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_millis();
    set_prop(ns, "tickcount", Value::F64(ms as f64));

    // Functions called with args — keep as function refs
    set_prop(ns, "getenvironmentvariable", host_fn_ref(vm, "wasi:cli", "getEnv"));
    set_prop(ns, "setenvironmentvariable", Value::Null); // noop
    set_prop(ns, "getfolderpath", host_fn_ref(vm, "wasi:cli", "getFolderPath"));
    set_prop(ns, "commandline", host_fn_ref(vm, "wasi:cli", "args"));
    set_prop(ns, "exit", host_fn_ref(vm, "wasi:cli", "exit"));
}
