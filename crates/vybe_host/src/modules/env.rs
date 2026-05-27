use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::Object;

pub(crate) fn machine_name_value() -> Value {
    let name = std::env::var("HOSTNAME")
        .or_else(|_| std::env::var("COMPUTERNAME"))
        .unwrap_or_else(|_| "unknown".into());
    Value::String(Arc::from(name.as_str()))
}

// `register_dotnet_net` retired — `Dns.GetHostName()` lowers to
// `node:os.hostname()` via `emitter::dotnet::core::sockets_adapter`.

pub fn register(vm: &mut VM) {
    // ── wasi:cli/environment — WASI CLI proposal interface ───────────
    // 0.2.x exports `initial-cwd`; 0.3.x renamed it to
    // `get-initial-cwd`. We expose both names on the unversioned
    // interface module so callers targeting either proposal revision
    // can bind the actual CLI environment surface.
    vm.register_host_fn("wasi:cli/environment", "get-environment", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let pairs: Vec<Value> = std::env::vars()
            .map(|(key, value)| {
                Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
                    Value::String(Arc::from(key.as_str())),
                    Value::String(Arc::from(value.as_str())),
                ]))))
            })
            .collect();
        Value::Object(Arc::new(Mutex::new(Object::new_array(pairs))))
    }));

    vm.register_host_fn("wasi:cli/environment", "get-arguments", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let args: Vec<Value> = std::env::args()
            .map(|arg| Value::String(Arc::from(arg.as_str())))
            .collect();
        Value::Object(Arc::new(Mutex::new(Object::new_array(args))))
    }));

    vm.register_host_fn("wasi:cli/environment", "initial-cwd", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        match std::env::current_dir() {
            Ok(path) => Value::String(Arc::from(path.to_string_lossy().as_ref())),
            Err(_) => Value::Null,
        }
    }));

    vm.register_host_fn("wasi:cli/environment", "get-initial-cwd", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        match std::env::current_dir() {
            Ok(path) => Value::String(Arc::from(path.to_string_lossy().as_ref())),
            Err(_) => Value::Null,
        }
    }));

    // ── wasi:cli (legacy flat compatibility namespace) ───────────────
    // Get command line arguments as array of strings
    vm.register_host_fn("wasi:cli", "args", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let args: Vec<Value> = std::env::args()
            .map(|a| Value::String(Arc::from(a.as_str())))
            .collect();
        Value::Object(Arc::new(Mutex::new(Object::new_array(args))))
    }));

    // Get environment variable (returns null if not set)
    vm.register_host_fn("wasi:cli", "getEnv", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        match std::env::var(&key) {
            Ok(val) => Value::String(Arc::from(val.as_str())),
            Err(_) => Value::Null,
        }
    }));

    // Get current working directory
    vm.register_host_fn("wasi:cli", "cwd", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        match std::env::current_dir() {
            Ok(p) => Value::String(Arc::from(p.to_string_lossy().as_ref())),
            Err(_) => Value::Null,
        }
    }));

    // Get platform name
    vm.register_host_fn("wasi:cli", "platform", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::String(Arc::from(std::env::consts::OS))
    }));

    // Get architecture
    vm.register_host_fn("wasi:cli", "arch", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::String(Arc::from(std::env::consts::ARCH))
    }));

    // Machine name
    vm.register_host_fn("wasi:cli", "machineName", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        machine_name_value()
    }));

    // User name
    vm.register_host_fn("wasi:cli", "userName", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::String(Arc::from(std::env::var("USER").or_else(|_| std::env::var("USERNAME")).unwrap_or_else(|_| "unknown".into()).as_str()))
    }));

    // Newline
    vm.register_host_fn("wasi:cli", "newLine", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::String(Arc::from("\n"))
    }));

    // Tick count (ms since boot, approximated as ms since epoch)
    vm.register_host_fn("wasi:cli", "tickCount", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let ms = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_millis();
        Value::F64(ms as f64)
    }));

    // GetFolderPath — returns home dir for any special folder enum value
    vm.register_host_fn("wasi:cli", "getFolderPath", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let path = std::env::var("HOME")
            .or_else(|_| std::env::var("USERPROFILE"))
            .unwrap_or_else(|_| ".".into());
        Value::String(Arc::from(path.as_str()))
    }));
}
