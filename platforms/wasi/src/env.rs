use std::sync::Arc;
use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

// `register_dotnet_net` retired — `Dns.GetHostName()` lowers to
// `node:os.hostname()` via `emitter::dotnet::core::sockets_adapter`.

pub fn register(vm: &mut VM) {
    // ── wasi:cli/environment — WASI CLI proposal interface ───────────
    // 0.2.x exports `initial-cwd`; 0.3.x renamed it to
    // `get-initial-cwd`. We expose both names on the unversioned
    // interface module so callers targeting either proposal revision
    // can bind the actual CLI environment surface.
    // Spec: get-environment() → list<(string, string)>
    // Extension: if called with 1 arg (key), returns value for that key or null.
    vm.register_host_fn(
        "wasi:cli/environment",
        "get-environment",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(key_val) = args.first() {
                let key = format!("{}", key_val);
                return match std::env::var(&key) {
                    Ok(val) => Value::String(Arc::from(val.as_str())),
                    Err(_) => Value::Null,
                };
            }
            let pairs: Vec<Value> = std::env::vars()
                .map(|(key, value)| {
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
                        Value::String(Arc::from(key.as_str())),
                        Value::String(Arc::from(value.as_str())),
                    ])))
                })
                .collect();
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(pairs)))
        }),
    );

    vm.register_host_fn(
        "wasi:cli/environment",
        "get-arguments",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let args: Vec<Value> = std::env::args()
                .map(|arg| Value::String(Arc::from(arg.as_str())))
                .collect();
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(args)))
        }),
    );

    vm.register_host_fn(
        "wasi:cli/environment",
        "initial-cwd",
        Box::new(
            |_ctx: &mut HostContext, _args: &[Value]| match std::env::current_dir() {
                Ok(path) => Value::String(Arc::from(path.to_string_lossy().as_ref())),
                Err(_) => Value::Null,
            },
        ),
    );

    vm.register_host_fn(
        "wasi:cli/environment",
        "get-initial-cwd",
        Box::new(
            |_ctx: &mut HostContext, _args: &[Value]| match std::env::current_dir() {
                Ok(path) => Value::String(Arc::from(path.to_string_lossy().as_ref())),
                Err(_) => Value::Null,
            },
        ),
    );
}
