use std::cell::RefCell;
use std::sync::Arc;
use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

// `register_dotnet_net` retired — `Dns.GetHostName()` lowers to
// `node:os.hostname()` via `emitter::dotnet::core::sockets_adapter`.

thread_local! {
    /// Variables the host installed for the instance running on this thread.
    ///
    /// WASI's model is one component instance per request, and
    /// `get-environment` reports *that instance's* environment — which is how
    /// CGI has always passed `DOCUMENT_ROOT`, `SCRIPT_FILENAME`,
    /// `SERVER_SOFTWARE` and friends. `vybex --serve` runs each request on its
    /// own blocking worker, so "this instance" is "this thread", exactly as
    /// `node::http::install_context` already models the request itself.
    ///
    /// This is a Rust-side install API, NOT a host function: nothing new is
    /// exported under `wasi:cli/environment`, so the namespace stays spec-pure.
    static INSTANCE_ENVIRONMENT: RefCell<Vec<(String, String)>> = const { RefCell::new(Vec::new()) };
}

/// Install `vars` as the current instance's environment and return a guard
/// that restores the previous set on drop (including panic unwind).
///
/// Installed variables shadow the process environment; everything the process
/// inherited is still reported, so `get-environment` remains the whole
/// environment rather than just the overlay.
#[must_use = "dropping the guard immediately clears the installed environment"]
pub fn install_environment(vars: Vec<(String, String)>) -> EnvironmentGuard {
    let prev = INSTANCE_ENVIRONMENT.with(|slot| slot.replace(vars));
    EnvironmentGuard { prev }
}

/// Restore-on-drop guard returned by [`install_environment`].
pub struct EnvironmentGuard {
    prev: Vec<(String, String)>,
}

impl Drop for EnvironmentGuard {
    fn drop(&mut self) {
        let prev = std::mem::take(&mut self.prev);
        INSTANCE_ENVIRONMENT.with(|slot| *slot.borrow_mut() = prev);
    }
}

/// The whole environment: what the process inherited, with the instance's own
/// variables layered on top.
fn environment_pairs() -> Vec<(String, String)> {
    let installed = INSTANCE_ENVIRONMENT.with(|slot| slot.borrow().clone());
    let mut pairs: Vec<(String, String)> = std::env::vars()
        .filter(|(key, _)| !installed.iter().any(|(name, _)| name == key))
        .collect();
    pairs.extend(installed);
    pairs
}

pub fn register(vm: &mut VM) {
    // ── wasi:cli/environment — WASI CLI proposal interface ───────────
    // 0.2.x exports `initial-cwd`; 0.3.x renamed it to
    // `get-initial-cwd`. We expose both names on the unversioned
    // interface module so callers targeting either proposal revision
    // can bind the actual CLI environment surface.
    //
    // `get-environment: func() -> list<tuple<string, string>>` — NO argument,
    // in both `wasi-cli/wit/environment.wit` (0.2.0) and
    // `wasi-cli/wit-0.3.0-draft/environment.wit`. A single-key lookup is not
    // part of the interface; a caller that wants one scans the list, which is
    // what the language adapters do.
    vm.register_host_fn(
        "wasi:cli/environment",
        "get-environment",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let pairs: Vec<Value> = environment_pairs()
                .into_iter()
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
