use std::cell::RefCell;
use std::sync::Arc;
use vybe_runtime::value::Object;
use vybe_runtime::vm::HostFnDecl;
use vybe_runtime::{FuncSig, HostContext, VM, ValType, Value};

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

/// Declare a `wasi:cli/environment` function.
///
/// No resource: this interface owns no handles — every function here is a
/// free function returning owned data, so there is nothing to `borrow`.
fn env_fn(
    vm: &mut VM,
    name: &str,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(
        HostFnDecl::new("wasi:cli/environment", name, call).with_sig(FuncSig {
            name: name.to_string(),
            // Every function on this interface is nullary — see the WIT note
            // above `register`. The empty vector is the whole point of
            // declaring them: it is what makes a call that passes an argument
            // report itself.
            params: vec![],
            results,
        }),
    );
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
    env_fn(
        vm,
        "get-environment",
        // `list<tuple<string, string>>`. The Component Model defines a tuple as
        // a record with positional field names, which is exactly what the
        // closure builds: a list of two-element arrays.
        vec![ValType::List(Box::new(ValType::Record(vec![
            ("0".to_string(), ValType::String),
            ("1".to_string(), ValType::String),
        ])))],
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

    env_fn(
        vm,
        "get-arguments",
        // list<string>
        vec![ValType::List(Box::new(ValType::String))],
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let args: Vec<Value> = std::env::args()
                .map(|arg| Value::String(Arc::from(arg.as_str())))
                .collect();
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(args)))
        }),
    );

    // Both spellings of the same 0.2/0.3 function return `option<string>`:
    // the closure answers `Null` when the cwd cannot be read, which is the
    // `none` case, not an error.
    env_fn(
        vm,
        "initial-cwd",
        vec![ValType::Option(Box::new(ValType::String))],
        Box::new(
            |_ctx: &mut HostContext, _args: &[Value]| match std::env::current_dir() {
                Ok(path) => Value::String(Arc::from(path.to_string_lossy().as_ref())),
                Err(_) => Value::Null,
            },
        ),
    );

    env_fn(
        vm,
        "get-initial-cwd",
        vec![ValType::Option(Box::new(ValType::String))],
        Box::new(
            |_ctx: &mut HostContext, _args: &[Value]| match std::env::current_dir() {
                Ok(path) => Value::String(Arc::from(path.to_string_lossy().as_ref())),
                Err(_) => Value::Null,
            },
        ),
    );
}
