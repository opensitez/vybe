use super::*;

pub fn register(vm: &mut VM) {
    // System.Threading.Tasks.Task — spawn/join compile to Op::THREAD_SPAWN /
    // Op::THREAD_JOIN (WASM threads proposal) via compiler_common::threading.
    // No namespace host fns: `Task.Run(fn)` goes through the compiler's
    // method-call path, not property lookup.
    ensure_namespace(vm, &["System", "Threading", "Tasks", "Task"]);
    ensure_namespace(vm, &["Task"]);

    // System.Diagnostics.Stopwatch — real WASI clocks backing.
    let sw = ensure_namespace(vm, &["System", "Diagnostics", "Stopwatch"]);
    set_prop(&sw, "startnew", host_fn_ref(vm, "wasi:clocks", "stopwatchNew"));
    set_prop(&sw, "new", host_fn_ref(vm, "wasi:clocks", "stopwatchNew"));
    let sw_bare = ensure_namespace(vm, &["Stopwatch"]);
    set_prop(&sw_bare, "startnew", host_fn_ref(vm, "wasi:clocks", "stopwatchNew"));
    set_prop(&sw_bare, "new", host_fn_ref(vm, "wasi:clocks", "stopwatchNew"));

    // System.Diagnostics.Debug / Trace (no-op stubs)
    let debug = ensure_namespace(vm, &["System", "Diagnostics", "Debug"]);
    set_prop(&debug, "writeline", host_fn_ref(vm, "wasi:cli", "log"));
    set_prop(&debug, "write", host_fn_ref(vm, "wasi:cli", "log"));
    set_prop(&debug, "assert", host_fn_ref(vm, "wasi:cli", "log"));
    let debug_bare = ensure_namespace(vm, &["Debug"]);
    set_prop(&debug_bare, "writeline", host_fn_ref(vm, "wasi:cli", "log"));
    set_prop(&debug_bare, "write", host_fn_ref(vm, "wasi:cli", "log"));
    set_prop(&debug_bare, "assert", host_fn_ref(vm, "wasi:cli", "log"));

    let trace = ensure_namespace(vm, &["System", "Diagnostics", "Trace"]);
    set_prop(&trace, "writeline", host_fn_ref(vm, "wasi:cli", "log"));
    let trace_bare = ensure_namespace(vm, &["Trace"]);
    set_prop(&trace_bare, "writeline", host_fn_ref(vm, "wasi:cli", "log"));

    // System.Diagnostics.Process
    let proc_ns = ensure_namespace(vm, &["System", "Diagnostics", "Process"]);
    set_prop(&proc_ns, "start", host_fn_ref(vm, "vybe:types", "processStart"));
    set_prop(&proc_ns, "getcurrentprocess", host_fn_ref(vm, "vybe:types", "processGetCurrent"));

    // System.Random — constructor bound via known_types ctor_mapping in
    // builtin_types.rs. No namespace host fn needed; `new Random()` goes
    // through the compiler's construction path.
    ensure_namespace(vm, &["System", "Random"]);
    ensure_namespace(vm, &["Random"]);

    // System.Threading.Thread — Sleep uses real WASI clocks. Spawn/Start
    // compile to Op::THREAD_SPAWN via compiler_common::threading.
    let thread = ensure_namespace(vm, &["System", "Threading", "Thread"]);
    set_prop(&thread, "sleep", host_fn_ref(vm, "wasi:clocks", "sleep"));
    let thread_bare = ensure_namespace(vm, &["Thread"]);
    set_prop(&thread_bare, "sleep", host_fn_ref(vm, "wasi:clocks", "sleep"));

    // System.Threading.Timer — not yet rewired to wasi:clocks. Placeholder
    // namespace keeps lookups non-null; actual `new Timer(...)` will fail
    // until a WASI-backed timer primitive lands.
    ensure_namespace(vm, &["System", "Threading", "Timer"]);
    ensure_namespace(vm, &["Timer"]);
}
