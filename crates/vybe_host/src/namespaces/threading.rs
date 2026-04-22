use super::*;

pub fn register(vm: &mut VM) {
    // System.Threading.Tasks.Task
    let task = ensure_namespace(vm, &["System", "Threading", "Tasks", "Task"]);
    set_prop(&task, "run", host_fn_ref(vm, "vybe:threading", "taskRun"));
    set_prop(&task, "delay", host_fn_ref(vm, "vybe:threading", "taskDelay"));
    set_prop(&task, "fromresult", host_fn_ref(vm, "vybe:threading", "taskFromResult"));
    set_prop(&task, "completedtask", host_fn_ref(vm, "vybe:threading", "taskCompleted"));
    set_prop(&task, "whenall", host_fn_ref(vm, "vybe:threading", "taskCompleted")); // simplified
    set_prop(&task, "whenany", host_fn_ref(vm, "vybe:threading", "taskCompleted")); // simplified

    // Bare Task alias
    let task_bare = ensure_namespace(vm, &["Task"]);
    set_prop(&task_bare, "run", host_fn_ref(vm, "vybe:threading", "taskRun"));
    set_prop(&task_bare, "delay", host_fn_ref(vm, "vybe:threading", "taskDelay"));
    set_prop(&task_bare, "fromresult", host_fn_ref(vm, "vybe:threading", "taskFromResult"));
    set_prop(&task_bare, "completedtask", host_fn_ref(vm, "vybe:threading", "taskCompleted"));
    set_prop(&task_bare, "whenall", host_fn_ref(vm, "vybe:threading", "taskCompleted"));
    set_prop(&task_bare, "whenany", host_fn_ref(vm, "vybe:threading", "taskCompleted"));

    // System.Diagnostics.Stopwatch
    let sw = ensure_namespace(vm, &["System", "Diagnostics", "Stopwatch"]);
    set_prop(&sw, "startnew", host_fn_ref(vm, "wasi:clocks", "stopwatchNew"));
    set_prop(&sw, "new", host_fn_ref(vm, "wasi:clocks", "stopwatchNew"));

    // Bare Stopwatch alias
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

    // System.Random
    let rnd = ensure_namespace(vm, &["System", "Random"]);
    set_prop(&rnd, "new", host_fn_ref(vm, "vybe:threading", "randomNew"));

    // Also direct shortcut
    let rnd_short = ensure_namespace(vm, &["Random"]);
    set_prop(&rnd_short, "new", host_fn_ref(vm, "vybe:threading", "randomNew"));

    // System.Threading.Thread
    let thread = ensure_namespace(vm, &["System", "Threading", "Thread"]);
    set_prop(&thread, "sleep", host_fn_ref(vm, "wasi:clocks", "sleep"));
    let thread_bare = ensure_namespace(vm, &["Thread"]);
    set_prop(&thread_bare, "sleep", host_fn_ref(vm, "wasi:clocks", "sleep"));

    // System.Threading.Timer
    let timer_ns = ensure_namespace(vm, &["System", "Threading", "Timer"]);
    set_prop(&timer_ns, "new", host_fn_ref(vm, "vybe:threading", "taskDelay")); // simplified
    let timer_bare = ensure_namespace(vm, &["Timer"]);
    set_prop(&timer_bare, "new", host_fn_ref(vm, "vybe:threading", "taskDelay")); // simplified
}
