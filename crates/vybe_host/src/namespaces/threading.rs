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

    // System.Diagnostics.Stopwatch
    let sw = ensure_namespace(vm, &["System", "Diagnostics", "Stopwatch"]);
    set_prop(&sw, "startnew", host_fn_ref(vm, "vybe:threading", "stopwatchNew"));

    // System.Random
    let rnd = ensure_namespace(vm, &["System", "Random"]);
    set_prop(&rnd, "new", host_fn_ref(vm, "vybe:threading", "randomNew"));

    // Also direct shortcut
    let rnd_short = ensure_namespace(vm, &["Random"]);
    set_prop(&rnd_short, "new", host_fn_ref(vm, "vybe:threading", "randomNew"));
}
