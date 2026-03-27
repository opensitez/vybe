use super::*;

pub fn register(vm: &mut VM) {
    // --- DateTime ---
    let dt = ensure_namespace(vm, &["DateTime"]);
    set_prop(&dt, "now", host_fn_ref(vm, "vybe:types", "dateTimeNow"));
    set_prop(&dt, "parse", host_fn_ref(vm, "vybe:types", "dateTimeParse"));

    let sys_dt = ensure_namespace(vm, &["System", "DateTime"]);
    set_prop(&sys_dt, "now", host_fn_ref(vm, "vybe:types", "dateTimeNow"));
    set_prop(&sys_dt, "parse", host_fn_ref(vm, "vybe:types", "dateTimeParse"));
    set_prop(&sys_dt, "today", host_fn_ref(vm, "vybe:types", "dateTimeNow")); // simplified

    // --- StringBuilder ---
    let sb = ensure_namespace(vm, &["StringBuilder"]);
    set_prop(&sb, "new", host_fn_ref(vm, "vybe:types", "stringBuilderNew")); // not really used this way

    let sys_sb = ensure_namespace(vm, &["System", "Text", "StringBuilder"]);
    set_prop(&sys_sb, "new", host_fn_ref(vm, "vybe:types", "stringBuilderNew"));

    // --- Process ---
    let proc = ensure_namespace(vm, &["Process"]);
    set_prop(&proc, "start", host_fn_ref(vm, "vybe:types", "processStart"));

    let sys_proc = ensure_namespace(vm, &["System", "Diagnostics", "Process"]);
    set_prop(&sys_proc, "start", host_fn_ref(vm, "vybe:types", "processStart"));
}
