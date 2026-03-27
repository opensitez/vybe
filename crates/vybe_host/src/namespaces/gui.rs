use super::*;

pub fn register(vm: &mut VM) {
    // vybe.gui.* (low-level API)
    let gui = ensure_namespace(vm, &["vybe", "gui"]);
    set_prop(&gui, "controlsadd", host_fn_ref(vm, "vybe:gui", "controlsAdd"));
    set_prop(&gui, "setproperty", host_fn_ref(vm, "vybe:gui", "setProperty"));
    set_prop(&gui, "onevent", host_fn_ref(vm, "vybe:gui", "onEvent"));
    set_prop(&gui, "msgbox", host_fn_ref(vm, "vybe:gui", "msgBox"));
    set_prop(&gui, "showform", host_fn_ref(vm, "vybe:gui", "showForm"));
    set_prop(&gui, "closeform", host_fn_ref(vm, "vybe:gui", "closeForm"));

    // Application.Run
    let app = ensure_namespace(vm, &["Application"]);
    set_prop(&app, "run", host_fn_ref(vm, "vybe:gui", "runApplication"));
}
