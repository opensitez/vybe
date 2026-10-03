//! Guest callbacks while the system browser owns the visible window.

use std::time::Duration;

use vybe_runtime::scheduler::DeferredSource;

pub fn launch_gui(mut vm: vybe_runtime::VM) {
    let node = vybe_platform_web::engine::DOCUMENT;
    let event = crate::gui_document::event_object("load", node);
    vybe_platform_web::engine_osbrowser::begin_batch();
    vybe_platform_web::engine_osbrowser::finish_startup_batch();
    for listener in crate::gui_document::listeners_for(node, "load") {
        if let Err(error) = vm.invoke(&listener, &[event.clone()]) {
            eprintln!("Load handler error: {error}");
        }
    }
    vybe_platform_web::engine_osbrowser::end_batch();

    while vybe_platform_web::engine_osbrowser::connected() {
        let stamp = vybe_runtime::Value::F64(vybe_runtime::event_loop::monotonic_now_ms());
        vybe_platform_web::engine_osbrowser::begin_batch();
        for dispatch in crate::gui_document::drain() {
            if let Err(error) = vm.invoke(&dispatch.callback, &[dispatch.event]) {
                eprintln!("Event handler error: {error}");
            }
        }
        let clock = vybe_platform_web::animation::callbacks();
        if clock.pending_count() > 0 {
            while let Some(callback) = clock.pop_due() {
                let args = if crate::gui_document::fn_arity(&callback) == 0 {
                    Vec::new()
                } else {
                    vec![stamp.clone()]
                };
                if let Err(error) = vm.invoke(&callback, &args) {
                    eprintln!("Animation callback error: {error}");
                }
            }
        }
        vybe_platform_web::engine_osbrowser::end_batch();
        vybe_platform_web::engine_osbrowser::wait_for_event(Duration::from_millis(16));
    }
}
