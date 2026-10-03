//! Launch or capture the document owned by the selected browser engine.

use vybe_platform_web::engine::{self, WindowOp};
use vybe_platform_web::engine_select::{self, Engine};

pub fn launch_gui(vm: vybe_runtime::VM) {
    match engine_select::live() {
        Some(Engine::WebCore) => crate::gui_launch_webcore::launch_gui(vm),
        Some(Engine::OsBrowser) => crate::gui_launch_osbrowser::launch_gui(vm),
        None => eprintln!("Cannot launch GUI: no browser engine is installed"),
    }
}

pub fn capture_gui(
    mut vm: vybe_runtime::VM,
    path: &str,
    control: Option<&str>,
) -> Result<(u32, u32), String> {
    if engine_select::live() != Some(Engine::WebCore) {
        return Err("headless capture requires the in-process WebCore engine".into());
    }

    let (width, height) = crate::gui_document::viewport().unwrap_or((800, 600));
    engine::window(WindowOp::ViewportChanged(
        crate::gui_document::active(),
        f64::from(width),
        f64::from(height),
    ));

    let node = engine::DOCUMENT;
    let event = crate::gui_document::event_object("load", node);
    for listener in crate::gui_document::listeners_for(node, "load") {
        vm.invoke(&listener, &[event.clone()])
            .map_err(|error| format!("Load handler error: {error}"))?;
    }

    crate::gui_capture::capture_to_png(path, control, 1.0)
}
