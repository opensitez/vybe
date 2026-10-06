//! Guest callbacks for Webcore's browser-owned native context.

use std::time::Duration;

use webcore::browser::{Browser, BrowserChrome, BrowserGuest, BrowserOptions};

struct Guest {
    vm: vybe_runtime::VM,
}

impl Guest {
    fn dispatch_document_events(&mut self) -> bool {
        let pending = crate::gui_document::drain();
        let changed = !pending.is_empty();
        for dispatch in pending {
            if let Err(error) = self.vm.invoke(&dispatch.callback, &[dispatch.event]) {
                eprintln!("Event handler error: {error}");
            }
        }
        changed
    }
}

impl BrowserGuest for Guest {
    fn context_closed(&mut self) {
        use vybe_platform_web::engine::{self, WindowOp, WindowValue};
        if let WindowValue::Window(context) = engine::window(WindowOp::DefaultView(crate::gui_document::active())) {
            engine::window(WindowOp::Close(context));
        }
    }
    fn context_ready(&mut self) {
        let node = vybe_platform_web::engine::DOCUMENT;
        let event = crate::gui_document::event_object("load", node);
        for listener in crate::gui_document::listeners_for(node, "load") {
            if let Err(error) = self.vm.invoke(&listener, &[event.clone()]) {
                eprintln!("Load handler error: {error}");
            }
        }
    }

    fn dispatch_pending(&mut self) -> bool {
        use vybe_runtime::scheduler::DeferredSource;
        let clock = vybe_platform_web::animation::callbacks();
        let stamp = vybe_runtime::Value::F64(vybe_runtime::event_loop::monotonic_now_ms());
        let mut changed = false;
        while let Some(callback) = clock.pop_due() {
            changed = true;
            let args = if crate::gui_document::fn_arity(&callback) == 0 {
                Vec::new()
            } else {
                vec![stamp.clone()]
            };
            if let Err(error) = self.vm.invoke(&callback, &args) {
                eprintln!("Animation callback error: {error}");
            }
        }
        let frames: std::sync::Arc<dyn DeferredSource> = clock;
        let callback = self.vm.deferred_sources.iter()
            .filter(|source| !std::sync::Arc::ptr_eq(source, &frames))
            .find_map(|source| source.pop_due());
        if let Some(callback) = callback {
            changed = true;
            if let Err(error) = self.vm.invoke(&callback, &[]) {
                eprintln!("Deferred callback error: {error}");
            }
        }
        self.dispatch_document_events() || changed
    }

    fn next_wakeup(&self) -> Option<Duration> {
        self.vm.deferred_sources.iter()
            .filter(|source| source.has_pending())
            .filter_map(|source| source.earliest_deadline_ms())
            .reduce(f64::min).map(|deadline| {
            Duration::from_secs_f64(
                ((deadline - vybe_runtime::event_loop::monotonic_now_ms()) / 1000.0).max(0.0),
            )
        })
    }
}

pub fn launch_gui(vm: vybe_runtime::VM) {
    let Some(document) =
        vybe_platform_web::engine_webcore::browser_document(crate::gui_document::active())
    else {
        eprintln!("Cannot open Webcore GUI: the active document is closed");
        return;
    };
    let (width, height) = crate::gui_document::viewport().unwrap_or((800, 600));
    let title = document.read().title.clone();
    let options = BrowserOptions {
        title,
        width: width as f32,
        height: height as f32,
        chrome: BrowserChrome::NONE,
        debug: webcore::browser::BrowserDebugOptions {
            port: std::env::var("VYBE_BROWSER_DEBUG_PORT").ok().and_then(|value| value.parse().ok()),
            ..Default::default()
        },
        ..BrowserOptions::default()
    };
    let mut browser = match Browser::open(options) {
        Ok(browser) => browser,
        Err(error) => {
            eprintln!("Cannot open Webcore browser: {error}");
            return;
        }
    };
    browser.attach_document(document);
    browser.set_guest(Guest { vm });
    if let Err(error) = browser.run() {
        eprintln!("Webcore browser error: {error}");
    }
}
