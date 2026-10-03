//! Guest callbacks for Webcore's own native browser shell.

use std::cell::RefCell;
use std::rc::Rc;

use tiny_skia::{Color, Pixmap};
use webcore::embedded_window::{self, EmbeddedPage};
use webcore::ui_events::UiEvent;

use vybe_platform_web::engine::{self, DomOp, UiEventFields, WindowOp};

struct Page {
    vm: Rc<RefCell<vybe_runtime::VM>>,
    loaded: bool,
}

impl Page {
    fn dispatch_document_events(&mut self) -> bool {
        let pending = crate::gui_document::drain();
        let changed = !pending.is_empty();
        for dispatch in pending {
            if let Err(error) = self.vm.borrow_mut().invoke(&dispatch.callback, &[dispatch.event]) {
                eprintln!("Event handler error: {error}");
            }
        }
        changed
    }

    fn fire_load(&mut self) {
        let node = engine::DOCUMENT;
        let event = crate::gui_document::event_object("load", node);
        for listener in crate::gui_document::listeners_for(node, "load") {
            if let Err(error) = self.vm.borrow_mut().invoke(&listener, &[event.clone()]) {
                eprintln!("Load handler error: {error}");
            }
        }
    }
}

impl EmbeddedPage for Page {
    fn title(&self) -> String {
        match engine::apply(crate::gui_document::active(), DomOp::Title) {
            engine::DomValue::Text(title) => title,
            _ => String::new(),
        }
    }

    fn resize(&mut self, width: f32, height: f32) {
        engine::window(WindowOp::ViewportChanged(
            crate::gui_document::active(), f64::from(width), f64::from(height),
        ));
        if !self.loaded {
            self.loaded = true;
            self.fire_load();
        }
    }

    fn paint(&mut self, pixmap: &mut Pixmap, scale: f32) {
        pixmap.fill(Color::WHITE);
        vybe_platform_web::present::render(crate::gui_document::active(), pixmap, scale);
    }

    fn input(&mut self, event: &UiEvent) {
        match event.kind.as_str() {
            "mousemove" | "mousedown" | "mouseup" => {
                engine::apply(crate::gui_document::active(), DomOp::DispatchPointer {
                    kind: event.kind.clone(),
                    client_x: event.client_x as f32,
                    client_y: event.client_y as f32,
                    button: event.button,
                });
            }
            "keydown" | "keyup" | "keypress" => {
                engine::apply(crate::gui_document::active(), DomOp::DispatchKeyboard(UiEventFields {
                    kind: event.kind.clone(), key: event.key.clone(), code: event.code.clone(),
                    key_code: event.key_code, ctrl_key: event.ctrl_key,
                    shift_key: event.shift_key, alt_key: event.alt_key, meta_key: event.meta_key,
                    ..UiEventFields::default()
                }));
                if event.kind == "keydown" && !event.ctrl_key && !event.meta_key
                    && event.key.chars().count() == 1 {
                    engine::apply(crate::gui_document::active(), DomOp::DispatchKeyboard(UiEventFields {
                        kind: "keypress".into(), key: event.key.clone(), code: event.code.clone(),
                        key_code: event.key_code, ctrl_key: event.ctrl_key,
                        shift_key: event.shift_key, alt_key: event.alt_key, meta_key: event.meta_key,
                        ..UiEventFields::default()
                    }));
                }
            }
            "wheel" => {
                engine::apply(crate::gui_document::active(), DomOp::DispatchWheel(UiEventFields {
                    kind: event.kind.clone(), client_x: event.client_x,
                    client_y: event.client_y, delta_y: event.delta_y,
                    ctrl_key: event.ctrl_key, shift_key: event.shift_key,
                    alt_key: event.alt_key, meta_key: event.meta_key,
                    ..UiEventFields::default()
                }));
            }
            _ => {}
        }
        self.dispatch_document_events();
    }

    fn tick(&mut self) -> bool {
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
            if let Err(error) = self.vm.borrow_mut().invoke(&callback, &args) {
                eprintln!("Animation callback error: {error}");
            }
        }
        self.dispatch_document_events() || changed
    }
}

pub fn launch_gui(vm: vybe_runtime::VM) {
    let (width, height) = crate::gui_document::viewport().unwrap_or((800, 600));
    embedded_window::run(width, height, Page {
        vm: Rc::new(RefCell::new(vm)), loaded: false,
    });
}
