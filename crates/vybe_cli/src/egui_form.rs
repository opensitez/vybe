use eframe::egui;
use std::rc::Rc;
use std::cell::RefCell;

pub fn launch_egui_form(
    mut vm: vybe_bytecode::VM,
    queue: Rc<RefCell<vybe_host::SideEffectQueue>>,
    form: &vybe_forms::Form,
    title: &str,
) {
    let app = EguiApp::new(vm, queue, form.clone(), title.to_string());
    let native_options = eframe::NativeOptions::default();
    eframe::run_native(
        title,
        native_options,
        Box::new(|_cc| Box::new(app)),
    );
}

struct EguiApp {
    vm: RefCell<vybe_bytecode::VM>,
    queue: Rc<RefCell<vybe_host::SideEffectQueue>>,
    form: vybe_forms::Form,
    title: String,
}

impl EguiApp {
    fn new(
        vm: vybe_bytecode::VM,
        queue: Rc<RefCell<vybe_host::SideEffectQueue>>,
        form: vybe_forms::Form,
        title: String,
    ) -> Self {
        Self { vm: RefCell::new(vm), queue, form, title }
    }

    fn handle_event(&self, control_name: &str, event_name: &str) {
        let callback = {
            let q = self.queue.borrow();
            q.get_event_handler(control_name, event_name).cloned()
        };
        if let Some(cb) = callback {
            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let mut vm = self.vm.borrow_mut();
                let me = vm.globals.get("__f").cloned()
                    .or_else(|| vm.globals.get("me").cloned())
                    .unwrap_or(vybe_bytecode::Value::Null);
                let arity = match &cb {
                    vybe_bytecode::Value::Object(obj) => {
                        match &obj.borrow().kind {
                            vybe_bytecode::value::ObjectKind::Function(f) => f.arity as usize,
                            _ => 0,
                        }
                    }
                    _ => 0,
                };
                let sender = vybe_bytecode::Value::String(std::rc::Rc::from(control_name));
                match arity {
                    0 => vm.invoke(&cb, &[]),
                    1 => vm.invoke(&cb, &[me]),
                    2 => vm.invoke(&cb, &[me, sender]),
                    _ => vm.invoke(&cb, &[me, sender, vybe_bytecode::Value::Null]),
                }
            }));
            if let Err(panic) = result {
                let _ = panic; // swallow panic; VM will have logged
            }
            // Process side effects produced by callback
            let new_effects = self.queue.borrow_mut().drain();
            for effect in new_effects {
                match effect {
                    vybe_host::SideEffect::ConsoleOutput(msg) => print!("{msg}"),
                    vybe_host::SideEffect::MsgBox { text, title } => println!("[MsgBox] {}: {}", title, text),
                    _ => {}
                }
            }
        }
    }
}

impl eframe::App for EguiApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default().show(ctx, |ui| {
            let form_w = self.form.width as f32;
            let form_h = self.form.height as f32;
            let painter = ui.painter();

            // Draw simple background
            let rect = egui::Rect::from_min_size(egui::pos2(0.0, 0.0), egui::vec2(form_w, form_h));
            painter.rect_filled(rect, 0.0, egui::Color32::from_gray(240));

            for ctrl in &self.form.controls {
                let x = ctrl.bounds.x as f32;
                let y = ctrl.bounds.y as f32;
                let w = ctrl.bounds.width as f32;
                let h = ctrl.bounds.height as f32;
                let r = egui::Rect::from_min_size(egui::pos2(x, y), egui::vec2(w.max(1.0), h.max(1.0)));
                let response = ui.interact(r, ui.make_persistent_id(&ctrl.name), egui::Sense::click());
                // draw control background
                painter.rect_filled(r, 4.0, egui::Color32::WHITE);
                // draw text
                let text = ctrl.properties.get_string("Text").unwrap_or_default();
                painter.text(
                    egui::pos2(x + 4.0, y + 4.0),
                    egui::Align2::LEFT_TOP,
                    text,
                    egui::FontId::proportional(14.0),
                    egui::Color32::BLACK,
                );
                if response.clicked() {
                    self.handle_event(&ctrl.name, "Click");
                }
            }
        });
    }
}
