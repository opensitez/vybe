use egui::Ui;
use crate::state::{EditorState, RunStatus};

pub fn show(ui: &mut Ui, state: &mut EditorState) {
    ui.heading("Output");
    ui.separator();
    match &state.run_status {
        RunStatus::Idle => { ui.label("Not running."); }
        RunStatus::Running => { ui.label("Running..."); }
        RunStatus::Done(msg) => {
            let msg = msg.clone();
            egui::ScrollArea::vertical().show(ui, |ui| {
                ui.add(egui::TextEdit::multiline(&mut msg.as_str())
                    .font(egui::TextStyle::Monospace)
                    .desired_width(f32::INFINITY));
            });
        }
    }
}
