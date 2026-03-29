use egui::Ui;
use crate::state::EditorState;

pub fn show(ui: &mut Ui, state: &mut EditorState) {
    let name = state.current_code_file.clone()
        .or_else(|| state.current_form.clone());

    let Some(name) = name else {
        ui.label("No file selected.");
        return;
    };

    ui.horizontal(|ui| {
        ui.heading(format!("Code — {}", name));
    });
    ui.separator();

    let code = state.get_code_buffer(&name);
    let response = egui::ScrollArea::both()
        .show(ui, |ui| {
            ui.add_sized(
                ui.available_size(),
                egui::TextEdit::multiline(code)
                    .font(egui::TextStyle::Monospace)
                    .desired_width(f32::INFINITY),
            )
        });
}
