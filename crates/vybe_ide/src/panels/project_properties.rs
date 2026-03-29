use eframe::egui;
use crate::state::EditorState;

pub fn show_dialog(ctx: &egui::Context, state: &mut EditorState) {
    if !state.show_project_properties {
        return;
    }
    
    let mut is_open = state.show_project_properties;
    let mut close_clicked = false;

    egui::Window::new("Project Properties")
        .open(&mut is_open)
        .collapsible(false)
        .resizable(false)
        .default_width(320.0)
        .show(ctx, |ui| {
            if let Some(proj) = state.project.as_mut() {
                egui::Grid::new("proj_props_grid").num_columns(2).spacing([16.0, 12.0]).striped(true).show(ui, |ui| {
                    ui.label(egui::RichText::new("Project Name").strong());
                    ui.add(egui::TextEdit::singleline(&mut proj.name).desired_width(200.0));
                    ui.end_row();
                    
                    ui.label(egui::RichText::new("Startup Object").strong());
                    
                    let current_startup = match &proj.startup_object {
                        vybe_project::StartupObject::Form(n) => n.clone(),
                        vybe_project::StartupObject::SubMain => "Sub Main".to_string(),
                        vybe_project::StartupObject::None => "(None)".to_string(),
                    };
                    
                    egui::ComboBox::from_id_salt("startup_cb")
                        .selected_text(current_startup)
                        .width(200.0)
                        .show_ui(ui, |ui| {
                            let mut changed = false;
                            if ui.selectable_label(matches!(proj.startup_object, vybe_project::StartupObject::SubMain), "Sub Main").clicked() {
                                proj.startup_object = vybe_project::StartupObject::SubMain;
                                proj.startup_form = None;
                                changed = true;
                            }
                            if ui.selectable_label(matches!(proj.startup_object, vybe_project::StartupObject::None), "(None)").clicked() {
                                proj.startup_object = vybe_project::StartupObject::None;
                                proj.startup_form = None;
                                changed = true;
                            }
                            // List ALL forms
                            let form_names: Vec<String> = proj.forms.iter().map(|f| f.form.name.clone()).collect();
                            for name in form_names {
                                let is_sel = matches!(&proj.startup_object, vybe_project::StartupObject::Form(n) if n == &name);
                                if ui.selectable_label(is_sel, &name).clicked() {
                                    proj.startup_object = vybe_project::StartupObject::Form(name.clone());
                                    // Required for backward compat macro inside vybe_project:
                                    proj.startup_form = Some(name.clone());
                                    changed = true;
                                }
                            }
                            if changed {
                                ui.ctx().request_repaint();
                            }
                        });
                    ui.end_row();
                });
                
                ui.add_space(8.0);
                ui.label(egui::RichText::new("More settings coming soon...").weak().small());
                ui.add_space(16.0);
                
                ui.horizontal(|ui| {
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        if ui.button("OK").clicked() {
                            close_clicked = true;
                        }
                    });
                });
            } else {
                ui.label("No project loaded.");
                if ui.button("Close").clicked() {
                    close_clicked = true;
                }
            }
        });
        
    if close_clicked { is_open = false; }
    state.show_project_properties = is_open;
}
