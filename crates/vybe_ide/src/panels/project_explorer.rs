use egui::Ui;
use crate::state::{EditorState, View};

pub fn show(ui: &mut Ui, state: &mut EditorState) {
    ui.heading("Project");
    ui.separator();

    if state.project.is_none() {
        ui.label("No project loaded.");
        return;
    }

    let proj_name = state.project.as_ref().unwrap().name.clone();
    let forms: Vec<String> = state.project.as_ref().unwrap().forms.iter().map(|f| f.form.name.clone()).collect();
    let code_files: Vec<String> = state.project.as_ref().unwrap().code_files.iter().map(|c| c.name.clone()).collect();

    ui.label(format!("📁 {}", proj_name));
    ui.indent("forms", |ui| {
        ui.label("📂 Forms");
        ui.indent("form_list", |ui| {
            for name in &forms {
                let selected = state.current_form.as_deref() == Some(name.as_str());
                if ui.selectable_label(selected, format!("🗒 {}", name)).clicked() {
                    state.current_form = Some(name.clone());
                    state.current_code_file = None;
                    state.view = View::FormDesigner;
                    state.selected_controls.clear();
                }
            }
        });
        ui.label("📂 Code");
        ui.indent("code_list", |ui| {
            for name in &code_files {
                let selected = state.current_code_file.as_deref() == Some(name.as_str());
                if ui.selectable_label(selected, format!("📄 {}", name)).clicked() {
                    state.current_code_file = Some(name.clone());
                    state.current_form = None;
                    state.view = View::CodeEditor;
                }
            }
        });
    });

    ui.separator();
    ui.horizontal(|ui| {
        if ui.small_button("+ Form").clicked() { add_form(state); }
        if ui.small_button("+ Module").clicked() { add_module(state); }
    });
}

fn add_form(state: &mut EditorState) {
    let Some(proj) = state.project.as_mut() else { return };
    let n = proj.forms.len() + 1;
    let name = format!("Form{}", n);
    let mut form = vybe_forms::Form::new(&name);
    form.text = name.clone();
    form.width = 640;
    form.height = 480;
    let designer = vybe_forms::serialization::designer_codegen::generate_designer_code(&form);
    let user = vybe_forms::serialization::designer_codegen::generate_user_code_stub(&name);
    proj.forms.push(vybe_project::FormModule::new_vbnet(form, designer, user));
    state.current_form = Some(name);
    state.view = View::FormDesigner;
}

fn add_module(state: &mut EditorState) {
    let Some(proj) = state.project.as_mut() else { return };
    let n = proj.code_files.len() + 1;
    let name = format!("Module{}", n);
    proj.code_files.push(vybe_project::CodeFile { name: name.clone(), code: format!("Module {}\n\nEnd Module\n", name) });
    state.current_code_file = Some(name);
    state.current_form = None;
    state.view = View::CodeEditor;
}
