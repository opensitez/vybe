mod state;
mod designer;
mod vb_syntax;

use eframe::egui;
use state::EditorState;
use irys_forms::{ControlType, Form};
use irys_project::Project;
use designer::ResizeHandle;

fn main() -> Result<(), eframe::Error> {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([1200.0, 800.0])
            .with_title("irys Basic Editor"),
        ..Default::default()
    };

    eframe::run_native(
        "irys Editor",
        options,
        Box::new(|_cc| Box::new(irysEditorApp::new())),
    )
}

struct irysEditorApp {
    state: EditorState,
    // Window states
    show_project_explorer: bool,
    show_properties: bool,
    show_toolbox: bool,
    show_code_editor: bool,
    properties_tab_is_events: bool,
    // UI state
    selected_tool: Option<ControlType>,
    dragging_control: Option<(uuid::Uuid, egui::Pos2)>,
    resizing_control: Option<(uuid::Uuid, ResizeHandle)>,
    // Runtime state
    run_mode: bool,
    interpreter: Option<irys_runtime::Interpreter>,
    runtime_project: Option<Project>,
    runtime_active_form: Option<String>,
    // List editing
    editing_list_for_control: Option<uuid::Uuid>,
    list_edit_items: Vec<ListEditorItem>,
    list_edit_loaded_for: Option<uuid::Uuid>,
}

#[derive(Clone, Default)]
struct ListEditorItem {
    text: String,
    value: String,
}

impl irysEditorApp {
    fn new() -> Self {
        // Auto-create default project
        let mut state = EditorState::new();
        let mut project = Project::new("Project1");
        let mut form = Form::new("Form1");
        form.caption = "Form1".to_string();
        project.add_form(form);
        state.project = Some(project);
        state.current_form = Some("Form1".to_string());

        Self {
            state,
            show_project_explorer: true,
            show_properties: true,
            show_toolbox: true,
            show_code_editor: false,
            properties_tab_is_events: false,
            selected_tool: None,
            dragging_control: None,
            resizing_control: None,
            run_mode: false,
            interpreter: None,
            runtime_project: None,
            runtime_active_form: None,
            editing_list_for_control: None,
            list_edit_items: Vec::new(),
            list_edit_loaded_for: None,
        }
    }
}

impl eframe::App for irysEditorApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        // MENU BAR
        egui::TopBottomPanel::top("menu_bar").show(ctx, |ui| {
            egui::menu::bar(ui, |ui| {
                // File Menu
                ui.menu_button("File", |ui| {
                    if ui.button("New Project").clicked() {
                        self.state.project = Some(Project::new("Project1"));
                        self.state.current_project_path = None;
                        ui.close_menu();
                    }
                    if ui.button("Open Project...").clicked() {
                        if let Some(path) = rfd::FileDialog::new()
                            .add_filter("irys Project", &["vbp", "vbproj"])
                            .pick_file() 
                        {
                            match irys_project::load_project_auto(&path) {
                                Ok(project) => {
                                    println!("Loaded project from {:?}", path);
                                    self.state.project = Some(project);
                                    self.state.current_project_path = Some(path);
                                    // Set current form if any
                                    if let Some(proj) = &self.state.project {
                                        if let Some(first) = proj.forms.first() {
                                            self.state.current_form = Some(first.form.name.clone());
                                        }
                                    }
                                }
                                Err(e) => {
                                    eprintln!("Failed to load project: {}", e);
                                }
                            }
                        }
                        ui.close_menu();
                    }
                    if ui.button("Save Project").clicked() {
                        if let Some(project) = &self.state.project {
                            if let Some(path) = &self.state.current_project_path {
                                match irys_project::save_project_auto(project, path) {
                                    Ok(_) => println!("Project saved to {:?}", path),
                                    Err(e) => eprintln!("Failed to save project: {}", e),
                                }
                            } else {
                                // Save As behavior
                                if let Some(path) = rfd::FileDialog::new()
                                    .set_file_name(&format!("{}.vbp", project.name))
                                    .add_filter("irys Project", &["vbp", "vbproj"])
                                    .save_file() 
                                {
                                    match irys_project::save_project_auto(project, &path) {
                                        Ok(_) => {
                                            println!("Project saved to {:?}", path);
                                            self.state.current_project_path = Some(path);
                                        }
                                        Err(e) => eprintln!("Failed to save project: {}", e),
                                    }
                                }
                            }
                        }
                        ui.close_menu();
                    }
                    if ui.button("Save Project As...").clicked() {
                        if let Some(project) = &self.state.project {
                            if let Some(path) = rfd::FileDialog::new()
                                .set_file_name(&format!("{}.vbp", project.name))
                                .add_filter("irys Project", &["vbp", "vbproj"])
                                .save_file() 
                            {
                                match irys_project::save_project_auto(project, &path) {
                                    Ok(_) => {
                                        println!("Project saved to {:?}", path);
                                        self.state.current_project_path = Some(path);
                                    }
                                    Err(e) => eprintln!("Failed to save project: {}", e),
                                }
                            }
                        }
                        ui.close_menu();
                    }
                    ui.separator();
                    if ui.button("Exit").clicked() {
                        ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                    }
                });

                // Edit Menu
                ui.menu_button("Edit", |ui| {
                    if ui.button("Cut").clicked() { ui.close_menu(); }
                    if ui.button("Copy").clicked() { ui.close_menu(); }
                    if ui.button("Paste").clicked() { ui.close_menu(); }
                    ui.separator();
                    if ui.button("Find...").clicked() { ui.close_menu(); }
                    if ui.button("Replace...").clicked() { ui.close_menu(); }
                });

                // View Menu
                ui.menu_button("View", |ui| {
                    ui.checkbox(&mut self.show_project_explorer, "Project Explorer");
                    ui.checkbox(&mut self.show_properties, "Properties Window");
                    ui.checkbox(&mut self.show_toolbox, "Toolbox");
                    ui.separator();
                    if ui.button("Code Editor").clicked() {
                        self.show_code_editor = !self.show_code_editor;
                        ui.close_menu();
                    }
                });

                // Project Menu
                ui.menu_button("Project", |ui| {
                    if ui.button("Add Form").clicked() {
                        if let Some(project) = &mut self.state.project {
                            let form_num = project.forms.len() + 1;
                            let form_name = format!("Form{}", form_num);
                            let mut form = Form::new(&form_name);
                            form.caption = form_name.clone();
                            project.add_form(form);
                        }
                        ui.close_menu();
                    }
                    if ui.button("Project Properties...").clicked() {
                        self.state.show_project_properties = true;
                        ui.close_menu();
                    }
                    // ... other project items
                });

                // Run Menu
                ui.menu_button("Run", |ui| {
                    if ui.button("▶ Start").clicked() {
                        self.run_project();
                        ui.close_menu();
                    }
                    if ui.button("⏹ Stop").clicked() {
                        self.run_mode = false;
                        self.interpreter = None;
                        self.runtime_active_form = None;
                        ui.close_menu();
                    }
                    ui.separator();
                    if ui.button("Restart").clicked() {
                        self.run_mode = false;
                        self.runtime_active_form = None;
                        self.run_project(); // simple restart
                        ui.close_menu();
                    }
                });

                // Window/Help
                ui.menu_button("Window", |ui| {
                    if ui.button("Cascade").clicked() { ui.close_menu(); }
                });
                ui.menu_button("Help", |ui| {
                    if ui.button("About").clicked() { ui.close_menu(); }
                });
            });
        });

        // TOOLBAR
        egui::TopBottomPanel::top("toolbar").show(ctx, |ui| {
            ui.horizontal(|ui| {
                if ui.button("📄 New").clicked() {
                    self.state.project = Some(Project::new("NewProject"));
                }
                if ui.button("📁 Open").clicked() {
                     if let Some(path) = rfd::FileDialog::new()
                        .add_filter("VB Project", &["vbp", "vbproj"])
                        .pick_file() {
                        match irys_project::load_project_auto(&path) {
                            Ok(project) => {
                                self.state.project = Some(project);
                            }
                            Err(e) => {
                                eprintln!("Error loading project: {}", e);
                            }
                        }
                    }
                }
                if ui.button("💾 Save").clicked() {
                    if let Some(project) = &self.state.project {
                        if let Some(path) = &self.state.current_project_path {
                            match irys_project::save_project_auto(project, path) {
                                Ok(_) => println!("Project saved to {:?}", path),
                                Err(e) => eprintln!("Failed to save project: {}", e),
                            }
                        } else {
                            let path = std::path::Path::new(&project.name).with_extension("vbp");
                            match irys_project::save_project_auto(project, &path) {
                                Ok(_) => println!("Project saved to {:?}", path),
                                Err(e) => eprintln!("Failed to save project: {}", e),
                            }
                        }
                    }
                }
                ui.separator();
                
                let start_btn = ui.add_enabled(!self.run_mode, egui::Button::new("▶ Run"));
                if start_btn.clicked() {
                    self.run_project();
                }
                
                let stop_btn = ui.add_enabled(self.run_mode, egui::Button::new("⏹ Stop"));
                if stop_btn.clicked() {
                    self.run_mode = false;
                    self.interpreter = None;
                    self.runtime_project = None;
                    self.runtime_active_form = None;
                }

                ui.separator();
                if ui.button("➕ Add Form").clicked() {
                    if let Some(project) = &mut self.state.project {
                        let form_num = project.forms.len() + 1;
                        let form_name = format!("Form{}", form_num);
                        let mut form = Form::new(&form_name);
                        form.caption = form_name.clone();
                        project.add_form(form);
                    }
                }
            });
        });

        // STATUS BAR
        egui::TopBottomPanel::bottom("status_bar").show(ctx, |ui| {
            ui.horizontal(|ui| {
                if self.run_mode {
                    ui.label("Running...");
                } else {
                    ui.label("Ready");
                }
            });
        });

        if self.run_mode {
            // RUNTIME VIEW
            egui::CentralPanel::default().show(ctx, |ui| {
                self.show_runtime_panel(ui);
            });
        } else {
            // DESIGNER VIEW
            // PROJECT EXPLORER (Left Side)
            if self.show_project_explorer {
                egui::SidePanel::left("project_explorer")
                    .default_width(200.0)
                    .show(ctx, |ui| {
                        ui.heading("Project Explorer");
                        ui.separator();

                        if let Some(project) = &self.state.project {
                            ui.label(format!("📁 {}", project.name));
                            ui.indent("forms", |ui| {
                                ui.label("📋 Forms");
                                for form_module in &project.forms {
                                    let is_selected = self.state.current_form.as_ref()
                                        == Some(&form_module.form.name);
                                    if ui.selectable_label(is_selected, &form_module.form.name).clicked() {
                                        self.state.current_form = Some(form_module.form.name.clone());
                                        self.state.selected_control = None;
                                    }
                                }
                            });
                        }
                    });
            }

            // TOOLBOX (Left Side, below Project Explorer)
            if self.show_toolbox {
                egui::SidePanel::left("toolbox")
                    .default_width(150.0)
                    .show(ctx, |ui| {
                        ui.heading("Toolbox");
                        ui.separator();

                        let controls = [
                            ("Pointer", None),
                            ("Button", Some(ControlType::Button)),
                            ("Label", Some(ControlType::Label)),
                            ("TextBox", Some(ControlType::TextBox)),
                            ("CheckBox", Some(ControlType::CheckBox)),
                            ("RadioButton", Some(ControlType::RadioButton)),
                            ("ComboBox", Some(ControlType::ComboBox)),
                            ("ListBox", Some(ControlType::ListBox)),
                            ("Frame", Some(ControlType::Frame)),
                            ("PictureBox", Some(ControlType::PictureBox)),
                            ("RichTextBox", Some(ControlType::RichTextBox)),
                            ("WebBrowser", Some(ControlType::WebBrowser)),
                        ];

                        for (name, control_type) in controls {
                            let is_selected = self.selected_tool == control_type;
                            if ui.selectable_label(is_selected, name).clicked() {
                                self.selected_tool = control_type;
                            }
                        }
                    });
            }

            // PROPERTIES PANEL (Right Side)
            if self.show_properties {
                egui::SidePanel::right("properties")
                    .default_width(250.0)
                    .show(ctx, |ui| {
                        self.show_properties_panel(ui, ctx);
                    });
            }

            // CENTRAL PANEL - Form Designer
            egui::CentralPanel::default().show(ctx, |ui| {
                if self.show_code_editor {
                    self.show_code_editor_panel(ui);
                } else {
                    self.show_designer_panel(ui);
                }
            });
        }

        if self.state.show_project_properties {
            self.show_project_properties_window(ctx);
        }
        
        self.process_side_effects();
        self.show_msgbox(ctx);
    }
}

impl irysEditorApp {
    fn show_designer_panel(&mut self, ui: &mut egui::Ui) {
        let form_name = match &self.state.current_form {
            Some(name) => name.clone(),
            None => {
                ui.label("No form selected");
                return;
            }
        };

        designer::show_designer(
            ui,
            &mut self.state.project,
            &form_name,
            &mut self.state.selected_control,
            &mut self.selected_tool,
            &mut self.dragging_control,
            &mut self.resizing_control,
            &mut self.show_code_editor,
        );
    }
    fn show_code_editor_panel(&mut self, ui: &mut egui::Ui) {
        ui.horizontal(|ui| {
            ui.heading("Code Editor");
            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                if ui.button("⬅ Back to Designer").clicked() {
                    self.show_code_editor = false;
                }
            });
        });

        // Procedure navigation dropdown (like irys)
        if let Some(code) = self.get_current_code_mut() {
            let procedures = vb_syntax::extract_procedures(code);

            ui.horizontal(|ui| {
                ui.label("Object:");
                egui::ComboBox::from_id_source("object_selector")
                    .selected_text(self.state.current_form.as_ref().unwrap_or(&"Form".to_string()))
                    .width(150.0)
                    .show_ui(ui, |ui| {
                        if let Some(form_name) = &self.state.current_form {
                            ui.selectable_label(true, form_name);
                        }
                    });

                ui.separator();
                ui.label("Procedure:");

                let current_proc = procedures.first().map(|p| format!("{} {}", p.proc_type.icon(), p.name))
                    .unwrap_or_else(|| "(Declarations)".to_string());

                egui::ComboBox::from_id_source("procedure_selector")
                    .selected_text(current_proc)
                    .width(200.0)
                    .show_ui(ui, |ui| {
                        ui.selectable_label(false, "(Declarations)");
                        ui.separator();

                        for proc in &procedures {
                            if ui.selectable_label(false, format!("{} {}", proc.proc_type.icon(), proc.name)).clicked() {
                                // TODO: Scroll to line proc.line
                                // For now, egui_code_editor doesn't support programmatic scrolling easily
                            }
                        }
                    });
            });
        }

        ui.separator();

        // Syntax highlighted code editor
        if let Some(code) = self.get_current_code_mut() {
            let font_id = egui::FontId::monospace(14.0);
            let font_id_clone = font_id.clone();

            let mut layouter = |ui: &egui::Ui, string: &str, _wrap_width: f32| {
                let layout_job = vb_syntax::highlight_irys(string, font_id_clone.clone());
                ui.fonts(|f| f.layout_job(layout_job))
            };

            egui::ScrollArea::vertical().show(ui, |ui| {
                ui.add(
                    egui::TextEdit::multiline(code)
                        .font(font_id)
                        .code_editor()
                        .desired_width(f32::INFINITY)
                        .desired_rows(30)
                        .layouter(&mut layouter)
                );
            });
        }
    }

    fn show_properties_panel(&mut self, ui: &mut egui::Ui, ctx: &egui::Context) {
        ui.heading("Properties");
        
        // Tab switcher
        ui.horizontal(|ui| {
            ui.selectable_value(&mut self.properties_tab_is_events, false, "Properties");
            ui.selectable_value(&mut self.properties_tab_is_events, true, "Events");
        });
        ui.separator();

        let mut delete_control = false;
        let mut open_code_editor = false;
        let mut navigate_to_event: Option<(String, irys_forms::EventType)> = None;
        let mut open_list_editor: Option<uuid::Uuid> = None;
        let is_events_tab = self.properties_tab_is_events;

        if let Some(control_id) = self.state.selected_control {
            if let Some(form) = self.get_current_form_mut() {
                if let Some(control) = form.get_control_mut(control_id) {
                    // Object selector
                    ui.horizontal(|ui| {
                        ui.label(&control.name);
                        ui.label(control.control_type.as_str());
                    });
                    ui.separator();

                    if is_events_tab {
                        // Events tab
                        let control_name = control.name.clone();
                        let control_type = Some(control.control_type);
                        let applicable_events = self.get_applicable_events(control_type);

                        egui::ScrollArea::vertical().show(ui, |ui| {
                            for event in applicable_events {
                                let has_handler = self.find_event_handler(&control_name, &event).is_some();
                                ui.horizontal(|ui| {
                                    if has_handler {
                                        ui.label("✓");
                                    } else {
                                        ui.label("  ");
                                    }
                                    
                                    if ui.button(event.as_str()).clicked() {
                                        navigate_to_event = Some((control_name.clone(), event.clone()));
                                    }
                                });
                            }
                        });
                    } else {
                        // Properties grid (irys style)
                        egui::Grid::new("properties_grid")
                            .num_columns(2)
                            .striped(true)
                            .spacing([40.0, 4.0])
                            .show(ui, |ui| {
                                // Name
                                ui.label("(Name)");
                                let mut name = control.name.clone();
                                if ui.add(egui::TextEdit::singleline(&mut name).desired_width(150.0)).changed() {
                                    control.name = name;
                                }
                                ui.end_row();

                                // Caption
                                if let Some(caption) = control.get_caption() {
                                    ui.label("Caption");
                                    let mut cap = caption.to_string();
                                    if ui.add(egui::TextEdit::singleline(&mut cap).desired_width(150.0)).changed() {
                                        control.set_caption(cap);
                                    }
                                    ui.end_row();
                                }

                                // Text (for TextBox)
                                if let Some(text) = control.get_text() {
                                    ui.label("Text");
                                    let mut txt = text.to_string();
                                    if ui.add(egui::TextEdit::singleline(&mut txt).desired_width(150.0)).changed() {
                                        control.set_text(txt);
                                    }
                                    ui.end_row();
                                }

                                // Enabled
                                ui.label("Enabled");
                                let mut enabled = control.is_enabled();
                                if ui.checkbox(&mut enabled, "").changed() {
                                    control.set_enabled(enabled);
                                }
                                ui.end_row();

                                // Visible
                                ui.label("Visible");
                                let mut visible = control.is_visible();
                                if ui.checkbox(&mut visible, "").changed() {
                                    control.set_visible(visible);
                                }
                                ui.end_row();

                                // Position and Size
                                ui.label("Left");
                                ui.add(egui::DragValue::new(&mut control.bounds.x).speed(1));
                                ui.end_row();

                                ui.label("Top");
                                ui.add(egui::DragValue::new(&mut control.bounds.y).speed(1));
                                ui.end_row();

                                ui.label("Width");
                                ui.add(egui::DragValue::new(&mut control.bounds.width).speed(1).clamp_range(10..=1000));
                                ui.end_row();

                                ui.label("Height");
                                ui.add(egui::DragValue::new(&mut control.bounds.height).speed(1).clamp_range(10..=1000));
                                ui.end_row();

                                // TabIndex
                                ui.label("TabIndex");
                                ui.add(egui::DragValue::new(&mut control.tab_index).speed(1).clamp_range(0..=100));
                                ui.end_row();
                                
                                // List property for ComboBox and ListBox
                                if control.control_type == ControlType::ComboBox || control.control_type == ControlType::ListBox {
                                    ui.label("List");
                                    ui.vertical(|ui| {
                                        // Get current list or create empty one
                                        let current_list = control.properties.get_string_array("List")
                                            .map(|v| v.clone())
                                            .unwrap_or_default();
                                        
                                        // Show items
                                        ui.label(format!("{} items", current_list.len()));
                                        
                                        if ui.button("Edit List...").clicked() {
                                            // For now, show a simple text input for comma-separated values
                                            open_list_editor = Some(control_id);
                                        }
                                    });
                                    ui.end_row();
                                }
                            });
                    }

                    ui.separator();
                    ui.horizontal(|ui| {
                        if ui.button("Delete").clicked() {
                            delete_control = true;
                        }

                        if ui.button("View Events").clicked() {
                            open_code_editor = true;
                        }
                    });
                }
            }

            // Handle actions after releasing borrow
            if delete_control {
                if let Some(form) = self.get_current_form_mut() {
                    form.remove_control(control_id);
                }
                self.state.selected_control = None;
            }
            if let Some(control_id) = open_list_editor {
                self.editing_list_for_control = Some(control_id);
            }
            if open_code_editor {
                self.show_code_editor = true;
            }
        } else {
            // Show form properties
            if let Some(form) = self.get_current_form_mut() {
                ui.horizontal(|ui| {
                    ui.label(&form.name);
                    ui.label("Form");
                });
                ui.separator();

                if is_events_tab {
                    // Events tab for form
                    let form_name = form.name.clone();
                    let applicable_events = self.get_applicable_events(None);

                    egui::ScrollArea::vertical().show(ui, |ui| {
                        for event in applicable_events {
                            let has_handler = self.find_event_handler(&form_name, &event).is_some()
                                || (event == irys_forms::EventType::Load && self.find_event_handler("Form", &event).is_some());
                            ui.horizontal(|ui| {
                                if has_handler {
                                    ui.label("✓");
                                } else {
                                    ui.label("  ");
                                }
                                
                                if ui.button(event.as_str()).clicked() {
                                    navigate_to_event = Some(("Form".to_string(), event.clone()));
                                }
                            });
                        }
                    });
                } else {
                    egui::Grid::new("form_properties_grid")
                        .num_columns(2)
                        .striped(true)
                        .spacing([40.0, 4.0])
                        .show(ui, |ui| {
                            ui.label("(Name)");
                            let mut name = form.name.clone();
                            if ui.add(egui::TextEdit::singleline(&mut name).desired_width(150.0)).changed() {
                                form.name = name;
                            }
                            ui.end_row();

                            ui.label("Caption");
                            ui.text_edit_singleline(&mut form.caption);
                            ui.end_row();

                            ui.label("Width");
                            ui.add(egui::DragValue::new(&mut form.width).speed(10).clamp_range(100..=2000));
                            ui.end_row();

                            ui.label("Height");
                            ui.add(egui::DragValue::new(&mut form.height).speed(10).clamp_range(100..=2000));
                            ui.end_row();
                        });
                }
            } else {
                ui.label("No form selected");
            }
        }
        
        // Handle event navigation after all borrows are released
        if let Some((control_name, event_type)) = navigate_to_event {
            self.navigate_to_or_create_handler(&control_name, event_type);
        }
        
        // List editor dialog
        if let Some(control_id) = self.editing_list_for_control {
            let mut close_dialog = false;
            let mut save_list = false;

            // Load current list if switching controls or first open
            if self.list_edit_loaded_for != Some(control_id) {
                self.list_edit_items.clear();
                if let Some(project) = &self.state.project {
                    if let Some(form_name) = &self.state.current_form {
                        if let Some(form_module) = project.get_form(form_name) {
                            if let Some(control) = form_module.form.get_control(control_id) {
                                let list = control
                                    .properties
                                    .get_string_array("List")
                                    .cloned()
                                    .unwrap_or_default();
                                let values = control
                                    .properties
                                    .get_string_array("ListValues")
                                    .cloned()
                                    .unwrap_or_default();

                                let max_len = list.len().max(values.len());
                                for i in 0..max_len {
                                    let text = list.get(i).cloned().unwrap_or_default();
                                    let value = values.get(i).cloned().unwrap_or_default();
                                    self.list_edit_items.push(ListEditorItem { text, value });
                                }
                            }
                        }
                    }
                }
                self.list_edit_loaded_for = Some(control_id);
            }

            egui::Window::new("Edit List Items")
                .collapsible(false)
                .resizable(true)
                .default_width(520.0)
                .show(ctx, |ui| {
                    ui.label("Edit item text and value:");
                    ui.separator();

                    egui::Grid::new("list_items_grid")
                        .num_columns(3)
                        .spacing([8.0, 4.0])
                        .show(ui, |ui| {
                            ui.label("Item Text");
                            ui.label("Value");
                            ui.label("");
                            ui.end_row();

                            let mut remove_index: Option<usize> = None;
                            for (idx, item) in self.list_edit_items.iter_mut().enumerate() {
                                ui.add(egui::TextEdit::singleline(&mut item.text).desired_width(200.0));
                                ui.add(egui::TextEdit::singleline(&mut item.value).desired_width(180.0));
                                if ui.button("Remove").clicked() {
                                    remove_index = Some(idx);
                                }
                                ui.end_row();
                            }

                            if let Some(index) = remove_index {
                                if index < self.list_edit_items.len() {
                                    self.list_edit_items.remove(index);
                                }
                            }
                        });

                    ui.separator();
                    ui.horizontal(|ui| {
                        if ui.button("Add Item").clicked() {
                            self.list_edit_items.push(ListEditorItem::default());
                        }
                        ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                            if ui.button("OK").clicked() {
                                save_list = true;
                                close_dialog = true;
                            }
                            if ui.button("Cancel").clicked() {
                                close_dialog = true;
                            }
                        });
                    });
                });

            if save_list {
                let mut items: Vec<String> = Vec::new();
                let mut values: Vec<String> = Vec::new();
                for item in &self.list_edit_items {
                    let text = item.text.trim().to_string();
                    if !text.is_empty() {
                        items.push(text);
                        values.push(item.value.trim().to_string());
                    }
                }

                if let Some(project) = &mut self.state.project {
                    if let Some(form_name) = &self.state.current_form {
                        if let Some(form_module) = project.get_form_mut(form_name) {
                            if let Some(control) = form_module.form.get_control_mut(control_id) {
                                use irys_forms::properties::PropertyValue;
                                control
                                    .properties
                                    .set_raw("List", PropertyValue::StringArray(items));
                                control
                                    .properties
                                    .set_raw("ListValues", PropertyValue::StringArray(values));
                            }
                        }
                    }
                }
            }

            if close_dialog {
                self.editing_list_for_control = None;
                self.list_edit_items.clear();
                self.list_edit_loaded_for = None;
            }
        }
    }

    fn get_current_form_mut(&mut self) -> Option<&mut Form> {
        let form_name = self.state.current_form.clone()?;
        self.state.project.as_mut()?
            .get_form_mut(&form_name)
            .map(|fm| &mut fm.form)
    }

    fn get_current_code_mut(&mut self) -> Option<&mut String> {
        let form_name = self.state.current_form.clone()?;
        self.state.project.as_mut()?
            .get_form_mut(&form_name)
            .map(|fm| &mut fm.code)
    }

    fn show_runtime_panel(&mut self, ui: &mut egui::Ui) {
        let mut pending_list_selection: Vec<(uuid::Uuid, String, i32, String, String)> = Vec::new();
        let mut pending_checkbox_updates: Vec<(uuid::Uuid, bool)> = Vec::new();
        let mut pending_events: Vec<(String, irys_forms::EventType)> = Vec::new();
        let mut pending_env_updates: Vec<(String, irys_runtime::Value)> = Vec::new();
        let mut pending_url_updates: Vec<(uuid::Uuid, String)> = Vec::new();
        let mut pending_text_updates: Vec<(uuid::Uuid, String)> = Vec::new();

        if let Some(project) = &self.runtime_project {
             // Show the active form (set by Form.Show or initially by startup_form)
             if let Some(form_name) = &self.runtime_active_form {
                 if let Some(form_mod) = project.get_form(form_name) {
                     let form = &form_mod.form;
                     
                     // Form background
                     let rect = ui.max_rect();
                     ui.painter().rect_filled(rect, 0.0, egui::Color32::from_rgb(240, 240, 240));

                     // Detect clicks on form background
                     let form_response = ui.interact(rect, ui.id().with("form_background"), egui::Sense::click().union(egui::Sense::drag()));
                     
                     // Click events
                     if form_response.clicked() {
                         println!("Runtime: Form clicked");
                         pending_events.push(("Form".to_string(), irys_forms::EventType::Click));
                     }
                     if form_response.double_clicked() {
                         println!("Runtime: Form double-clicked");
                         pending_events.push(("Form".to_string(), irys_forms::EventType::DblClick));
                     }
                     
                     // Mouse events
                     if form_response.drag_started() {
                         pending_events.push(("Form".to_string(), irys_forms::EventType::MouseDown));
                     }
                     if form_response.drag_stopped() {
                         pending_events.push(("Form".to_string(), irys_forms::EventType::MouseUp));
                     }
                     if form_response.dragged() {
                         pending_events.push(("Form".to_string(), irys_forms::EventType::MouseMove));
                     }
                     
                     // Keyboard events (when form has focus)
                     if ui.input(|i| i.key_pressed(egui::Key::Enter)) {
                         pending_events.push(("Form".to_string(), irys_forms::EventType::KeyPress));
                     }

                     let offset = rect.min;

                     // Draw controls
                     // Note: We need to handle interactivity if it's a button.
                     
                     // Ideally we iterate in z-order.
                     for control in &form.controls {
                         let pos = offset + egui::vec2(control.bounds.x as f32, control.bounds.y as f32);
                         let size = egui::vec2(control.bounds.width as f32, control.bounds.height as f32);
                         let control_rect = egui::Rect::from_min_size(pos, size);

                         if !control.is_visible() { continue; }

                         match control.control_type {
                             ControlType::Button => {
                                 let text = control.get_caption().unwrap_or("");
                                 let enabled = control.is_enabled();
                                 let btn = egui::Button::new(text).sense(egui::Sense::click().union(egui::Sense::drag()));
                                 
                                 // We need to place it at exact coords
                                 let btn_response = ui.put(control_rect, btn);
                                 
                                 if enabled && btn_response.clicked() {
                                     println!("Runtime: Button {} clicked", control.name);
                                     pending_events.push((control.name.clone(), irys_forms::EventType::Click));
                                 }
                                 
                                 if enabled && btn_response.double_clicked() {
                                     println!("Runtime: Button {} double-clicked", control.name);
                                     pending_events.push((control.name.clone(), irys_forms::EventType::DblClick));
                                 }
                                 
                                 // Mouse events
                                 if enabled && btn_response.drag_started() {
                                     pending_events.push((control.name.clone(), irys_forms::EventType::MouseDown));
                                 }
                                 if enabled && btn_response.drag_stopped() {
                                     pending_events.push((control.name.clone(), irys_forms::EventType::MouseUp));
                                 }
                                 if enabled && btn_response.hovered() {
                                     pending_events.push((control.name.clone(), irys_forms::EventType::MouseMove));
                                 }
                             }
                             ControlType::Label => {
                                 let text = control.get_caption().unwrap_or("");
                                 ui.put(control_rect, egui::Label::new(text));
                             }
                             ControlType::TextBox => {
                                 // Get current text from runtime project
                                 let mut text = control.get_text().unwrap_or("").to_string();
                                 let text_edit = egui::TextEdit::singleline(&mut text);
                                 let response = ui.put(control_rect, text_edit);

                                 // Trigger Change event if text was modified
                                 if response.changed() {
                                     println!("Runtime: TextBox {} changed", control.name);
                                     pending_text_updates.push((control.id, text.clone()));
                                     pending_events.push((control.name.clone(), irys_forms::EventType::Change));
                                     pending_env_updates.push((
                                         format!("{}.Text", control.name),
                                         irys_runtime::Value::String(text.clone()),
                                     ));
                                 }
                                 
                                 // Focus events
                                 if response.gained_focus() {
                                     pending_events.push((control.name.clone(), irys_forms::EventType::GotFocus));
                                 }
                                 if response.lost_focus() {
                                     pending_events.push((control.name.clone(), irys_forms::EventType::LostFocus));
                                 }
                             }
                             ControlType::CheckBox => {
                                 let text = control.get_caption().unwrap_or("");
                                let current_value = control
                                    .properties
                                    .get_int("Value")
                                    .unwrap_or_else(|| {
                                        if control.properties.get_bool("Value").unwrap_or(false) { 1 } else { 0 }
                                    });
                                let mut checked = current_value != 0;
                                let response = ui.put(control_rect, egui::Checkbox::new(&mut checked, text));
                                if response.clicked() || response.changed() {
                                    let new_value = if checked { 1 } else { 0 };
                                    let changed = new_value != current_value;
                                    pending_checkbox_updates.push((control.id, checked));
                                    pending_events.push((control.name.clone(), irys_forms::EventType::Click));
                                    if changed {
                                        pending_events.push((control.name.clone(), irys_forms::EventType::Change));
                                    }
                                    pending_env_updates.push((
                                        format!("{}.Value", control.name),
                                        irys_runtime::Value::Integer(new_value),
                                    ));
                                    pending_env_updates.push((
                                        format!("{}.Text", control.name),
                                        irys_runtime::Value::String(text.to_string()),
                                    ));
                                }
                             }
                             ControlType::RadioButton => {
                                 let text = control.get_caption().unwrap_or("");
                                let selected = control
                                    .properties
                                    .get_bool("Value")
                                    .unwrap_or_else(|| control.properties.get_int("Value").unwrap_or(0) != 0);
                                 ui.put(control_rect, egui::RadioButton::new(selected, text));
                             }
                             ControlType::ComboBox => {
                                 let items = control.properties.get_string_array("List");
                                 let list_values = control.properties.get_string_array("ListValues");
                                 let selected_index = control.properties.get_int("ListIndex").unwrap_or(-1);
                                 let display_text = items
                                     .and_then(|list| list.get(selected_index.max(0) as usize))
                                     .map(|s| s.as_str())
                                     .or_else(|| control.properties.get_string("Text"))
                                     .unwrap_or("(Select)");

                                let mut child_ui = ui.child_ui(control_rect, *ui.layout());
                                egui::ComboBox::from_id_source(control.id)
                                    .selected_text(display_text)
                                    .width(control_rect.width())
                                    .show_ui(&mut child_ui, |ui| {
                                        if let Some(list) = items {
                                            for (index, item) in list.iter().enumerate() {
                                                let is_selected = selected_index == index as i32;
                                                if ui.selectable_label(is_selected, item).clicked() {
                                                    let new_index = index as i32;
                                                    let new_text = item.clone();
                                                    let new_value = list_values
                                                        .and_then(|v| v.get(index))
                                                        .cloned()
                                                        .unwrap_or_default();

                                                    pending_list_selection.push((
                                                        control.id,
                                                        control.name.clone(),
                                                        new_index,
                                                        new_text.clone(),
                                                        new_value.clone(),
                                                    ));

                                                    if new_index != selected_index {
                                                        pending_events.push((
                                                            control.name.clone(),
                                                            irys_forms::EventType::Change,
                                                        ));
                                                    }

                                                    pending_env_updates.push((
                                                        format!("{}.ListIndex", control.name),
                                                        irys_runtime::Value::Integer(new_index),
                                                    ));
                                                    pending_env_updates.push((
                                                        format!("{}.Text", control.name),
                                                        irys_runtime::Value::String(new_text),
                                                    ));
                                                    pending_env_updates.push((
                                                        format!("{}.Value", control.name),
                                                        irys_runtime::Value::String(new_value),
                                                    ));
                                                }
                                            }
                                        }
                                    });
                             }
                             ControlType::ListBox => {
                                 // Draw bordered area and show items
                                 ui.painter().rect_filled(control_rect, 0.0, egui::Color32::WHITE);
                                 ui.painter().rect_stroke(control_rect, 0.0, egui::Stroke::new(1.0, egui::Color32::GRAY));
                                 
                                 // Display list items
                                 if let Some(items) = control.properties.get_string_array("List") {
                                     let selected_index = control.properties.get_int("ListIndex").unwrap_or(-1);
                                     let row_height = 16.0;
                                     let max_rows = ((control_rect.height() - 6.0) / row_height)
                                         .floor()
                                         .max(1.0) as usize;
                                     let list_values = control.properties.get_string_array("ListValues");

                                     for (index, item) in items.iter().take(max_rows).enumerate() {
                                         let y_offset = 3.0 + (index as f32 * row_height);
                                         let row_rect = egui::Rect::from_min_size(
                                             control_rect.min + egui::vec2(2.0, y_offset),
                                             egui::vec2(control_rect.width() - 4.0, row_height),
                                         );

                                         let is_selected = selected_index == index as i32;
                                         if is_selected {
                                             ui.painter().rect_filled(
                                                 row_rect,
                                                 0.0,
                                                 egui::Color32::from_rgb(0, 120, 215),
                                             );
                                         }

                                         let text_pos = row_rect.min + egui::vec2(3.0, 0.0);
                                         ui.painter().text(
                                             text_pos,
                                             egui::Align2::LEFT_TOP,
                                             item,
                                             egui::FontId::default(),
                                             if is_selected { egui::Color32::WHITE } else { egui::Color32::BLACK },
                                         );

                                         let response = ui.interact(
                                             row_rect,
                                             ui.id().with((control.id, index)),
                                             egui::Sense::click(),
                                         );
                                         if response.clicked() {
                                             let new_index = index as i32;
                                             let new_text = item.clone();
                                             let new_value = list_values
                                                 .and_then(|v| v.get(index))
                                                 .cloned()
                                                 .unwrap_or_default();

                                             pending_list_selection.push((
                                                 control.id,
                                                 control.name.clone(),
                                                 new_index,
                                                 new_text.clone(),
                                                 new_value.clone(),
                                             ));

                                             pending_events.push((control.name.clone(), irys_forms::EventType::Click));
                                             if new_index != selected_index {
                                                 pending_events.push((control.name.clone(), irys_forms::EventType::Change));
                                             }

                                             pending_env_updates.push((
                                                 format!("{}.ListIndex", control.name),
                                                 irys_runtime::Value::Integer(new_index),
                                             ));
                                             pending_env_updates.push((
                                                 format!("{}.Text", control.name),
                                                 irys_runtime::Value::String(new_text),
                                             ));
                                             pending_env_updates.push((
                                                 format!("{}.Value", control.name),
                                                 irys_runtime::Value::String(new_value),
                                             ));
                                         }
                                     }
                                 }
                             }
                             ControlType::Frame => {
                                 // Draw a frame/group box
                                 let caption = control.get_caption().unwrap_or("Frame");
                                 ui.painter().rect_stroke(control_rect, 2.0, egui::Stroke::new(1.0, egui::Color32::GRAY));
                                 // Draw caption at top
                                 let text_pos = control_rect.min + egui::vec2(10.0, -7.0);
                                 ui.painter().text(
                                     text_pos,
                                     egui::Align2::LEFT_TOP,
                                     caption,
                                     egui::FontId::default(),
                                     egui::Color32::BLACK,
                                 );
                             }
                             ControlType::PictureBox => {
                                 // Draw a bordered area for pictures
                                 ui.painter().rect_filled(control_rect, 0.0, egui::Color32::from_rgb(240, 240, 240));
                                 ui.painter().rect_stroke(control_rect, 0.0, egui::Stroke::new(1.0, egui::Color32::DARK_GRAY));
                                 ui.put(control_rect, egui::Label::new("[Picture]"));
                             }
                             ControlType::RichTextBox => {
                                 // Get current text from runtime project
                                 let mut text = control.get_text().unwrap_or("").to_string();
                                 let text_edit = egui::TextEdit::multiline(&mut text)
                                     .desired_width(control_rect.width())
                                     .desired_rows((control_rect.height() / 20.0).max(1.0) as usize);
                                 let response = ui.put(control_rect, text_edit);

                                 // Trigger Change event if text was modified
                                 if response.changed() {
                                     println!("Runtime: RichTextBox {} changed", control.name);
                                     pending_text_updates.push((control.id, text.clone()));
                                     pending_events.push((control.name.clone(), irys_forms::EventType::Change));
                                     pending_env_updates.push((
                                         format!("{}.Text", control.name),
                                         irys_runtime::Value::String(text.clone()),
                                     ));
                                 }

                                 // Focus events
                                 if response.gained_focus() {
                                     pending_events.push((control.name.clone(), irys_forms::EventType::GotFocus));
                                 }
                                 if response.lost_focus() {
                                     pending_events.push((control.name.clone(), irys_forms::EventType::LostFocus));
                                 }
                             }
                             ControlType::WebBrowser => {
                                 // Check if URL was changed by VB code and schedule update
                                 if let Some(interp) = &self.interpreter {
                                     let url_key = format!("{}.URL", control.name);
                                     if let Ok(env_value) = interp.env.get(&url_key) {
                                         let env_url = env_value.as_string();
                                         let current_url = control.properties.get_string("URL").unwrap_or("about:blank");
                                         if env_url != current_url {
                                             pending_url_updates.push((control.id, env_url.to_string()));
                                         }
                                     }
                                 }

                                 // Draw a bordered area with URL display
                                 ui.painter().rect_filled(control_rect, 0.0, egui::Color32::WHITE);
                                 ui.painter().rect_stroke(control_rect, 0.0, egui::Stroke::new(1.0, egui::Color32::DARK_GRAY));

                                 let url = control.properties.get_string("URL").unwrap_or("about:blank");
                                 let label_text = format!("[Browser: {}]", url);
                                 ui.put(control_rect, egui::Label::new(label_text));
                             }
                         }
                     }
                 }
             }
        }

        if !pending_list_selection.is_empty() {
            if let Some(project) = &mut self.runtime_project {
                if let Some(form_name) = project.startup_form.clone().or_else(|| self.state.current_form.clone()) {
                    if let Some(form_mod) = project.get_form_mut(&form_name) {
                        for (control_id, _control_name, new_index, new_text, new_value) in &pending_list_selection {
                            if let Some(control) = form_mod.form.get_control_mut(*control_id) {
                                use irys_forms::properties::PropertyValue;
                                control
                                    .properties
                                    .set_raw("ListIndex", PropertyValue::Integer(*new_index));
                                control
                                    .properties
                                    .set_raw("Text", PropertyValue::String(new_text.clone()));
                                control
                                    .properties
                                    .set_raw("Value", PropertyValue::String(new_value.clone()));
                            }
                        }
                    }
                }
            }
        }

        if !pending_checkbox_updates.is_empty() {
            if let Some(project) = &mut self.runtime_project {
                if let Some(form_name) = project.startup_form.clone().or_else(|| self.state.current_form.clone()) {
                    if let Some(form_mod) = project.get_form_mut(&form_name) {
                        for (control_id, checked) in &pending_checkbox_updates {
                            if let Some(control) = form_mod.form.get_control_mut(*control_id) {
                                use irys_forms::properties::PropertyValue;
                                control.properties.set_raw(
                                    "Value",
                                    PropertyValue::Integer(if *checked { 1 } else { 0 }),
                                );
                            }
                        }
                    }
                }
            }
        }

        if !pending_url_updates.is_empty() {
            if let Some(project) = &mut self.runtime_project {
                if let Some(form_name) = &self.runtime_active_form.clone() {
                    if let Some(form_mod) = project.get_form_mut(form_name) {
                        for (control_id, new_url) in &pending_url_updates {
                            if let Some(control) = form_mod.form.get_control_mut(*control_id) {
                                use irys_forms::properties::PropertyValue;
                                control.properties.set_raw("URL", PropertyValue::String(new_url.clone()));
                            }
                        }
                    }
                }
            }
        }

        if !pending_text_updates.is_empty() {
            if let Some(project) = &mut self.runtime_project {
                if let Some(form_name) = &self.runtime_active_form.clone() {
                    if let Some(form_mod) = project.get_form_mut(form_name) {
                        for (control_id, new_text) in &pending_text_updates {
                            if let Some(control) = form_mod.form.get_control_mut(*control_id) {
                                use irys_forms::properties::PropertyValue;
                                control.properties.set_raw("Text", PropertyValue::String(new_text.clone()));
                            }
                        }
                    }
                }
            }
        }

        if let Some(interp) = &mut self.interpreter {
            for (key, value) in pending_env_updates {
                let _ = interp.env.set(&key, value);
            }
            for (control_name, event_type) in pending_events {
                if let Err(e) = interp.trigger_event(&control_name, event_type, None) {
                    println!("Runtime Event Error: {}", e);
                }
            }
        }
    }

    fn run_project(&mut self) {
        if let Some(master_project) = &self.state.project {
            let project = master_project.clone();
            self.runtime_project = Some(project.clone());
            println!("\n=== Starting Runtime: {} ===", project.name);

            // Create Interpreter
            let mut interpreter = irys_runtime::Interpreter::new();
            
            // Parse all forms code
            for form_module in &project.forms {
                let code_to_run = if form_module.code.trim().is_empty() {
                    "" // Allow empty code
                } else {
                    &form_module.code
                };
                
                match irys_parser::parse_program(code_to_run) {
                    Ok(program) => {
                        // 1. Load module with proper scoping
                        if let Err(e) = interpreter.load_module(&form_module.form.name, &program) {
                             eprintln!("Runtime Error during init of {}: {}", form_module.form.name, e);
                        }

                        // 2. Register explicit event handlers
                        for binding in &form_module.form.event_bindings {
                            interpreter.events.register(
                                &binding.control_name,
                                binding.event_type.clone(),
                                &binding.handler_name,
                            );
                        }

                        // 3. Auto-register event handlers based on naming convention
                        // Check for ControlName_EventName patterns
                        for (sub_name, _) in &interpreter.subs {
                            // Extract procedure name (handle qualified names like "form1.btn1_click")
                            let proc_name = if let Some(dot_pos) = sub_name.rfind('.') {
                                &sub_name[dot_pos + 1..]
                            } else {
                                sub_name.as_str()
                            };

                            if let Some(underscore_pos) = proc_name.rfind('_') {
                                let control_name = &proc_name[..underscore_pos];
                                let event_name = &proc_name[underscore_pos + 1..];

                                // Try to match event name to EventType
                                for event_type in irys_forms::EventType::all_events() {
                                    if event_type.as_str().to_lowercase() == event_name {
                                        interpreter.events.register(
                                            control_name,
                                            event_type,
                                            sub_name,  // Register with full qualified name
                                        );
                                        break;
                                    }
                                }
                            }
                        }

                        // 4. Special case: support "Form_Load" convention
                        let form_load_key = format!("{}.form_load", form_module.form.name.to_lowercase());
                        if interpreter.subs.contains_key(&form_load_key) {
                            interpreter.events.register(
                                &form_module.form.name,
                                irys_forms::EventType::Load,
                                &form_load_key,
                            );
                        } else if interpreter.subs.contains_key("form_load") {
                            interpreter.events.register(
                                &form_module.form.name,
                                irys_forms::EventType::Load,
                                "form_load",
                            );
                        }

                        // 5. Populate environment with current control properties
                        for control in &form_module.form.controls {
                            // Caption
                            if let Some(caption) = control.get_caption() {
                                interpreter.env.define(format!("{}.Caption", control.name), irys_runtime::Value::String(caption.to_string()));
                            }
                            // Text
                            if let Some(text) = control.get_text() {
                                interpreter.env.define(format!("{}.Text", control.name), irys_runtime::Value::String(text.to_string()));
                            }
                            // Enabled
                            interpreter.env.define(format!("{}.Enabled", control.name), irys_runtime::Value::Boolean(control.is_enabled()));
                            // Visible
                            interpreter.env.define(format!("{}.Visible", control.name), irys_runtime::Value::Boolean(control.is_visible()));

                            // CheckBox/RadioButton value (irys uses 1/0)
                            if matches!(control.control_type, ControlType::CheckBox | ControlType::RadioButton) {
                                let checked = control.properties.get_bool("Value").unwrap_or(false);
                                interpreter.env.define(
                                    format!("{}.Value", control.name),
                                    irys_runtime::Value::Integer(if checked { 1 } else { 0 }),
                                );
                            }

                            // ListBox/ComboBox selection
                            if matches!(control.control_type, ControlType::ComboBox | ControlType::ListBox) {
                                let list_index = control.properties.get_int("ListIndex").unwrap_or(-1);
                                interpreter.env.define(
                                    format!("{}.ListIndex", control.name),
                                    irys_runtime::Value::Integer(list_index),
                                );

                                let text = control.properties.get_string("Text").unwrap_or("");
                                interpreter.env.define(
                                    format!("{}.Text", control.name),
                                    irys_runtime::Value::String(text.to_string()),
                                );

                                let value = control.properties.get_string("Value").unwrap_or("");
                                interpreter.env.define(
                                    format!("{}.Value", control.name),
                                    irys_runtime::Value::String(value.to_string()),
                                );
                            }

                            // WebBrowser URL
                            if matches!(control.control_type, ControlType::WebBrowser) {
                                let url = control.properties.get_string("URL").unwrap_or("about:blank");
                                interpreter.env.define(
                                    format!("{}.URL", control.name),
                                    irys_runtime::Value::String(url.to_string()),
                                );
                            }
                        }

                        // 6. Trigger Form_Load
                        if let Err(e) = interpreter.trigger_event(&form_module.form.name, irys_forms::EventType::Load, None) {
                             eprintln!("Runtime Error during Load of {}: {}", form_module.form.name, e);
                        }
                        
                        // Form properties
                        interpreter.env.define(format!("{}.Caption", form_module.form.name), irys_runtime::Value::String(form_module.form.caption.clone()));
                    }
                    Err(e) => {
                        eprintln!("Parse error in {}: {}", form_module.form.name, e);
                    }
                }
            }
            
            self.interpreter = Some(interpreter);
            // Set initial active form to startup form or first form
            self.runtime_active_form = project.startup_form.clone()
                .or_else(|| project.forms.first().map(|f| f.form.name.clone()));
            self.run_mode = true;
            println!("Runtime Mode Active");
        }
    }

    fn show_project_properties_window(&mut self, ctx: &egui::Context) {
        let mut open = self.state.show_project_properties;
        let mut close_requested = false;
        egui::Window::new("Project Properties")
            .open(&mut open)
            .collapsible(false)
            .resizable(false)
            .show(ctx, |ui| {
                if let Some(project) = &mut self.state.project {
                    ui.vertical(|ui| {
                        ui.horizontal(|ui| {
                            ui.label("Project Name:");
                            ui.text_edit_singleline(&mut project.name);
                        });

                        ui.add_space(8.0);

                        ui.horizontal(|ui| {
                            ui.label("Startup Object:");
                            let mut selected = project.startup_form.clone().unwrap_or_default();
                            egui::ComboBox::from_id_source("startup_obj")
                                .selected_text(if selected.is_empty() { "(None)" } else { &selected })
                                .show_ui(ui, |ui| {
                                    ui.selectable_value(&mut selected, "".to_string(), "(None)");
                                    for form_mod in &project.forms {
                                        ui.selectable_value(&mut selected, form_mod.form.name.clone(), &form_mod.form.name);
                                    }
                                });
                            if selected.is_empty() {
                                project.startup_form = None;
                            } else {
                                project.startup_form = Some(selected);
                            }
                        });

                        ui.add_space(16.0);
                        ui.horizontal(|ui| {
                            if ui.button("OK").clicked() {
                                close_requested = true;
                            }
                        });
                    });
                } else {
                    ui.label("No project loaded.");
                    if ui.button("Close").clicked() {
                        close_requested = true;
                    }
                }
            });
        
        if close_requested {
            open = false;
        }
        self.state.show_project_properties = open;
    }

    fn process_side_effects(&mut self) {
        let Some(interp) = &mut self.interpreter else { return };
        while let Some(effect) = interp.side_effects.pop_front() {
            match effect {
                irys_runtime::RuntimeSideEffect::MsgBox(msg) => {
                    self.state.msgbox_pending = Some(msg);
                }
                irys_runtime::RuntimeSideEffect::PropertyChange { object, property, value } => {
                    if let Some(project) = &mut self.runtime_project {
                        for form_mod in &mut project.forms {
                            if let Some(control) = form_mod.form.controls.iter_mut().find(|c| c.name.to_lowercase() == object.to_lowercase()) {
                                match property.to_lowercase().as_str() {
                                    "caption" => control.set_caption(value.as_string()),
                                    "text" => control.set_text(value.as_string()),
                                    "enabled" => control.set_enabled(value.as_bool().unwrap_or(true)),
                                    "visible" => control.set_visible(value.as_bool().unwrap_or(true)),
                                    "listindex" => {
                                        use irys_forms::properties::PropertyValue;
                                        let idx = value.as_integer().unwrap_or(-1);
                                        control.properties.set_raw("ListIndex", PropertyValue::Integer(idx));
                                    }
                                    "value" => {
                                        use irys_forms::properties::PropertyValue;
                                        match &value {
                                            irys_runtime::Value::Integer(i) => {
                                                control.properties.set_raw("Value", PropertyValue::Integer(*i));
                                            }
                                            irys_runtime::Value::Long(i) => {
                                                control.properties.set_raw("Value", PropertyValue::Integer(*i as i32));
                                            }
                                            irys_runtime::Value::Boolean(b) => {
                                                control.properties.set_raw(
                                                    "Value",
                                                    PropertyValue::Integer(if *b { 1 } else { 0 }),
                                                );
                                            }
                                            irys_runtime::Value::String(s) => {
                                                control
                                                    .properties
                                                    .set_raw("Value", PropertyValue::String(s.clone()));
                                            }
                                            _ => {
                                                control
                                                    .properties
                                                    .set_raw("Value", PropertyValue::String(value.as_string()));
                                            }
                                        }
                                    }
                                    _ => {}
                                }
                            } else if form_mod.form.name.to_lowercase() == object.to_lowercase() {
                                // Property on the form itself
                                match property.to_lowercase().as_str() {
                                    "caption" => form_mod.form.caption = value.as_string(),
                                    "visible" => {
                                        // When a form's Visible property is set to true (Show), make it the active form
                                        if value.as_bool().unwrap_or(false) {
                                            self.runtime_active_form = Some(form_mod.form.name.clone());
                                            println!("Switching to form: {}", form_mod.form.name);
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    fn show_msgbox(&mut self, ctx: &egui::Context) {
        let Some(msg) = self.state.msgbox_pending.clone() else { return };
        
        let mut open = true;
        let mut close_requested = false;
        egui::Window::new("irys Basic")
            .open(&mut open)
            .collapsible(false)
            .resizable(false)
            .anchor(egui::Align2::CENTER_CENTER, [0.0, 0.0])
            .show(ctx, |ui| {
                ui.vertical_centered(|ui| {
                    ui.label(&msg);
                    ui.add_space(12.0);
                    if ui.button("OK").clicked() {
                        close_requested = true;
                    }
                });
            });
            
        if !open || close_requested {
            self.state.msgbox_pending = None;
        }
    }

    fn find_event_handler(&self, control_name: &str, event_type: &irys_forms::EventType) -> Option<usize> {
        let code = self.get_current_code()?;
        let handler_name = format!("{}_{}", control_name, event_type.as_str()).to_lowercase();
        
        for (line_num, line) in code.lines().enumerate() {
            let line_lower = line.to_lowercase();
            if line_lower.contains("sub") && line_lower.contains(&handler_name) {
                return Some(line_num);
            }
        }
        None
    }

    fn get_current_code(&self) -> Option<&String> {
        let form_name = self.state.current_form.as_ref()?;
        self.state.project.as_ref()?
            .get_form(form_name)
            .map(|fm| &fm.code)
    }

    fn get_applicable_events(&self, control_type: Option<ControlType>) -> Vec<irys_forms::EventType> {
        irys_forms::EventType::all_events()
            .into_iter()
            .filter(|event| event.is_applicable_to(control_type))
            .collect()
    }

    fn generate_event_stub(&self, control_name: &str, event_type: &irys_forms::EventType) -> String {
        let handler_name = format!("{}_{}", control_name, event_type.as_str());
        let params = event_type.parameters();
        
        if params.is_empty() {
            format!("Private Sub {}()\n    ' TODO: Implement {}\nEnd Sub\n\n", handler_name, event_type.as_str())
        } else {
            format!("Private Sub {}({})\n    ' TODO: Implement {}\nEnd Sub\n\n", handler_name, params, event_type.as_str())
        }
    }

    fn navigate_to_or_create_handler(&mut self, control_name: &str, event_type: irys_forms::EventType) {
        if let Some(_line_num) = self.find_event_handler(control_name, &event_type) {
            // Handler exists, just open code editor
            // TODO: Implement scrolling to specific line
            self.show_code_editor = true;
        } else {
            // Generate and add stub
            let stub = self.generate_event_stub(control_name, &event_type);
            if let Some(code) = self.get_current_code_mut() {
                code.push_str(&stub);
            }
            self.show_code_editor = true;
        }
    }
}

