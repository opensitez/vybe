// App state management using Dioxus signals
use dioxus::prelude::*;
use irys_forms::{ControlType, Form};
use irys_project::Project;
use rfd::FileDialog;
use std::path::PathBuf;
use uuid::Uuid;

// Helper function to strip HTML tags for plain text
fn strip_html_tags(html: &str) -> String {
    let mut result = String::new();
    let mut in_tag = false;
    
    for c in html.chars() {
        match c {
            '<' => in_tag = true,
            '>' => in_tag = false,
            _ if !in_tag => result.push(c),
            _ => {}
        }
    }
    
    // Decode common HTML entities
    result
        .replace("&nbsp;", " ")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
        .replace("&quot;", "\"")
        .replace("&#39;", "'")
}

#[derive(Clone, Copy)]
pub struct AppState {
    pub project: Signal<Option<Project>>,
    pub current_form: Signal<Option<String>>,
    pub current_project_path: Signal<Option<PathBuf>>,
    pub selected_control: Signal<Option<Uuid>>,
    pub selected_tool: Signal<Option<ControlType>>,
    pub run_mode: Signal<bool>,
    pub show_code_editor: Signal<bool>,
    pub show_properties: Signal<bool>,
    pub show_toolbox: Signal<bool>,
    pub show_project_explorer: Signal<bool>,
    pub show_project_properties: Signal<bool>,
    pub show_resources: Signal<bool>,
    pub clipboard_control: Signal<Option<irys_forms::Control>>,
}

impl AppState {
    pub fn new() -> Self {
        Self {
            project: Signal::new(None),
            current_form: Signal::new(None),
            current_project_path: Signal::new(None),
            selected_control: Signal::new(None),
            selected_tool: Signal::new(None),
            run_mode: Signal::new(false),
            show_code_editor: Signal::new(false),
            show_properties: Signal::new(true),
            show_toolbox: Signal::new(true),
            show_project_explorer: Signal::new(true),
            show_project_properties: Signal::new(false),
            show_resources: Signal::new(false),
            clipboard_control: Signal::new(None),
        }
    }
    
    pub fn get_current_form(&self) -> Option<Form> {
        let project = self.project.read();
        let form_name = self.current_form.read();
        
        if let (Some(proj), Some(name)) = (project.as_ref(), form_name.as_ref()) {
            proj.get_form(name).map(|fm| fm.form.clone())
        } else {
            None
        }
    }

    pub fn new_project(&self) {
        let mut project_signal = self.project;
        let mut project_write = project_signal.write();
        let mut current_form_signal = self.current_form;
        let mut current_form_write = current_form_signal.write();
        let mut path_signal = self.current_project_path;
        let mut path_write = path_signal.write();

        let mut project = Project::new("Project1");
        let mut form = Form::new("Form1");
        form.caption = "Form1".to_string();
        project.add_form(form);

        *project_write = Some(project);
        *current_form_write = Some("Form1".to_string());
        *path_write = None;
    }

    pub fn open_project_dialog(&self) {
        if let Some(path) = FileDialog::new()
            .add_filter("Irys Project", &["vbp", "vbproj"])
            .pick_file()
        {
            eprintln!("[DEBUG] open_project_dialog: path={:?}", path);
            match irys_project::load_project_auto(&path) {
                Ok(project) => {
                    eprintln!("[DEBUG] Project loaded: '{}' with {} forms", project.name, project.forms.len());
                    for f in &project.forms {
                        eprintln!("[DEBUG]   Form '{}': {} controls, caption='{}'", f.form.name, f.form.controls.len(), f.form.caption);
                        for c in &f.form.controls {
                            eprintln!("[DEBUG]     Control: {} ({:?}) at ({},{}) {}x{}", c.name, c.control_type, c.bounds.x, c.bounds.y, c.bounds.width, c.bounds.height);
                        }
                    }
                    let mut project_signal = self.project;
                    let mut project_write = project_signal.write();
                    let mut current_form_signal = self.current_form;
                    let mut current_form_write = current_form_signal.write();
                    let mut path_signal = self.current_project_path;
                    let mut path_write = path_signal.write();

                    *project_write = Some(project);
                    if let Some(proj) = project_write.as_ref() {
                        if let Some(first) = proj.forms.first() {
                            eprintln!("[DEBUG] Setting current form to: '{}'", first.form.name);
                            *current_form_write = Some(first.form.name.clone());
                        } else {
                            *current_form_write = None;
                        }
                    }
                    *path_write = Some(path);
                }
                Err(e) => {
                    eprintln!("Failed to load project: {}", e);
                }
            }
        }
    }

    pub fn save_project(&self) {
        let project_read = self.project.read();
        let Some(project) = project_read.as_ref() else { return };
        let current_path = self.current_project_path.read().clone();

        if let Some(path) = current_path {
            if let Err(e) = irys_project::save_project_auto(project, &path) {
                eprintln!("Failed to save project: {}", e);
            }
        } else {
            self.save_project_as();
        }
    }

    pub fn save_project_as(&self) {
        let project_read = self.project.read();
        let Some(project) = project_read.as_ref() else { return };

        if let Some(path) = FileDialog::new()
            .set_file_name(&format!("{}.vbp", project.name))
            .add_filter("Irys Project", &["vbp", "vbproj"])
            .save_file()
        {
            if let Err(e) = irys_project::save_project_auto(project, &path) {
                eprintln!("Failed to save project: {}", e);
                return;
            }
            let mut path_signal = self.current_project_path;
            let mut path_write = path_signal.write();
            *path_write = Some(path);
        }
    }

    pub fn get_current_code(&self) -> String {
        let project = self.project.read();
        let form_name = self.current_form.read();

        if let (Some(proj), Some(name)) = (project.as_ref(), form_name.as_ref()) {
            // Try form first
            if let Some(fm) = proj.get_form(name) {
                return fm.get_user_code().to_string();
            }
            // Try code file
            if let Some(cf) = proj.get_code_file(name) {
                return cf.code.clone();
            }
        }
        String::new()
    }

    pub fn get_current_designer_code(&self) -> String {
        let project = self.project.read();
        let form_name = self.current_form.read();

        if let (Some(proj), Some(name)) = (project.as_ref(), form_name.as_ref()) {
            if let Some(fm) = proj.get_form(name) {
                return fm.get_designer_code().to_string();
            }
        }
        String::new()
    }

    pub fn is_current_form_vbnet(&self) -> bool {
        let project = self.project.read();
        let form_name = self.current_form.read();

        if let (Some(proj), Some(name)) = (project.as_ref(), form_name.as_ref()) {
            if let Some(fm) = proj.get_form(name) {
                return fm.is_vbnet();
            }
        }
        false
    }

    pub fn update_current_code(&self, new_code: String) {
        let mut project_signal = self.project;
        let mut project_write = project_signal.write();
        let form_name = self.current_form.read();

        if let (Some(proj), Some(name)) = (project_write.as_mut(), form_name.as_ref()) {
            // Try form first
            if let Some(form_module) = proj.get_form_mut(name) {
                form_module.set_user_code(new_code);
                return;
            }
            // Try code file
            if let Some(cf) = proj.get_code_file_mut(name) {
                cf.code = new_code;
            }
        }
    }



    pub fn update_control_property(&self, control_id: uuid::Uuid, property: &str, value: String) {
        let mut project_signal = self.project;
        let mut project_write = project_signal.write();
        let form_name = self.current_form.read();

        if let (Some(proj), Some(name)) = (project_write.as_mut(), form_name.as_ref()) {
            if let Some(form_module) = proj.get_form_mut(name) {
                // Validate Name property before taking mutable borrow
                if property == "Name" {
                    let trimmed = value.trim().to_string();
                    if trimmed.is_empty() {
                        return;
                    }
                    let has_duplicate = form_module.form.controls.iter().any(|c| {
                        c.id != control_id && c.name.eq_ignore_ascii_case(&trimmed)
                    });
                    if has_duplicate {
                        return;
                    }
                }

                // Handle Index property before taking mutable control borrow
                if property == "Index" {
                    let trimmed = value.trim();
                    if trimmed.is_empty() {
                        // Remove from array
                        if let Some(control) = form_module.form.get_control_mut(control_id) {
                            control.index = None;
                        }
                        form_module.sync_designer_code();
                        return;
                    }
                    if let Ok(new_idx) = trimmed.parse::<i32>() {
                        // Check for duplicate name+index
                        let ctrl_name = form_module.form.get_control(control_id)
                            .map(|c| c.name.clone())
                            .unwrap_or_default();
                        let has_dup = form_module.form.controls.iter().any(|c| {
                            c.id != control_id
                                && c.name.eq_ignore_ascii_case(&ctrl_name)
                                && c.index == Some(new_idx)
                        });
                        if has_dup {
                            return; // Duplicate index
                        }
                        // If setting index on a non-array control, auto-set index=0 on
                        // another control with the same name that has no index
                        let other_no_index: Option<uuid::Uuid> = form_module.form.controls.iter()
                            .find(|c| {
                                c.id != control_id
                                    && c.name.eq_ignore_ascii_case(&ctrl_name)
                                    && c.index.is_none()
                            })
                            .map(|c| c.id);
                        if let Some(other_id) = other_no_index {
                            if let Some(other) = form_module.form.get_control_mut(other_id) {
                                other.index = Some(0);
                            }
                        }
                        if let Some(control) = form_module.form.get_control_mut(control_id) {
                            control.index = Some(new_idx);
                        }
                        form_module.sync_designer_code();
                        return;
                    }
                    return; // Invalid input
                }

                if let Some(control) = form_module.form.get_control_mut(control_id) {
                    match property {
                        "Name" => {
                            control.name = value.trim().to_string();
                        }
                        "Caption" => control.set_caption(value),
                        "Text" => control.set_text(value),
                        "BackColor" => control.set_back_color(value),
                        "ForeColor" => control.set_fore_color(value),
                        "Font" => control.set_font(value),
                        "Enabled" => {
                            if let Ok(enabled) = value.parse::<bool>() {
                                control.set_enabled(enabled);
                            }
                        },
                        "Visible" => {
                            if let Ok(visible) = value.parse::<bool>() {
                                control.set_visible(visible);
                            }
                        },
                        "TabIndex" => {
                            if let Ok(tab_index) = value.parse::<i32>() {
                                control.tab_index = tab_index;
                            }
                        },
                        "List" => {
                            let items: Vec<String> = value
                                .split('\n')
                                .map(|s| s.trim().to_string())
                                .filter(|s| !s.is_empty())
                                .collect();
                            control.set_list_items(items);
                        },
                        "Value" => {
                            if let Ok(val) = value.parse::<i32>() {
                                use irys_forms::properties::PropertyValue;
                                control.properties.set_raw("Value", PropertyValue::Integer(val));
                            }
                        },
                        "URL" => {
                            control.properties.set("URL", value);
                        },
                        "HTML" => {
                            control.properties.set("HTML", value.clone());
                            // Also update Text property for RichTextBox to keep them in sync
                            // Strip HTML tags for plain text version
                            let plain_text = strip_html_tags(&value);
                            control.set_text(plain_text);
                        },
                        "ToolbarVisible" => {
                            if let Ok(visible) = value.parse::<bool>() {
                                control.properties.set("ToolbarVisible", visible);
                            }
                        },
                        // Data binding properties
                        "DataSource" | "DataMember" | "Filter" | "Sort" |
                        "BindingSource" | "DataSetName" | "TableName" |
                        "SelectCommand" | "ConnectionString" |
                        "DisplayMember" | "ValueMember" |
                        "DbType" | "DbPath" | "DbHost" | "DbPort" |
                        "DbName" | "DbUser" | "DbPassword" => {
                            control.properties.set(property, value);
                        },
                        // Simple data bindings (DataBindings.Text, DataBindings.Checked, etc.)
                        prop if prop.starts_with("DataBindings.") => {
                            control.properties.set(prop, value);
                        },
                        _ => {}
                    }
                }
                form_module.sync_designer_code();
            }
        }
    }

    pub fn update_control_geometry(&self, control_id: uuid::Uuid, x: i32, y: i32, width: i32, height: i32) {
        let mut project_signal = self.project;
        let mut project_write = project_signal.write();
        let form_name = self.current_form.read();

        if let (Some(proj), Some(name)) = (project_write.as_mut(), form_name.as_ref()) {
            if let Some(form_module) = proj.get_form_mut(name) {
                if let Some(control) = form_module.form.get_control_mut(control_id) {
                    control.bounds.x = x;
                    control.bounds.y = y;
                    control.bounds.width = width;
                    control.bounds.height = height;
                }
                form_module.sync_designer_code();
            }
        }
    }

    pub fn add_control_at(&self, control_type: ControlType, x: i32, y: i32) {
        let mut project_signal = self.project;
        let mut project_write = project_signal.write();
        let form_name = self.current_form.read();

        if let (Some(proj), Some(name)) = (project_write.as_mut(), form_name.as_ref()) {
            if let Some(form_module) = proj.get_form_mut(name) {
                let prefix = control_type.default_name_prefix();
                let mut counter = 1;
                let mut control_name = format!("{}{}", prefix, counter);

                while form_module.form.get_control_by_name(&control_name).is_some() {
                    counter += 1;
                    control_name = format!("{}{}", prefix, counter);
                }

                let new_control = irys_forms::Control::new(
                    control_type,
                    control_name,
                    x,
                    y
                );

                form_module.form.add_control(new_control);
                form_module.sync_designer_code();
            }
        }
    }

    pub fn add_new_form(&self) {
        let mut project_signal = self.project;
        let mut project_write = project_signal.write();
        
        if let Some(proj) = project_write.as_mut() {
            let mut counter = 1;
            let mut name = format!("Form{}", counter);
            while proj.get_form(&name).is_some() {
                counter += 1;
                name = format!("Form{}", counter);
            }
            
            let mut form = Form::new(&name);
            form.caption = name.clone();
            form.width = 400;
            form.height = 300;
            
            proj.add_form(form);
            
            // Switch to new form
            let mut form_signal = self.current_form;
            *form_signal.write() = Some(name);
        }
    }

    pub fn update_form_property(&self, property: &str, value: String) {
        let mut project_signal = self.project;
        let mut project_write = project_signal.write();
        let form_name = self.current_form.read();

        if let (Some(proj), Some(name)) = (project_write.as_mut(), form_name.as_ref()) {
            if let Some(form_module) = proj.get_form_mut(name) {
                match property {
                    "Caption" => form_module.form.caption = value,
                    "Width" => {
                        if let Ok(w) = value.parse::<i32>() { form_module.form.width = w; }
                    }
                    "Height" => {
                        if let Ok(h) = value.parse::<i32>() { form_module.form.height = h; }
                    }
                    "BackColor" => form_module.form.back_color = Some(value),
                    "ForeColor" => form_module.form.fore_color = Some(value),
                    "Font" => form_module.form.font = Some(value),
                    _ => {}
                }
                form_module.sync_designer_code();
            }
        }
    }

    pub fn add_new_vbnet_form(&self) {
        let mut project_signal = self.project;
        let mut project_write = project_signal.write();

        if let Some(proj) = project_write.as_mut() {
            let mut counter = 1;
            let mut name = format!("Form{}", counter);
            while proj.get_form(&name).is_some() {
                counter += 1;
                name = format!("Form{}", counter);
            }

            let mut form = Form::new(&name);
            form.caption = name.clone();
            form.width = 640;
            form.height = 480;

            let designer_code = irys_forms::serialization::designer_codegen::generate_designer_code(&form);
            let user_code = irys_forms::serialization::designer_codegen::generate_user_code_stub(&name);

            let form_module = irys_project::FormModule::new_vbnet(form, designer_code, user_code);
            proj.forms.push(form_module);

            // Switch to new form
            let mut form_signal = self.current_form;
            *form_signal.write() = Some(name);
        }
    }

    pub fn add_code_file(&self) {
        let mut project_signal = self.project;
        let mut project_write = project_signal.write();

        if let Some(proj) = project_write.as_mut() {
            let mut counter = 1;
            let mut name = format!("Code{}", counter);
            while proj.get_code_file(&name).is_some() {
                counter += 1;
                name = format!("Code{}", counter);
            }

            let mut code_file = irys_project::CodeFile::new(&name);
            code_file.code = format!("' Code file: {}\n\n", name);

            proj.add_code_file(code_file);

            // Switch to new code file
            let mut form_signal = self.current_form;
            *form_signal.write() = Some(name);

            // Switch to code view
            let mut code_editor_signal = self.show_code_editor;
            *code_editor_signal.write() = true;
        }
    }

    pub fn delete_selected_control(&self) {
        let control_id = match *self.selected_control.read() {
            Some(id) => id,
            None => return,
        };

        let mut project_signal = self.project;
        let mut project_write = project_signal.write();
        let form_name = self.current_form.read();

        if let (Some(proj), Some(name)) = (project_write.as_mut(), form_name.as_ref()) {
            if let Some(form_module) = proj.get_form_mut(name) {
                // Collect IDs of the control and all its descendants
                let mut to_remove = vec![control_id];
                let mut i = 0;
                while i < to_remove.len() {
                    let parent = to_remove[i];
                    for c in &form_module.form.controls {
                        if c.parent_id == Some(parent) && !to_remove.contains(&c.id) {
                            to_remove.push(c.id);
                        }
                    }
                    i += 1;
                }

                form_module.form.controls.retain(|c| !to_remove.contains(&c.id));
                form_module.sync_designer_code();
            }
        }

        let mut selected = self.selected_control;
        selected.set(None);
    }

    pub fn copy_selected_control(&self) {
        let control_id = match *self.selected_control.read() {
            Some(id) => id,
            None => return,
        };

        let project = self.project.read();
        let form_name = self.current_form.read();

        if let (Some(proj), Some(name)) = (project.as_ref(), form_name.as_ref()) {
            if let Some(fm) = proj.get_form(name) {
                if let Some(control) = fm.form.get_control(control_id) {
                    let mut clipboard = self.clipboard_control;
                    clipboard.set(Some(control.clone()));
                }
            }
        }
    }

    pub fn cut_selected_control(&self) {
        self.copy_selected_control();
        self.delete_selected_control();
    }

    pub fn paste_control(&self) {
        let clipboard_control = match self.clipboard_control.read().clone() {
            Some(c) => c,
            None => return,
        };

        let mut project_signal = self.project;
        let mut project_write = project_signal.write();
        let form_name = self.current_form.read();

        if let (Some(proj), Some(name)) = (project_write.as_mut(), form_name.as_ref()) {
            if let Some(form_module) = proj.get_form_mut(name) {
                let mut new_control = clipboard_control.clone();
                new_control.id = Uuid::new_v4();
                new_control.parent_id = None;
                new_control.bounds.x += 20;
                new_control.bounds.y += 20;

                if clipboard_control.is_array_member() {
                    // Array-aware paste: keep same name, assign next array index
                    let next_idx = form_module.form.next_array_index(&clipboard_control.name);
                    new_control.name = clipboard_control.name.clone();
                    new_control.index = Some(next_idx);
                } else {
                    // Generate unique name
                    let prefix = new_control.control_type.default_name_prefix();
                    let mut counter = 1;
                    let mut new_name = format!("{}{}", prefix, counter);
                    while form_module.form.get_control_by_name(&new_name).is_some() {
                        counter += 1;
                        new_name = format!("{}{}", prefix, counter);
                    }
                    new_control.name = new_name;
                }

                // Update clipboard with offset position for next paste
                let mut clipboard = self.clipboard_control;
                let mut updated = clipboard_control.clone();
                updated.bounds.x += 20;
                updated.bounds.y += 20;
                clipboard.set(Some(updated));

                let new_id = new_control.id;
                form_module.form.add_control(new_control);
                form_module.sync_designer_code();

                let mut selected = self.selected_control;
                selected.set(Some(new_id));
            }
        }
    }

    pub fn reparent_control(&self, control_id: uuid::Uuid, new_parent_id: Option<uuid::Uuid>) {
        if Some(control_id) == new_parent_id {
            return; // Cannot reparent to self
        }

        let mut project_signal = self.project;
        let mut project_write = project_signal.write();
        let form_name = self.current_form.read();
        
        if let (Some(proj), Some(name)) = (project_write.as_mut(), form_name.as_ref()) {
            if let Some(form_module) = proj.get_form_mut(name) {
                // 1. Build a lookup for geometry calculations and hierarchy check
                let mut control_map = std::collections::HashMap::new();
                for c in &form_module.form.controls {
                    control_map.insert(c.id, (c.bounds.clone(), c.parent_id));
                }

                // Check for cycles: walk up from new_parent_id. If we hit control_id, abort.
                let mut ancestor = new_parent_id;
                while let Some(a_id) = ancestor {
                     if a_id == control_id {
                         return; // Cycle detected: trying to drop parent into child
                     }
                     if let Some((_, pid)) = control_map.get(&a_id) {
                         ancestor = *pid;
                     } else {
                         break;
                     }
                }

                // Helper to get global pos from map
                let get_global_pos = |target_id: Option<Uuid>, map: &std::collections::HashMap<Uuid, (irys_forms::Bounds, Option<Uuid>)>| -> (i32, i32) {
                    let mut x = 0;
                    let mut y = 0;
                    let mut curr = target_id;
                    while let Some(cid) = curr {
                        if let Some((bounds, pid)) = map.get(&cid) {
                            x += bounds.x;
                            y += bounds.y;
                            curr = *pid;
                        } else {
                            break;
                        }
                    }
                    (x, y)
                };

                let (old_global_x, old_global_y) = get_global_pos(Some(control_id), &control_map);
                let (new_parent_global_x, new_parent_global_y) = get_global_pos(new_parent_id, &control_map);

                // 2. Update the control
                if let Some(control) = form_module.form.get_control_mut(control_id) {
                    control.parent_id = new_parent_id;
                    
                    // Transform coordinates: new_local = old_global - new_parent_global
                    let mut new_x = old_global_x - new_parent_global_x;
                    let mut new_y = old_global_y - new_parent_global_y;
                    
                    // Snap to grid for sanity
                    new_x = (new_x / 10) * 10;
                    new_y = (new_y / 10) * 10;

                    control.bounds.x = new_x;
                    control.bounds.y = new_y;
                }
            }
        }
    }
}
