use std::collections::HashMap;
use std::path::PathBuf;
use std::process::Child;
use std::sync::{Arc, Mutex};
use uuid::Uuid;
use vybe_forms::{Control, ControlType, Form};
use vybe_project::{FormModule, Project, StartupObject};

#[derive(Clone, Debug, PartialEq)]
pub enum View { FormDesigner, CodeEditor, ResourceEditor }

#[derive(Clone, Debug, PartialEq)]
pub enum RunStatus { Idle, Running, Done(String) }

pub struct EditorState {
    pub project: Option<Project>,
    pub project_path: Option<PathBuf>,
    pub current_form: Option<String>,
    pub current_code_file: Option<String>,
    pub view: View,
    pub selected_controls: Vec<Uuid>,
    pub selected_tool: Option<ControlType>,
    pub clipboard: Vec<Control>,
    pub code_buffers: HashMap<String, String>,
    pub run_status: RunStatus,
    pub run_child: Option<Child>,
    pub run_output: Arc<Mutex<Vec<String>>>,
    pub run_done: Arc<Mutex<bool>>,
    pub run_error: Arc<Mutex<Option<String>>>,
    pub drag: Option<DragState>,
    pub lasso: Option<LassoState>,
    pub undo_stacks: HashMap<String, Vec<FormSnapshot>>,
    pub redo_stacks: HashMap<String, Vec<FormSnapshot>>,
    pub show_toolbox: bool,
    pub show_properties: bool,
    pub show_project_explorer: bool,
    pub show_project_properties: bool,
}

#[derive(Clone, Debug)]
pub struct FormSnapshot {
    pub controls: Vec<Control>,
    pub width: i32,
    pub height: i32,
    pub text: String,
    pub back_color: Option<String>,
}

#[derive(Clone, Debug)]
pub struct DragState {
    pub ids: Vec<Uuid>,
    pub start_mouse: egui::Pos2,
    pub initial_bounds: Vec<(Uuid, vybe_forms::Bounds)>,
}

#[derive(Clone, Debug)]
pub struct LassoState {
    pub origin: egui::Pos2,
    pub current: egui::Pos2,
}

impl EditorState {
    pub fn new(cli_path: Option<PathBuf>) -> Self {
        let mut s = Self {
            project: None, project_path: None,
            current_form: None, current_code_file: None,
            view: View::FormDesigner,
            selected_controls: Vec::new(), selected_tool: None, clipboard: Vec::new(),
            code_buffers: HashMap::new(),
            run_status: RunStatus::Idle, run_child: None,
            run_output: Arc::new(Mutex::new(Vec::new())),
            run_done: Arc::new(Mutex::new(false)),
            run_error: Arc::new(Mutex::new(None)),
            drag: None, lasso: None,
            undo_stacks: HashMap::new(), redo_stacks: HashMap::new(),
            show_toolbox: true, show_properties: true,
            show_project_explorer: true, show_project_properties: false,
        };
        if let Some(path) = cli_path { s.load_project(&path); } else { s.new_project(); }
        s
    }

    // ── Project ───────────────────────────────────────────────────────────────

    pub fn new_project(&mut self) {
        let mut project = Project::new("Project1");
        let mut form = Form::new("Form1");
        form.text = "Form1".to_string(); form.width = 640; form.height = 480;
        let designer = vybe_forms::serialization::designer_codegen::generate_designer_code(&form);
        let user = vybe_forms::serialization::designer_codegen::generate_user_code_stub("Form1");
        project.forms.push(FormModule::new_vbnet(form, designer, user));
        project.startup_object = StartupObject::Form("Form1".to_string());
        project.startup_form = Some("Form1".to_string());
        self.project = Some(project);
        self.project_path = None;
        self.current_form = Some("Form1".to_string());
        self.current_code_file = None;
        self.view = View::FormDesigner;
        self.selected_controls.clear();
        self.code_buffers.clear();
        self.undo_stacks.clear();
        self.redo_stacks.clear();
    }

    pub fn load_project(&mut self, path: &PathBuf) {
        match vybe_project::load_project_auto(path) {
            Ok(proj) => {
                self.current_form = proj.forms.first().map(|f| f.form.name.clone());
                self.current_code_file = None;
                self.project_path = Some(path.clone());
                self.project = Some(proj);
                self.view = View::FormDesigner;
                self.selected_controls.clear();
                self.code_buffers.clear();
                self.undo_stacks.clear();
                self.redo_stacks.clear();
            }
            Err(e) => eprintln!("Failed to load project: {}", e),
        }
    }

    pub fn open_project_dialog(&mut self) {
        if let Some(path) = rfd::FileDialog::new()
            .add_filter("VB Project", &["vbp", "vbproj", "vybe"])
            .pick_file()
        { self.load_project(&path); }
    }

    pub fn save_project(&mut self) {
        if let Some(path) = self.project_path.clone() {
            self.flush_code_buffers();
            if let Some(proj) = &self.project { let _ = vybe_project::save_project_auto(proj, &path); }
        } else { self.save_project_as(); }
    }

    pub fn save_project_as(&mut self) {
        if let Some(path) = rfd::FileDialog::new().add_filter("VB Project", &["vbproj"]).save_file() {
            self.flush_code_buffers();
            if let Some(proj) = &self.project { let _ = vybe_project::save_project_auto(proj, &path); }
            self.project_path = Some(path);
        }
    }

    pub fn add_new_form(&mut self) {
        let Some(proj) = self.project.as_mut() else { return };
        let mut n = 1;
        let mut name = format!("Form{}", n);
        while proj.get_form(&name).is_some() { n += 1; name = format!("Form{}", n); }
        let mut form = Form::new(&name);
        form.text = name.clone(); form.width = 640; form.height = 480;
        let designer = vybe_forms::serialization::designer_codegen::generate_designer_code(&form);
        let user = vybe_forms::serialization::designer_codegen::generate_user_code_stub(&name);
        proj.forms.push(FormModule::new_vbnet(form, designer, user));
        self.current_form = Some(name);
        self.current_code_file = None;
        self.view = View::FormDesigner;
    }

    pub fn add_code_file(&mut self) {
        let Some(proj) = self.project.as_mut() else { return };
        let mut n = 1;
        let mut name = format!("Module{}", n);
        while proj.get_code_file(&name).is_some() { n += 1; name = format!("Module{}", n); }
        let mut cf = vybe_project::CodeFile::new(&name);
        cf.code = format!("Module {}\n\nEnd Module\n", name);
        proj.add_code_file(cf);
        self.current_code_file = Some(name);
        self.current_form = None;
        self.view = View::CodeEditor;
    }

    pub fn add_existing_form(&mut self) {
        let Some(paths) = rfd::FileDialog::new()
            .set_title("Add Existing Form").add_filter("VB Forms", &["vb"]).pick_files()
        else { return };
        for path in paths {
            match vybe_project::load_form_vb(&path) {
                Ok(fm) => {
                    let name = fm.form.name.clone();
                    if let Some(proj) = self.project.as_mut() {
                        if proj.get_form(&name).is_none() {
                            proj.forms.push(fm);
                            self.current_form = Some(name);
                            self.view = View::FormDesigner;
                        }
                    }
                }
                Err(e) => eprintln!("Failed to load form: {}", e),
            }
        }
    }

    pub fn add_existing_code_file(&mut self) {
        let Some(paths) = rfd::FileDialog::new()
            .set_title("Add Existing Code File").add_filter("Code Files", &["vb", "bas"]).pick_files()
        else { return };
        for path in paths {
            let name = path.file_stem().unwrap_or_default().to_string_lossy().to_string();
            let code = match vybe_project::read_text_file(&path) {
                Ok(c) => c, Err(e) => { eprintln!("Failed to read: {}", e); continue; }
            };
            if let Some(proj) = self.project.as_mut() {
                if proj.get_code_file(&name).is_none() {
                    let mut cf = vybe_project::CodeFile::new(&name);
                    cf.code = code;
                    proj.add_code_file(cf);
                    self.current_code_file = Some(name);
                    self.view = View::CodeEditor;
                }
            }
        }
    }

    pub fn remove_project_item(&mut self, name: &str) {
        let Some(proj) = self.project.as_mut() else { return };
        let removed = proj.remove_form(name) || proj.remove_code_file(name);
        if removed && self.current_form.as_deref() == Some(name) {
            self.current_form = proj.forms.first().map(|f| f.form.name.clone());
            if self.current_form.is_none() {
                self.current_code_file = proj.code_files.first().map(|c| c.name.clone());
                if self.current_code_file.is_some() { self.view = View::CodeEditor; }
            }
        }
    }

    // ── Code buffers ──────────────────────────────────────────────────────────

    pub fn flush_code_buffers(&mut self) {
        let Some(proj) = self.project.as_mut() else { return };
        for (name, code) in &self.code_buffers {
            if let Some(fm) = proj.forms.iter_mut().find(|f| &f.form.name == name) {
                fm.set_user_code(code.clone());
            } else if let Some(cf) = proj.code_files.iter_mut().find(|c| &c.name == name) {
                cf.code = code.clone();
            }
        }
    }

    pub fn current_edit_name(&self) -> Option<&str> {
        self.current_form.as_deref().or(self.current_code_file.as_deref())
    }

    pub fn get_code_buffer(&mut self, name: &str) -> &mut String {
        if !self.code_buffers.contains_key(name) {
            let mut code = String::new();
            if let Some(p) = &self.project {
                if let Some(fm) = p.forms.iter().find(|f| f.form.name == name) {
                    code = fm.get_user_code().to_string();
                } else if let Some(cf) = p.code_files.iter().find(|c| c.name == name) {
                    code = cf.code.clone();
                }
            }
            self.code_buffers.insert(name.to_string(), code);
        }
        self.code_buffers.get_mut(name).unwrap()
    }

    // ── Form data access ──────────────────────────────────────────────────────

    pub fn current_form_data(&self) -> Option<&Form> {
        let name = self.current_form.as_ref()?;
        self.project.as_ref()?.forms.iter().find(|f| &f.form.name == name).map(|f| &f.form)
    }

    pub fn current_form_data_mut(&mut self) -> Option<&mut Form> {
        let name = self.current_form.as_ref()?.clone();
        self.project.as_mut()?.forms.iter_mut().find(|f| f.form.name == name).map(|f| &mut f.form)
    }

    // ── Undo / Redo ───────────────────────────────────────────────────────────

    pub fn push_undo(&mut self) {
        let name = match &self.current_form { Some(n) => n.clone(), None => return };
        let snap = match self.current_form_data() {
            Some(f) => FormSnapshot { controls: f.controls.clone(), width: f.width, height: f.height, text: f.text.clone(), back_color: f.back_color.clone() },
            None => return,
        };
        let stack = self.undo_stacks.entry(name.clone()).or_default();
        stack.push(snap);
        if stack.len() > 50 { stack.remove(0); }
        self.redo_stacks.remove(&name);
    }

    pub fn undo(&mut self) {
        let name = match &self.current_form { Some(n) => n.clone(), None => return };
        let snap = self.undo_stacks.entry(name.clone()).or_default().pop();
        if let Some(s) = snap {
            let cur = self.current_form_data().map(|f| FormSnapshot { controls: f.controls.clone(), width: f.width, height: f.height, text: f.text.clone(), back_color: f.back_color.clone() });
            if let Some(c) = cur { self.redo_stacks.entry(name).or_default().push(c); }
            if let Some(form) = self.current_form_data_mut() {
                form.controls = s.controls; form.width = s.width; form.height = s.height; form.text = s.text; form.back_color = s.back_color;
            }
        }
    }

    pub fn redo(&mut self) {
        let name = match &self.current_form { Some(n) => n.clone(), None => return };
        let snap = self.redo_stacks.entry(name.clone()).or_default().pop();
        if let Some(s) = snap {
            let cur = self.current_form_data().map(|f| FormSnapshot { controls: f.controls.clone(), width: f.width, height: f.height, text: f.text.clone(), back_color: f.back_color.clone() });
            if let Some(c) = cur { self.undo_stacks.entry(name).or_default().push(c); }
            if let Some(form) = self.current_form_data_mut() {
                form.controls = s.controls; form.width = s.width; form.height = s.height; form.text = s.text; form.back_color = s.back_color;
            }
        }
    }

    pub fn can_undo(&self) -> bool {
        self.current_form.as_ref().and_then(|n| self.undo_stacks.get(n)).map(|s| !s.is_empty()).unwrap_or(false)
    }

    pub fn can_redo(&self) -> bool {
        self.current_form.as_ref().and_then(|n| self.redo_stacks.get(n)).map(|s| !s.is_empty()).unwrap_or(false)
    }

    // ── Controls ──────────────────────────────────────────────────────────────

    pub fn add_control(&mut self, ct: ControlType, x: i32, y: i32) {
        self.push_undo();
        let Some(proj) = self.project.as_mut() else { return };
        let name_ref = match &self.current_form { Some(n) => n.clone(), None => return };
        let Some(fm) = proj.forms.iter_mut().find(|f| f.form.name == name_ref) else { return };
        let prefix = ct.default_name_prefix();
        let mut n = 1;
        let mut ctrl_name = format!("{}{}", prefix, n);
        while fm.form.get_control_by_name(&ctrl_name).is_some() { n += 1; ctrl_name = format!("{}{}", prefix, n); }
        let ctrl = Control::new(ct, ctrl_name, x, y);
        fm.form.add_control(ctrl);
        fm.sync_designer_code();
    }

    pub fn delete_selected(&mut self) {
        if self.selected_controls.is_empty() { return; }
        self.push_undo();
        let ids = self.selected_controls.clone();
        let Some(proj) = self.project.as_mut() else { return };
        let name_ref = match &self.current_form { Some(n) => n.clone(), None => return };
        if let Some(fm) = proj.forms.iter_mut().find(|f| f.form.name == name_ref) {
            fm.form.controls.retain(|c| !ids.contains(&c.id));
            fm.sync_designer_code();
        }
        self.selected_controls.clear();
    }

    pub fn copy_selected(&mut self) {
        let ids = self.selected_controls.clone();
        if ids.is_empty() { return; }
        if let Some(form) = self.current_form_data() {
            self.clipboard = form.controls.iter().filter(|c| ids.contains(&c.id)).cloned().collect();
        }
    }

    pub fn cut_selected(&mut self) { self.copy_selected(); self.delete_selected(); }

    pub fn paste(&mut self) {
        if self.clipboard.is_empty() { return; }
        self.push_undo();
        let clipboard = self.clipboard.clone();
        let Some(proj) = self.project.as_mut() else { return };
        let name_ref = match &self.current_form { Some(n) => n.clone(), None => return };
        if let Some(fm) = proj.forms.iter_mut().find(|f| f.form.name == name_ref) {
            let mut new_ids = Vec::new();
            for src in &clipboard {
                let mut ctrl = src.clone();
                ctrl.id = Uuid::new_v4();
                ctrl.parent_id = None;
                ctrl.bounds.x += 20; ctrl.bounds.y += 20;
                if src.index.is_none() {
                    let prefix = ctrl.control_type.default_name_prefix();
                    let mut n = 1;
                    let mut new_name = format!("{}{}", prefix, n);
                    while fm.form.get_control_by_name(&new_name).is_some() { n += 1; new_name = format!("{}{}", prefix, n); }
                    ctrl.name = new_name;
                }
                new_ids.push(ctrl.id);
                fm.form.add_control(ctrl);
            }
            fm.sync_designer_code();
            self.selected_controls = new_ids;
        }
    }

    pub fn update_control_property(&mut self, id: Uuid, property: &str, value: String) {
        let Some(proj) = self.project.as_mut() else { return };
        let name_ref = match &self.current_form { Some(n) => n.clone(), None => return };
        let Some(fm) = proj.forms.iter_mut().find(|f| f.form.name == name_ref) else { return };
        let Some(ctrl) = fm.form.get_control_mut(id) else { return };
        match property {
            "Name" => { let v = value.trim().to_string(); if !v.is_empty() { ctrl.name = v; } }
            "Text" | "Caption" => ctrl.set_text(value),
            "BackColor" => ctrl.set_back_color(value),
            "ForeColor" => ctrl.set_fore_color(value),
            "Font" => ctrl.set_font(value),
            "Enabled" => { if let Ok(b) = value.parse::<bool>() { ctrl.set_enabled(b); } }
            "Visible" => { if let Ok(b) = value.parse::<bool>() { ctrl.set_visible(b); } }
            "TabIndex" => { if let Ok(n) = value.parse::<i32>() { ctrl.tab_index = n; } }
            "Checked" => {
                if let Ok(b) = value.parse::<bool>() {
                    ctrl.properties.set("Checked".to_string(), b);
                    use vybe_forms::properties::PropertyValue;
                    ctrl.properties.set_raw("CheckState", PropertyValue::Integer(if b { 1 } else { 0 }));
                    ctrl.properties.set_raw("Value", PropertyValue::Integer(if b { 1 } else { 0 }));
                }
            }
            "HTML" => { ctrl.properties.set("HTML".to_string(), value.clone()); ctrl.set_text(value); }
            "ThreeState" | "Multiline" | "ReadOnly" | "Sorted" | "ShowCheckBox" |
            "ShowUpDown" | "IsSplitterFixed" | "WrapContents" | "ShowToday" |
            "ShowWeekNumbers" | "AutoScroll" | "CheckOnClick" | "CheckBoxes" |
            "ShowLines" | "ShowRootLines" | "ShowPlusMinus" | "LabelEdit" |
            "FullRowSelect" | "GridLines" | "MultiSelect" | "AllowUserToAddRows" |
            "AllowUserToDeleteRows" | "AutoGenerateColumns" | "WordWrap" => {
                if let Ok(b) = value.parse::<bool>() { ctrl.properties.set(property.to_string(), b); }
            }
            "Minimum" | "Maximum" | "Step" | "Increment" | "DecimalPlaces" |
            "TickFrequency" | "SmallChange" | "LargeChange" | "SplitterDistance" |
            "MaxLength" | "MaxDropDownItems" => {
                if let Ok(n) = value.parse::<i32>() {
                    use vybe_forms::properties::PropertyValue;
                    ctrl.properties.set_raw(property, PropertyValue::Integer(n));
                }
            }
            "DataSource" | "DataMember" | "Filter" | "Sort" | "BindingSource" |
            "SelectCommand" | "ConnectionString" | "DisplayMember" | "ValueMember" |
            "DbType" | "DbPath" | "DbHost" | "DbPort" | "DbName" | "DbUser" | "DbPassword" |
            "Format" | "CustomFormat" | "Mask" | "PasswordChar" | "URL" |
            "BorderStyle" | "Alignment" | "Orientation" | "FlowDirection" => {
                ctrl.properties.set(property.to_string(), value);
            }
            prop if prop.starts_with("DataBindings.") => { ctrl.properties.set(prop.to_string(), value); }
            _ => { ctrl.properties.set(property.to_string(), value); }
        }
        fm.sync_designer_code();
    }

    pub fn update_control_geometry(&mut self, id: Uuid, x: i32, y: i32, w: i32, h: i32) {
        let Some(proj) = self.project.as_mut() else { return };
        let name_ref = match &self.current_form { Some(n) => n.clone(), None => return };
        if let Some(fm) = proj.forms.iter_mut().find(|f| f.form.name == name_ref) {
            if let Some(ctrl) = fm.form.get_control_mut(id) {
                ctrl.bounds.x = x; ctrl.bounds.y = y;
                ctrl.bounds.width = w.max(10); ctrl.bounds.height = h.max(10);
            }
            fm.sync_designer_code();
        }
    }

    // ── Run ───────────────────────────────────────────────────────────────────

    pub fn run(&mut self) {
        self.flush_code_buffers();
        *self.run_output.lock().unwrap() = Vec::new();
        *self.run_done.lock().unwrap() = false;
        *self.run_error.lock().unwrap() = None;
        let Some(proj) = &self.project else { return };
        let has_forms = !proj.forms.is_empty();
        let is_sub_main = proj.code_files.iter().any(|cf| cf.code.to_uppercase().contains("SUB MAIN"));
        if has_forms && !is_sub_main {
            if let Some(path) = &self.project_path {
                let vybec = std::env::current_exe().ok()
                    .and_then(|p| p.parent().map(|d| d.join("vybec")))
                    .unwrap_or_else(|| std::path::PathBuf::from("vybec"));
                match std::process::Command::new(&vybec).arg(path).spawn() {
                    Ok(child) => {
                        #[cfg(target_os = "macos")] bring_to_front(child.id());
                        self.run_child = Some(child);
                        self.run_status = RunStatus::Running;
                    }
                    Err(e) => self.run_status = RunStatus::Done(format!("Could not launch vybec: {}", e)),
                }
            } else {
                self.run_status = RunStatus::Done("Save the project first.".to_string());
            }
        } else {
            let (all_code, _) = vybe_cli::runner::build_project_code(proj);
            let out = self.run_output.clone();
            let done = self.run_done.clone();
            let err = self.run_error.clone();
            self.run_status = RunStatus::Running;
            std::thread::spawn(move || {
                let program = match vybe_parser_basic::parse_program(&all_code) {
                    Ok(p) => p,
                    Err(e) => { *err.lock().unwrap() = Some(format!("Parse error: {:?}", e)); *done.lock().unwrap() = true; return; }
                };
                let chunks = match vybe_compiler_vb::Compiler::new().compile(&program) {
                    Ok(c) => c,
                    Err(e) => { *err.lock().unwrap() = Some(format!("Compile error: {}", e)); *done.lock().unwrap() = true; return; }
                };
                let mut vm = vybe_bytecode::VM::new();
                let queue = std::rc::Rc::new(std::cell::RefCell::new(vybe_host::SideEffectQueue::new()));
                vybe_host::register_all_with_gui(&mut vm, queue.clone());
                vybe_host::setup_namespaces(&mut vm);
                let cap = out.clone();
                vm.register_host_fn("wasi:cli", "log", Box::new(move |args: &[vybe_bytecode::Value]| {
                    cap.lock().unwrap().push(args.iter().map(|v| format!("{v}")).collect::<Vec<_>>().join(" "));
                    vybe_bytecode::Value::Null
                }));
                if let Err(e) = vm.run(chunks) {
                    let msg = format!("{}", e);
                    if !msg.starts_with("__") { *err.lock().unwrap() = Some(format!("Runtime error: {}", msg)); }
                }
                for effect in queue.borrow_mut().drain() {
                    if let vybe_host::SideEffect::ConsoleOutput(msg) = effect { out.lock().unwrap().push(msg); }
                }
                *done.lock().unwrap() = true;
            });
        }
    }

    pub fn stop(&mut self) {
        if let Some(mut child) = self.run_child.take() { let _ = child.kill(); let _ = child.wait(); }
        self.run_status = RunStatus::Idle;
    }

    pub fn poll_run(&mut self) {
        if let RunStatus::Running = &self.run_status {
            if let Some(child) = &mut self.run_child {
                match child.try_wait() {
                    Ok(Some(_)) | Err(_) => { self.run_child = None; self.run_status = RunStatus::Done("Program finished.".to_string()); }
                    Ok(None) => {}
                }
            }
            if *self.run_done.lock().unwrap() {
                let lines = self.run_output.lock().unwrap().join("\n");
                let err = self.run_error.lock().unwrap().clone();
                let msg = if let Some(e) = err { if lines.is_empty() { e } else { format!("{}\n{}", lines, e) } } else { lines };
                self.run_status = RunStatus::Done(if msg.is_empty() { "Program finished.".to_string() } else { msg });
            }
        }
    }
}

#[cfg(target_os = "macos")]
fn bring_to_front(pid: u32) {
    std::thread::sleep(std::time::Duration::from_millis(300));
    let _ = std::process::Command::new("osascript")
        .arg("-e")
        .arg(format!("tell application \"System Events\" to set frontmost of (first process whose unix id is {}) to true", pid))
        .spawn();
}
