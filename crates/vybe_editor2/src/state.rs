use std::collections::HashMap;
use std::path::PathBuf;
use std::process::Child;
use std::sync::{Arc, Mutex};
use uuid::Uuid;
use vybe_forms::{Control, ControlType, Form};
use vybe_project::{FormModule, Project, StartupObject};

#[derive(Clone, Debug, PartialEq)]
pub enum View {
    FormDesigner,
    CodeEditor,
}

#[derive(Clone, Debug, PartialEq)]
pub enum RunStatus {
    Idle,
    Running,
    Done(String), // output/error message
}

pub struct EditorState {
    pub project: Option<Project>,
    pub project_path: Option<PathBuf>,

    pub current_form: Option<String>,   // form name
    pub current_code_file: Option<String>, // code file name
    pub view: View,

    pub selected_controls: Vec<Uuid>,
    pub selected_tool: Option<ControlType>,

    // Code editor buffer — keyed by form/file name
    pub code_buffers: HashMap<String, String>,

    // Run state
    pub run_status: RunStatus,
    pub run_child: Option<Child>,
    pub run_output: Arc<Mutex<Vec<String>>>,
    pub run_done: Arc<Mutex<bool>>,
    pub run_error: Arc<Mutex<Option<String>>>,

    // Form designer drag
    pub drag: Option<DragState>,

    // Undo stacks per form
    pub undo_stacks: HashMap<String, Vec<Vec<Control>>>,
    pub redo_stacks: HashMap<String, Vec<Vec<Control>>>,

    // UI toggles
    pub show_toolbox: bool,
    pub show_properties: bool,
    pub show_project_explorer: bool,
}

#[derive(Clone, Debug)]
pub struct DragState {
    pub control_id: Uuid,
    pub offset_x: f32,
    pub offset_y: f32,
    pub start_x: i32,
    pub start_y: i32,
}

impl EditorState {
    pub fn new(cli_path: Option<PathBuf>) -> Self {
        let mut s = Self {
            project: None,
            project_path: None,
            current_form: None,
            current_code_file: None,
            view: View::FormDesigner,
            selected_controls: Vec::new(),
            selected_tool: None,
            code_buffers: HashMap::new(),
            run_status: RunStatus::Idle,
            run_child: None,
            run_output: Arc::new(Mutex::new(Vec::new())),
            run_done: Arc::new(Mutex::new(false)),
            run_error: Arc::new(Mutex::new(None)),
            drag: None,
            undo_stacks: HashMap::new(),
            redo_stacks: HashMap::new(),
            show_toolbox: true,
            show_properties: true,
            show_project_explorer: true,
        };

        if let Some(path) = cli_path {
            s.load_project(&path);
        } else {
            s.new_project();
        }
        s
    }

    pub fn new_project(&mut self) {
        let mut project = Project::new("Project1");
        let mut form = Form::new("Form1");
        form.text = "Form1".to_string();
        form.width = 640;
        form.height = 480;
        let designer = vybe_forms::serialization::designer_codegen::generate_designer_code(&form);
        let user = vybe_forms::serialization::designer_codegen::generate_user_code_stub("Form1");
        project.forms.push(FormModule::new_vbnet(form, designer, user));
        project.startup_object = StartupObject::Form("Form1".to_string());
        project.startup_form = Some("Form1".to_string());
        self.project = Some(project);
        self.project_path = None;
        self.current_form = Some("Form1".to_string());
        self.view = View::FormDesigner;
        self.selected_controls.clear();
        self.code_buffers.clear();
    }

    pub fn load_project(&mut self, path: &PathBuf) {
        match vybe_project::load_project_auto(path) {
            Ok(proj) => {
                self.current_form = proj.forms.first().map(|f| f.form.name.clone());
                self.project_path = Some(path.clone());
                self.project = Some(proj);
                self.view = View::FormDesigner;
                self.selected_controls.clear();
                self.code_buffers.clear();
            }
            Err(e) => eprintln!("Failed to load project: {}", e),
        }
    }

    pub fn open_project_dialog(&mut self) {
        if let Some(path) = rfd::FileDialog::new()
            .add_filter("VB Project", &["vbp", "vbproj", "vybe"])
            .pick_file()
        {
            self.load_project(&path);
        }
    }

    pub fn save_project(&mut self) {
        if let Some(path) = &self.project_path.clone() {
            self.flush_code_buffers();
            if let Some(proj) = &self.project {
                let _ = vybe_project::save_project_auto(proj, path);
            }
        } else {
            self.save_project_as();
        }
    }

    pub fn save_project_as(&mut self) {
        if let Some(path) = rfd::FileDialog::new()
            .add_filter("VB Project", &["vbproj"])
            .save_file()
        {
            self.flush_code_buffers();
            if let Some(proj) = &self.project {
                let _ = vybe_project::save_project_auto(proj, &path);
            }
            self.project_path = Some(path);
        }
    }

    /// Flush code buffers back into the project before saving/running.
    pub fn flush_code_buffers(&mut self) {
        let Some(proj) = self.project.as_mut() else { return };
        for (name, code) in &self.code_buffers {
            // Try forms first
            if let Some(fm) = proj.forms.iter_mut().find(|f| &f.form.name == name) {
                fm.set_user_code(code.clone());
            } else if let Some(cf) = proj.code_files.iter_mut().find(|c| &c.name == name) {
                cf.code = code.clone();
            }
        }
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

    pub fn current_form_data(&self) -> Option<&Form> {
        let name = self.current_form.as_ref()?;
        self.project.as_ref()?.forms.iter().find(|f| &f.form.name == name).map(|f| &f.form)
    }

    pub fn current_form_data_mut(&mut self) -> Option<&mut Form> {
        let name = self.current_form.as_ref()?.clone();
        self.project.as_mut()?.forms.iter_mut().find(|f| f.form.name == name).map(|f| &mut f.form)
    }

    pub fn push_undo(&mut self) {
        let name = match &self.current_form { Some(n) => n.clone(), None => return };
        let controls = match self.current_form_data() { Some(f) => f.controls.clone(), None => return };
        self.undo_stacks.entry(name.clone()).or_default().push(controls);
        self.redo_stacks.remove(&name);
    }

    pub fn undo(&mut self) {
        let name = match &self.current_form { Some(n) => n.clone(), None => return };
        let snapshot = self.undo_stacks.entry(name.clone()).or_default().pop();
        if let Some(controls) = snapshot {
            let current = self.current_form_data().map(|f| f.controls.clone());
            if let Some(current) = current {
                self.redo_stacks.entry(name).or_default().push(current);
            }
            if let Some(form) = self.current_form_data_mut() {
                form.controls = controls;
            }
        }
    }

    pub fn redo(&mut self) {
        let name = match &self.current_form { Some(n) => n.clone(), None => return };
        let snapshot = self.redo_stacks.entry(name.clone()).or_default().pop();
        if let Some(controls) = snapshot {
            let current = self.current_form_data().map(|f| f.controls.clone());
            if let Some(current) = current {
                self.undo_stacks.entry(name).or_default().push(current);
            }
            if let Some(form) = self.current_form_data_mut() {
                form.controls = controls;
            }
        }
    }

    pub fn add_control(&mut self, ct: ControlType, x: i32, y: i32) {
        self.push_undo();
        let Some(form) = self.current_form_data_mut() else { return };
        let name = format!("{}{}", ct.default_name_prefix(), form.controls.len() + 1);
        let (w, h) = ct.default_size();
        let mut ctrl = Control::new(ct, name, x, y);
        ctrl.bounds = vybe_forms::Bounds::new(x, y, w, h);
        form.controls.push(ctrl);
    }

    pub fn delete_selected(&mut self) {
        self.push_undo();
        let ids = self.selected_controls.clone();
        if let Some(form) = self.current_form_data_mut() {
            form.controls.retain(|c| !ids.contains(&c.id));
        }
        self.selected_controls.clear();
    }

    pub fn run(&mut self) {
        self.flush_code_buffers();
        *self.run_output.lock().unwrap() = Vec::new();
        *self.run_done.lock().unwrap() = false;
        *self.run_error.lock().unwrap() = None;

        let Some(proj) = &self.project else { return };
        let has_forms = !proj.forms.is_empty();
        let is_sub_main = proj.code_files.iter().any(|cf| cf.code.to_uppercase().contains("SUB MAIN"));

        if has_forms && !is_sub_main {
            // Form project — spawn vybec
            if let Some(path) = &self.project_path {
                let vybec = std::env::current_exe()
                    .ok()
                    .and_then(|p| p.parent().map(|d| d.join("vybec")))
                    .unwrap_or_else(|| std::path::PathBuf::from("vybec"));

                match std::process::Command::new(&vybec).arg(path).spawn() {
                    Ok(child) => {
                        #[cfg(target_os = "macos")]
                        bring_to_front(child.id());
                        self.run_child = Some(child);
                        self.run_status = RunStatus::Running;
                    }
                    Err(e) => {
                        self.run_status = RunStatus::Done(format!("Could not launch vybec: {}", e));
                    }
                }
            } else {
                self.run_status = RunStatus::Done("Save the project first.".to_string());
            }
        } else {
            // Console project — run in thread
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
                    if let vybe_host::SideEffect::ConsoleOutput(msg) = effect {
                        out.lock().unwrap().push(msg);
                    }
                }
                *done.lock().unwrap() = true;
            });
        }
    }

    pub fn stop(&mut self) {
        if let Some(mut child) = self.run_child.take() {
            let _ = child.kill();
            let _ = child.wait();
        }
        self.run_status = RunStatus::Idle;
    }

    /// Poll run state — call every frame when running.
    pub fn poll_run(&mut self) {
        match &self.run_status {
            RunStatus::Running => {
                // Check subprocess
                if let Some(child) = &mut self.run_child {
                    match child.try_wait() {
                        Ok(Some(_)) => { self.run_child = None; self.run_status = RunStatus::Done("Program finished.".to_string()); }
                        Ok(None) => {}
                        Err(_) => { self.run_child = None; self.run_status = RunStatus::Done("Program finished.".to_string()); }
                    }
                }
                // Check console thread
                if *self.run_done.lock().unwrap() {
                    let lines = self.run_output.lock().unwrap().join("\n");
                    let err = self.run_error.lock().unwrap().clone();
                    let msg = if let Some(e) = err { format!("{}\n{}", lines, e) } else { lines };
                    self.run_status = RunStatus::Done(if msg.is_empty() { "Program finished.".to_string() } else { msg });
                }
            }
            _ => {}
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
