use uuid::Uuid;
use irys_forms::{Control, ControlType, Form};
use irys_project::Project;

#[derive(Debug, Clone, PartialEq)]
pub enum Tool {
    Select,
    Control(ControlType),
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum View {
    Designer,
    Code,
}

#[derive(Clone)]
pub struct DragState {
    pub control_id: Uuid,
    pub start_x: i32,
    pub start_y: i32,
    pub offset_x: i32,
    pub offset_y: i32,
}

pub struct EditorState {
    pub project: Option<Project>,
    pub current_form: Option<String>,
    pub selected_tool: Tool,
    pub selected_control: Option<Uuid>,
    pub form_selected: bool,
    pub current_view: View,
    pub drag_state: Option<DragState>,
    pub resize_state: Option<ResizeState>,
    pub control_counter: usize,

    // Window visibility
    pub show_project_explorer: bool,
    pub show_properties: bool,
    pub show_toolbox: bool,
    pub show_immediate: bool,
    pub show_project_properties: bool,

    // Running state
    pub is_running: bool,
    
    // Project state
    pub current_project_path: Option<std::path::PathBuf>,

    pub msgbox_pending: Option<String>,
}

#[derive(Clone)]
pub struct ResizeState {
    pub control_id: Uuid,
    pub handle: ResizeHandle,
    pub original_bounds: irys_forms::Bounds,
    pub start_x: i32,
    pub start_y: i32,
}

#[derive(Clone, Copy, PartialEq)]
pub enum ResizeHandle {
    TopLeft,
    Top,
    TopRight,
    Left,
    Right,
    BottomLeft,
    Bottom,
    BottomRight,
}

impl EditorState {
    pub fn new() -> Self {
        Self {
            project: None,
            current_form: None,
            selected_tool: Tool::Select,
            selected_control: None,
            form_selected: false,
            current_view: View::Designer,
            drag_state: None,
            resize_state: None,
            control_counter: 0,
            show_project_explorer: true,
            show_properties: true,
            show_toolbox: true,
            show_immediate: false,
            show_project_properties: false,
            is_running: false,
            current_project_path: None,
            msgbox_pending: None,
        }
    }

    pub fn get_current_form(&self) -> Option<&Form> {
        let name = self.current_form.as_ref()?;
        self.project
            .as_ref()
            .and_then(|p| p.get_form(name))
            .map(|fm| &fm.form)
    }

    pub fn get_current_form_mut(&mut self) -> Option<&mut Form> {
        let name = self.current_form.clone()?;
        self.project
            .as_mut()
            .and_then(|p| p.get_form_mut(&name))
            .map(|fm| &mut fm.form)
    }

    pub fn get_current_code(&self) -> String {
        let name = match self.current_form.as_ref() {
            Some(n) => n,
            None => return String::new(),
        };
        self.project
            .as_ref()
            .and_then(|p| p.get_form(name))
            .map(|fm| fm.code.clone())
            .unwrap_or_default()
    }

    pub fn set_current_code(&mut self, code: String) {
        let name = match &self.current_form {
            Some(n) => n.clone(),
            None => return,
        };

        if let Some(project) = &mut self.project {
            if let Some(form_module) = project.get_form_mut(&name) {
                form_module.code = code;
            }
        }
    }

    pub fn generate_control_name(&mut self, control_type: &ControlType) -> String {
        self.control_counter += 1;
        format!("{}{}", control_type.default_name_prefix(), self.control_counter)
    }

    pub fn get_selected_control(&self) -> Option<&Control> {
        self.selected_control
            .and_then(|_id| self.get_current_form())
            .and_then(|form| form.get_control(self.selected_control?))
    }
}

impl Default for EditorState {
    fn default() -> Self {
        Self::new()
    }
}
