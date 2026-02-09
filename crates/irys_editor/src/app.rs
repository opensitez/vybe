use crate::message::Message;
use crate::state::{EditorState, Tool, View};
use crate::views::{code_editor_view, menu_bar, toolbar, project_explorer, properties_panel, toolbox, Designer};
use iced::widget::{button, column, container, row, text};
use iced::{Application, Command, Element, Length, Theme};
use irys_forms::{Control, EventBinding, EventType, Form};
use irys_project::Project;

pub struct irysEditorApp {
    state: EditorState,
    designer: Designer,
}

impl irysEditorApp {
    pub fn new() -> (Self, Command<Message>) {
        // Auto-create a default project so the app is immediately usable
        let mut state = EditorState::new();
        let mut project = Project::new("NewProject");
        let mut form = Form::new("Form1");
        form.caption = "Form1".to_string();
        project.add_form(form);
        state.project = Some(project);
        state.current_form = Some("Form1".to_string());

        (
            Self {
                state,
                designer: Designer::new(),
            },
            Command::none(),
        )
    }

    pub fn title(&self) -> String {
        let project_name = self.state.project
            .as_ref()
            .map(|p| p.name.clone())
            .unwrap_or_else(|| "Untitled".to_string());

        format!("irys Editor - {}", project_name)
    }

    pub fn update(&mut self, message: Message) -> Command<Message> {
        match message {
            Message::NewProject => {
                let mut project = Project::new("NewProject");
                let mut form = Form::new("Form1");
                form.caption = "Form1".to_string();
                project.add_form(form);

                self.state.project = Some(project);
                self.state.current_form = Some("Form1".to_string());
                self.state.current_view = View::Designer;
                self.designer.clear_cache();
            }

            Message::OpenProject => {
                // TODO: Implement file dialog and loading
                println!("Open project not yet implemented");
            }

            Message::SaveProject => {
                // TODO: Implement file dialog and saving
                if let Some(project) = &self.state.project {
                    println!("Would save project: {}", project.name);
                }
            }

            Message::Start => {
                if let Some(project) = &self.state.project {
                    self.state.is_running = true;
                    self.run_project(project);
                    self.state.is_running = false;
                }
            }

            Message::Stop => {
                self.state.is_running = false;
            }

            Message::Restart => {
                self.state.is_running = false;
                // TODO: Implement restart logic
            }

            Message::NewForm => {
                if let Some(project) = &mut self.state.project {
                    let form_num = project.forms.len() + 1;
                    let form_name = format!("Form{}", form_num);
                    let mut form = Form::new(&form_name);
                    form.caption = form_name.clone();
                    project.add_form(form);
                    self.state.current_form = Some(form_name);
                    self.designer.clear_cache();
                }
            }

            Message::SelectForm(form_name) => {
                self.state.current_form = Some(form_name);
                self.state.selected_control = None;
                self.designer.clear_cache();
            }

            Message::GenerateEventHandlers => {
                if let Some(form) = self.state.get_current_form() {
                    let code = generate_form_code(form);
                    self.state.set_current_code(code);
                }
            }

            Message::SelectTool(control_type) => {
                println!("DEBUG: Tool selected: {:?}", control_type);
                self.state.selected_tool = Tool::Control(control_type);
            }

            Message::CanvasClicked(x, y) => {
                println!("DEBUG: Canvas clicked at ({}, {}), tool: {:?}", x, y, self.state.selected_tool);
                match self.state.selected_tool.clone() {
                    Tool::Select => {
                        // Select control at position and start drag
                        let drag_info = if let Some(form) = self.state.get_current_form() {
                            form.find_control_at(x, y).map(|control| {
                                (control.id, control.bounds.x, control.bounds.y)
                            })
                        } else {
                            None
                        };

                        if let Some((id, bounds_x, bounds_y)) = drag_info {
                            self.state.selected_control = Some(id);
                            self.state.form_selected = false;
                            self.state.drag_state = Some(crate::state::DragState {
                                control_id: id,
                                start_x: x,
                                start_y: y,
                                offset_x: x - bounds_x,
                                offset_y: y - bounds_y,
                            });
                        } else {
                            // Clicked on empty canvas - select the form
                            self.state.selected_control = None;
                            self.state.form_selected = true;
                            self.state.drag_state = None;
                        }
                        self.designer.clear_cache();
                    }
                    Tool::Control(control_type) => {
                        // Generate name first
                        let name = self.state.generate_control_name(&control_type);
                        println!("DEBUG: Placing {} control named {} at ({}, {})",
                                 control_type.as_str(), name, x, y);

                        // Place new control
                        if let Some(form) = self.state.get_current_form_mut() {
                            let control = Control::new(control_type, name.clone(), x, y);

                            // Auto-generate event binding for Click event
                            let binding = EventBinding::new(&name, EventType::Click);
                            form.add_event_binding(binding);

                            form.add_control(control);
                            println!("DEBUG: Control added. Total controls: {}", form.controls.len());
                        } else {
                            println!("DEBUG: ERROR - No current form!");
                        }
                        self.state.selected_tool = Tool::Select;
                        self.designer.clear_cache();
                    }
                }
            }

            Message::ControlSelected(id) => {
                self.state.selected_control = Some(id);
                self.designer.clear_cache();
            }

            Message::ControlMoved(id, new_x, new_y) => {
                if let Some(form) = self.state.get_current_form_mut() {
                    if let Some(control) = form.get_control_mut(id) {
                        control.bounds.x = new_x;
                        control.bounds.y = new_y;
                        self.designer.clear_cache();
                    }
                }
            }

            Message::ControlResized(id, new_width, new_height) => {
                if let Some(form) = self.state.get_current_form_mut() {
                    if let Some(control) = form.get_control_mut(id) {
                        control.bounds.width = new_width;
                        control.bounds.height = new_height;
                        self.designer.clear_cache();
                    }
                }
            }

            Message::PropertyChanged(prop_name, value) => {
                // Handle form properties
                if self.state.form_selected {
                    if let Some(form) = self.state.get_current_form_mut() {
                        match prop_name.as_str() {
                            "FormCaption" => {
                                form.caption = value;
                            }
                            "FormWidth" => {
                                if let Ok(width) = value.parse::<i32>() {
                                    form.width = width.max(200);
                                }
                            }
                            "FormHeight" => {
                                if let Ok(height) = value.parse::<i32>() {
                                    form.height = height.max(200);
                                }
                            }
                            _ => {}
                        }
                        self.designer.clear_cache();
                    }
                }
                // Handle control properties
                else if let Some(id) = self.state.selected_control {
                    if let Some(form) = self.state.get_current_form_mut() {
                        if let Some(control) = form.get_control_mut(id) {
                            match prop_name.as_str() {
                                "Caption" => control.set_caption(value),
                                "Text" => control.set_text(value),
                                "Enabled" => {
                                    if let Ok(enabled) = value.parse::<bool>() {
                                        control.set_enabled(enabled);
                                    }
                                }
                                "Visible" => {
                                    if let Ok(visible) = value.parse::<bool>() {
                                        control.set_visible(visible);
                                    }
                                }
                                "X" => {
                                    if let Ok(x) = value.parse::<i32>() {
                                        control.bounds.x = x;
                                    }
                                }
                                "Y" => {
                                    if let Ok(y) = value.parse::<i32>() {
                                        control.bounds.y = y;
                                    }
                                }
                                "Width" => {
                                    if let Ok(width) = value.parse::<i32>() {
                                        control.bounds.width = width.max(10);
                                    }
                                }
                                "Height" => {
                                    if let Ok(height) = value.parse::<i32>() {
                                        control.bounds.height = height.max(10);
                                    }
                                }
                                _ => {}
                            }
                            self.designer.clear_cache();
                        }
                    }
                }
            }

            Message::DeleteControl => {
                if let Some(id) = self.state.selected_control {
                    if let Some(form) = self.state.get_current_form_mut() {
                        form.remove_control(id);
                        self.state.selected_control = None;
                        self.designer.clear_cache();
                    }
                }
            }

            Message::CodeChanged(code) => {
                self.state.set_current_code(code);
            }

            Message::ViewCode => {
                self.state.current_view = View::Code;
            }

            Message::ViewDesigner => {
                self.state.current_view = View::Designer;
            }

            Message::DesignerMouseDown(x, y) => {
                let drag_info = if let Some(form) = self.state.get_current_form() {
                    form.find_control_at(x, y).map(|control| {
                        (control.id, control.bounds.x, control.bounds.y)
                    })
                } else {
                    None
                };

                if let Some((id, bounds_x, bounds_y)) = drag_info {
                    self.state.selected_control = Some(id);
                    self.state.drag_state = Some(crate::state::DragState {
                        control_id: id,
                        start_x: x,
                        start_y: y,
                        offset_x: x - bounds_x,
                        offset_y: y - bounds_y,
                    });
                    self.designer.clear_cache();
                }
            }

            Message::DesignerMouseMove(x, y) => {
                if let Some(drag_state) = self.state.drag_state.clone() {
                    let new_x = x - drag_state.offset_x;
                    let new_y = y - drag_state.offset_y;

                    if let Some(form) = self.state.get_current_form_mut() {
                        if let Some(control) = form.get_control_mut(drag_state.control_id) {
                            control.bounds.x = new_x;
                            control.bounds.y = new_y;
                            self.designer.clear_cache();
                        }
                    }
                }
            }

            Message::DesignerMouseUp => {
                self.state.drag_state = None;
                self.state.resize_state = None;
            }

            Message::ControlDoubleClicked(id) => {
                // Double-click opens code editor for that control's default event
                let control_name = if let Some(form) = self.state.get_current_form() {
                    form.get_control(id).map(|c| c.name.clone())
                } else {
                    None
                };

                if let Some(name) = control_name {
                    // Switch to code view
                    self.state.current_view = View::Code;

                    // Generate event handler if it doesn't exist
                    let handler_name = format!("{}_Click", name);
                    let mut code = self.state.get_current_code();

                    // Check if handler already exists
                    if !code.contains(&handler_name) {
                        // Add the event handler
                        code.push_str(&format!("\nPrivate Sub {}()\n", handler_name));
                        code.push_str(&format!("    ' TODO: Handle {} click event\n", name));
                        code.push_str("    MsgBox(\"Button clicked!\")\n");
                        code.push_str("End Sub\n");
                        self.state.set_current_code(code);
                    }
                }
            }

            Message::FormSelected => {
                self.state.selected_control = None;
                self.state.form_selected = true;
                self.designer.clear_cache();
            }

            Message::StartResize(_id, _handle) => {
                // TODO: Implement resize handle dragging
                // For now, users can resize via properties panel
            }

            // View menu
            Message::ToggleProjectExplorer => {
                self.state.show_project_explorer = !self.state.show_project_explorer;
            }

            Message::TogglePropertiesWindow => {
                self.state.show_properties = !self.state.show_properties;
            }

            Message::ToggleToolbox => {
                self.state.show_toolbox = !self.state.show_toolbox;
            }

            Message::ToggleImmediateWindow => {
                self.state.show_immediate = !self.state.show_immediate;
            }

            // Project menu
            Message::AddForm => {
                if let Some(project) = &mut self.state.project {
                    let form_num = project.forms.len() + 1;
                    let form_name = format!("Form{}", form_num);
                    let mut form = Form::new(&form_name);
                    form.caption = form_name.clone();
                    project.add_form(form);
                    self.state.current_form = Some(form_name);
                    self.designer.clear_cache();
                }
            }

            Message::AddModule => {
                // TODO: Implement add module
            }

            Message::AddClass => {
                // TODO: Implement add class
            }

            Message::RemoveForm(_name) => {
                // TODO: Implement remove form
            }

            Message::ProjectProperties => {
                // TODO: Show project properties dialog
            }

            Message::Components => {
                // TODO: Show components dialog
            }

            // Edit menu
            Message::Undo | Message::Redo | Message::Cut | Message::Copy |
            Message::Paste | Message::Delete | Message::SelectAll |
            Message::Find | Message::Replace => {
                // TODO: Implement edit operations
            }

            // File menu
            Message::SaveProjectAs | Message::CloseProject | Message::Exit => {
                // TODO: Implement file operations
            }

            // Window menu
            Message::CascadeWindows | Message::TileHorizontal | Message::TileVertical => {
                // TODO: Implement window management
            }

            // Debug menu
            Message::StepInto | Message::StepOver => {
                // TODO: Implement debugging
            }

            Message::None => {}
        }

        Command::none()
    }

    pub fn view(&self) -> Element<Message> {
        // irys-style layout with menu, toolbar, and windows
        let menu = menu_bar();
        let toolbar_view = toolbar();

        let mut header = column![menu, toolbar_view].spacing(2);

        // Main content area with dockable windows
        let content: Element<Message> = match self.state.current_view {
            View::Designer => {
                let mut left_panel = row![].spacing(5);

                // Project Explorer (left side)
                if self.state.show_project_explorer {
                    left_panel = left_panel.push(project_explorer(&self.state));
                }

                // Toolbox (left side)
                if self.state.show_toolbox {
                    left_panel = left_panel.push(toolbox());
                }

                // Center: Designer
                let designer_view = self.designer.view(&self.state);

                // Right panel: Properties
                let mut right_panel = row![].spacing(5);
                if self.state.show_properties {
                    right_panel = right_panel.push(properties_panel(&self.state));
                }

                row![
                    left_panel,
                    designer_view,
                    right_panel,
                ]
                .spacing(5)
                .into()
            }
            View::Code => {
                let mut left_panel = row![].spacing(5);

                if self.state.show_project_explorer {
                    left_panel = left_panel.push(project_explorer(&self.state));
                }

                let code_view = code_editor_view(&self.state);

                let mut right_panel = row![].spacing(5);
                if self.state.show_properties {
                    right_panel = right_panel.push(properties_panel(&self.state));
                }

                row![
                    left_panel,
                    code_view,
                    right_panel,
                ]
                .spacing(5)
                .into()
            }
        };

        // Status bar
        let status_text = if self.state.is_running {
            "Running..."
        } else if self.state.project.is_some() {
            "Ready"
        } else {
            "No project loaded"
        };

        let status_bar = container(
            text(status_text).size(11)
        )
        .padding(3)
        .width(Length::Fill);

        let main_view = column![header, content, status_bar]
            .spacing(2);

        container(main_view)
            .width(Length::Fill)
            .height(Length::Fill)
            .padding(2)
            .into()
    }

    fn run_project(&self, project: &Project) {
        use irys_parser::parse_program;
        use irys_runtime::Interpreter;

        println!("\n=== Running Project: {} ===", project.name);

        // Get startup form
        let startup_form = match project.get_startup_form() {
            Some(form) => form,
            None => {
                println!("Error: No startup form defined");
                return;
            }
        };

        println!("Startup form: {}", startup_form.form.name);
        println!("Controls: {}", startup_form.form.controls.len());

        // Parse and run code
        let code_to_run = if startup_form.code.trim().is_empty() {
             ""
        } else {
             &startup_form.code
        };

        println!("\nParsing code...");
        match parse_program(code_to_run) {
            Ok(program) => {
                println!("Parse successful!");
                println!("Declarations: {}", program.declarations.len());
                println!("Statements: {}", program.statements.len());

                println!("\nExecuting...");
                let mut interpreter = Interpreter::new();

                // Register event handlers
                for binding in &startup_form.form.event_bindings {
                    interpreter.events.register(
                        &binding.control_name,
                        binding.event_type.clone(),
                        &binding.handler_name,
                    );
                }

                match interpreter.run(&program) {
                    Ok(_) => println!("Execution completed successfully!"),
                    Err(e) => println!("Runtime error: {}", e),
                }
            }
            Err(e) => {
                println!("Parse error: {}", e);
            }
        }

        println!("=== End ===\n");
    }
}

impl Default for irysEditorApp {
    fn default() -> Self {
        Self {
            state: EditorState::new(),
            designer: Designer::new(),
        }
    }
}

fn generate_form_code(form: &Form) -> String {
    let mut code = String::new();

    code.push_str(&format!("' Form: {}\n", form.name));
    code.push_str(&format!("' Caption: {}\n", form.caption));
    code.push_str(&format!("' Controls: {}\n\n", form.controls.len()));

    // Generate Form_Load event
    code.push_str(&format!("Private Sub {}_Load()\n", form.name));
    code.push_str("    ' Form load initialization\n");
    code.push_str("End Sub\n\n");

    // Generate event handler stubs for all controls
    for control in &form.controls {
        for binding in &form.event_bindings {
            if binding.control_name == control.name {
                code.push_str(&format!("Private Sub {}_{}()\n", control.name, binding.event_type.as_str()));
                code.push_str(&format!("    ' {} event handler for {}\n", binding.event_type.as_str(), control.name));
                code.push_str(&format!("    MsgBox(\"You clicked {}!\")\n", control.name));
                code.push_str("End Sub\n\n");
            }
        }
    }

    // Add example procedures
    code.push_str("' Example: Add your own procedures here\n");
    code.push_str("' Function MyFunction(param As String) As Integer\n");
    code.push_str("'     MsgBox param\n");
    code.push_str("'     MyFunction = 42\n");
    code.push_str("' End Function\n");

    code
}

impl Application for irysEditorApp {
    type Executor = iced::executor::Default;
    type Message = Message;
    type Theme = Theme;
    type Flags = ();

    fn new(_flags: Self::Flags) -> (Self, Command<Self::Message>) {
        Self::new()
    }

    fn title(&self) -> String {
        self.title()
    }

    fn update(&mut self, message: Self::Message) -> Command<Self::Message> {
        self.update(message)
    }

    fn view(&self) -> Element<Self::Message> {
        self.view()
    }
}
