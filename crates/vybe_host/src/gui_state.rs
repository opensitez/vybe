//! GUI state — owns the widget form and event handlers.
//!
//! This is the bridge between the `vybe:gui` host module and `vybe_widgets`.
//! Host functions directly create widgets and register events here.
//! The CLI just takes the finished form and runs it in a window.

use std::collections::HashMap;
use vybe_bytecode::Value;
use vybe_widgets::{
    Form as WidgetForm,
    Button, Label, TextInput, Checkbox, Radio, Select, ListBox,
    Panel, GroupBox, PictureBox, ProgressBar, Slider, NumericUpDown,
    DateTimePicker, ScrollBar, LinkLabel, MaskedTextBox,
    TreeView, DataGrid, ListView, Tabs, MonthCalendar,
    MenuStrip, ContextMenu, StatusStrip, ToolStrip,
    SplitContainer, FlowLayoutPanel, TableLayoutPanel,
    BindingNavigator, PanelWidget, WidgetCommand, CommandValue,
};

/// Holds the live widget form + event callbacks.
/// Created before VM runs, shared with host fns via `Rc<RefCell<>>`.
pub struct GuiState {
    /// The widget form being built by host functions.
    pub form: WidgetForm,
    /// Event handlers: "controlname.eventname" → VM callback Value.
    pub event_handlers: HashMap<String, Value>,
    /// Control names in insertion order.
    pub control_names: Vec<String>,
    /// Form dimensions (logical pixels, before scaling).
    pub width: u32,
    pub height: u32,
    /// Whether runApplication/showForm was called — signals the CLI to open a window.
    pub should_run: bool,
    /// The form object reference for `this`/`Me` dispatch.
    pub form_object: Option<Value>,
    /// Pending close request.
    pub close_requested: bool,
    /// Pending MsgBox dialogs: (text, title). Drained by the runner after each VM invocation.
    pub pending_dialogs: Vec<(String, String)>,
}

impl GuiState {
    pub fn new() -> Self {
        Self {
            form: WidgetForm::new("Form1"),
            event_handlers: HashMap::new(),
            control_names: Vec::new(),
            width: 800,
            height: 600,
            should_run: false,
            form_object: None,
            close_requested: false,
            pending_dialogs: Vec::new(),
        }
    }

    /// Register an event handler: key = "controlname.eventname" (control lowercased).
    pub fn register_event(&mut self, control: &str, event: &str, callback: Value) {
        let key = format!("{}.{}", control.to_lowercase(), event);
        self.event_handlers.insert(key, callback);
    }

    /// Look up an event handler.
    pub fn get_event_handler(&self, control: &str, event: &str) -> Option<&Value> {
        let key = format!("{}.{}", control.to_lowercase(), event.to_lowercase());
        self.event_handlers.get(&key)
    }

    /// List all registered event keys (for debugging).
    pub fn event_keys(&self) -> Vec<String> {
        self.event_handlers.keys().cloned().collect()
    }

    /// Create a widget from a control type name and add it to the form.
    pub fn add_widget(&mut self, type_name: &str, name: &str, text: &str, x: i32, y: i32, w: i32, h: i32) {
        let name_lower = name.to_lowercase();
        let widget = make_widget(type_name, &name_lower, text, w as f32, h as f32);
        self.form.add_boxed_control(widget, x as f32, y as f32, w as f32, h as f32);
        self.control_names.push(name_lower);
    }

    /// Set a property on a control by name — directly updates the widget.
    pub fn set_property(&mut self, control: &str, property: &str, value: &str) {
        let name = control.to_lowercase();
        match property.to_lowercase().as_str() {
            "text" => {
                self.form.send_command(&name, &WidgetCommand::SetText(value.to_string()));
            }
            "enabled" => {
                let enabled = !matches!(value, "false" | "False" | "0" | "");
                self.form.send_command(&name, &WidgetCommand::SetEnabled(enabled));
            }
            "visible" => {
                let visible = !matches!(value, "false" | "False" | "0" | "");
                self.form.send_command(&name, &WidgetCommand::SetVisible(visible));
            }
            "readonly" => {
                let ro = matches!(value, "true" | "True" | "1");
                self.form.send_command(&name, &WidgetCommand::Custom("SetReadOnly".into(), CommandValue::Bool(ro)));
            }
            other => {
                self.form.send_command(&name, &WidgetCommand::Custom(
                    format!("Set{}", capitalize_first(other)),
                    CommandValue::Text(value.to_string()),
                ));
            }
        }
    }

    /// Get a property from a control by name.
    pub fn get_property(&mut self, control: &str, property: &str) -> String {
        let name = control.to_lowercase();
        match property.to_lowercase().as_str() {
            "text" => {
                let result = self.form.send_command(&name, &WidgetCommand::GetText);
                match result {
                    CommandValue::Text(s) => s,
                    _ => String::new(),
                }
            }
            _ => {
                let result = self.form.send_command(&name, &WidgetCommand::Custom(
                    format!("Get{}", capitalize_first(property)),
                    CommandValue::None,
                ));
                match result {
                    CommandValue::Text(s) => s,
                    _ => String::new(),
                }
            }
        }
    }
}

fn capitalize_first(s: &str) -> String {
    let mut c = s.chars();
    match c.next() {
        None => String::new(),
        Some(f) => f.to_uppercase().collect::<String>() + c.as_str(),
    }
}

/// Create a boxed PanelWidget from a type name string.
fn make_widget(type_name: &str, name: &str, text: &str, w: f32, h: f32) -> Box<dyn PanelWidget> {
    match type_name.to_lowercase().as_str() {
        "button" => {
            let mut b = Button::new(text).with_name(name);
            b.width = w;
            b.height = h;
            Box::new(b)
        }
        "label" | "linklabel" => {
            let mut l = Label::new(text).with_name(name);
            l.width = w;
            l.height = h;
            Box::new(l)
        }
        "textbox" | "richtextbox" => {
            let mut t = TextInput::new().with_name(name);
            t.value = text.to_string();
            t.width = w;
            t.height = h;
            Box::new(t)
        }
        "maskedtextbox" => {
            let mut t = MaskedTextBox::new().with_name(name);
            t.value = text.to_string();
            Box::new(t)
        }
        "checkbox" => Box::new(Checkbox::new(text).with_name(name)),
        "radiobutton" => Box::new(Radio::new(text).with_name(name)),
        "combobox" => {
            let mut s = Select::new(vec![]).with_name(name);
            s.width = w;
            s.height = h;
            Box::new(s)
        }
        "listbox" | "checkedlistbox" => {
            let mut l = ListBox::new().with_name(name);
            l.width = w;
            l.height = h;
            Box::new(l)
        }
        "panel" | "usercontrol" => {
            let mut p = Panel::new().with_name(name);
            p.width = w;
            p.height = h;
            Box::new(p)
        }
        "groupbox" | "frame" => {
            let mut g = GroupBox::new(text).with_name(name);
            g.width = w;
            g.height = h;
            Box::new(g)
        }
        "picturebox" => {
            let mut p = PictureBox::new().with_name(name);
            p.width = w;
            p.height = h;
            Box::new(p)
        }
        "progressbar" => {
            let mut p = ProgressBar::new().with_name(name);
            p.width = w;
            p.height = h;
            Box::new(p)
        }
        "trackbar" => Box::new(Slider::new(0.0, 100.0, 50.0).with_name(name)),
        "numericupdown" => Box::new(NumericUpDown::new().with_name(name)),
        "datetimepicker" => Box::new(DateTimePicker::new().with_name(name)),
        "treeview" => Box::new(TreeView::new("", 1.0).with_name(name)),
        "datagridview" | "datagrid" => Box::new(DataGrid::new(&[]).with_name(name)),
        "listview" => Box::new(ListView::new().with_name(name)),
        "tabcontrol" => {
            let mut t = Tabs::new(&["Tab1"]).with_name(name);
            t.width = w;
            t.height = h;
            Box::new(t)
        }
        "monthcalendar" => Box::new(MonthCalendar::new().with_name(name)),
        "hscrollbar" => {
            let mut s = ScrollBar::new(false).with_name(name);
            s.width = w;
            s.height = h;
            Box::new(s)
        }
        "vscrollbar" => {
            let mut s = ScrollBar::new(true).with_name(name);
            s.width = w;
            s.height = h;
            Box::new(s)
        }
        "menustrip" => Box::new(MenuStrip::new().with_name(name)),
        "toolstrip" => Box::new(ToolStrip::new().with_name(name)),
        "statusstrip" => Box::new(StatusStrip::new().with_name(name)),
        "contextmenustrip" => Box::new(ContextMenu::new().with_name(name)),
        "splitcontainer" => Box::new(SplitContainer::new(false).with_name(name)),
        "flowlayoutpanel" => Box::new(FlowLayoutPanel::new().with_name(name)),
        "tablelayoutpanel" => Box::new(TableLayoutPanel::new(2, 2).with_name(name)),
        "bindingnavigator" => Box::new(BindingNavigator::new(name)),
        _ => {
            // Unknown control type — render as label placeholder
            let mut l = Label::new(&format!("[{}]", name)).with_name(name);
            l.transparent = true;
            l.width = w;
            l.height = h;
            Box::new(l)
        }
    }
}
