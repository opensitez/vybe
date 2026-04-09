//! VybeWidget-based form renderer.
//!
//! Uses `vybe_widgets::Form` as the container for all controls, and
//! `vybe_widgets::Application` + `run_app()` for the window/event loop.
//! The CLI owns *only* the VM glue (event dispatch, data binding, dialogs).
//! All graphics, focus management, hover states, and keyboard routing live
//! in vybe_widgets.

use std::rc::Rc;
use std::cell::RefCell;
use std::sync::{Arc, Mutex};
use vybe_host::GuiState;

use vybe_widgets::{
    // Application framework
    Application, run_app, Pixmap, fill_background,
    FontSystem, SwashCache, RenderContext,
    // Layout types
    LayoutRect, MouseEvent, KeyEvent,
    // Widget trait + events/commands
    PanelWidget, WidgetEvent, WidgetCommand, CommandValue,
    // Form container
    Form as WidgetForm,
    // Widgets
    Button, Label, TextInput, Checkbox, Radio, ListBox, ProgressBar,
    Slider, NumericUpDown, DateTimePicker, ScrollBar, LinkLabel, MaskedTextBox,
    GroupBox, Panel, TreeView, DataGrid, ListView,
    Tabs, MenuStrip, ContextMenu, StatusStrip, ToolStrip,
    SplitContainer, FlowLayoutPanel, TableLayoutPanel, MonthCalendar,
    PictureBox, Select, BindingNavigator,
};

// ── Data binding types ─────────────────────────────────────────────────

#[derive(Clone, Debug)]
struct DataBindingEntry {
    control_name: String,
    property: String,
    source_name: String,
    column: String,
}

#[derive(Clone, Debug)]
struct BindingSourceInfo {
    name: String,
    data_adapter_name: String,
    data_member: String,
}

#[derive(Clone, Debug)]
struct NavigatorInfo {
    navigator_name: String,
    binding_source_name: String,
}

struct DataStore {
    columns: Vec<String>,
    rows: Vec<std::collections::HashMap<String, String>>,
    position: i32,
}

// ── Control type → Widget mapping ──────────────────────────────────────

/// Convert a `vybe_forms::Control` into a boxed `PanelWidget`.
fn make_widget(ctrl: &vybe_forms::Control) -> Box<dyn PanelWidget> {
    let text = ctrl.properties.get_string("Text").unwrap_or_default().to_string();
    let name = ctrl.name.to_lowercase();

    match ctrl.control_type {
        vybe_forms::ControlType::Button => {
            let mut w = Button::new(&text).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_forms::ControlType::Label | vybe_forms::ControlType::LinkLabel => {
            let mut w = Label::new(&text).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_forms::ControlType::TextBox | vybe_forms::ControlType::RichTextBox => {
            let mut w = TextInput::new().with_name(&name);
            w.value = text;
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_forms::ControlType::MaskedTextBox => {
            let mut w = MaskedTextBox::new().with_name(&name);
            w.value = text;
            Box::new(w)
        }
        vybe_forms::ControlType::CheckBox => {
            Box::new(Checkbox::new(&text).with_name(&name))
        }
        vybe_forms::ControlType::RadioButton => {
            Box::new(Radio::new(&text).with_name(&name))
        }
        vybe_forms::ControlType::ComboBox => {
            let items = ctrl.properties.get_string_array("Items").cloned().unwrap_or_default();
            let mut w = Select::new(items).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_forms::ControlType::ListBox | vybe_forms::ControlType::CheckedListBox => {
            let items = ctrl.properties.get_string_array("Items").cloned().unwrap_or_default();
            let mut w = ListBox::new().with_name(&name);
            w.items = items;
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_forms::ControlType::Panel | vybe_forms::ControlType::UserControl => {
            let mut w = Panel::new().with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_forms::ControlType::Frame => {
            let mut w = GroupBox::new(&text).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_forms::ControlType::PictureBox => {
            let mut w = PictureBox::new().with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_forms::ControlType::ProgressBar => {
            let mut w = ProgressBar::new().with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_forms::ControlType::TrackBar => {
            Box::new(Slider::new(0.0, 100.0, 50.0).with_name(&name))
        }
        vybe_forms::ControlType::NumericUpDown => {
            Box::new(NumericUpDown::new().with_name(&name))
        }
        vybe_forms::ControlType::DateTimePicker => {
            Box::new(DateTimePicker::new().with_name(&name))
        }
        vybe_forms::ControlType::TreeView => {
            Box::new(TreeView::new("", 1.0).with_name(&name))
        }
        vybe_forms::ControlType::DataGridView | vybe_forms::ControlType::DataGrid => {
            Box::new(DataGrid::new(&[]).with_name(&name))
        }
        vybe_forms::ControlType::ListView => {
            Box::new(ListView::new().with_name(&name))
        }
        vybe_forms::ControlType::TabControl => {
            let mut w = Tabs::new(&["Tab1"]).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_forms::ControlType::MonthCalendar => {
            Box::new(MonthCalendar::new().with_name(&name))
        }
        vybe_forms::ControlType::HScrollBar => {
            let mut w = ScrollBar::new(false).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_forms::ControlType::VScrollBar => {
            let mut w = ScrollBar::new(true).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_forms::ControlType::MenuStrip => {
            Box::new(MenuStrip::new().with_name(&name))
        }
        vybe_forms::ControlType::ToolStrip => {
            Box::new(ToolStrip::new().with_name(&name))
        }
        vybe_forms::ControlType::StatusStrip => {
            Box::new(StatusStrip::new().with_name(&name))
        }
        vybe_forms::ControlType::ContextMenuStrip => {
            Box::new(ContextMenu::new().with_name(&name))
        }
        vybe_forms::ControlType::SplitContainer => {
            Box::new(SplitContainer::new(false).with_name(&name))
        }
        vybe_forms::ControlType::FlowLayoutPanel => {
            Box::new(FlowLayoutPanel::new().with_name(&name))
        }
        vybe_forms::ControlType::TableLayoutPanel => {
            Box::new(TableLayoutPanel::new(2, 2).with_name(&name))
        }
        vybe_forms::ControlType::BindingNavigator => {
            Box::new(BindingNavigator::new(&name))
        }
        _ => {
            let mut w = Label::new(&format!("[{}]", ctrl.name));
            w.transparent = true;
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
    }
}

// ── FormApp — Application impl ─────────────────────────────────────────

struct FormApp {
    /// Control names in insertion order (index matches form control index).
    font_system: FontSystem,
    swash_cache: SwashCache,
    vm: Rc<RefCell<vybe_bytecode::VM>>,
    gui: Arc<Mutex<GuiState>>,
    // Data binding state
    data_bindings: Vec<DataBindingEntry>,
    binding_sources: Vec<BindingSourceInfo>,
    navigators: Vec<NavigatorInfo>,
    data_store: std::collections::HashMap<String, DataStore>,
    initialised: bool,
}

impl Application for FormApp {
    fn on_init(&mut self, width: f32, height: f32, _scale: f32) {
        self.gui.lock().unwrap().form.set_rect(LayoutRect::new(0.0, 0.0, width, height));
        if !self.initialised {
            self.initialised = true;
            self.fire_load_event();
            self.init_data_bindings();
        }
    }

    fn on_resize(&mut self, width: f32, height: f32) {
        self.gui.lock().unwrap().form.set_rect(LayoutRect::new(0.0, 0.0, width, height));
    }

    fn render(&mut self, pixmap: &mut Pixmap, scale: f32) {
        fill_background(pixmap, 240, 240, 240, 255);
        let mut g = self.gui.lock().unwrap();
        let mut ctx = RenderContext {
            pixmap,
            font_system: &mut self.font_system,
            swash_cache: &mut self.swash_cache,
            scale,
        };
        g.form.render(&mut ctx);
    }

    fn handle_mouse(&mut self, event: MouseEvent) -> bool {
        self.gui.lock().unwrap().form.handle_mouse(&event);
        self.process_widget_events();
        true
    }

    fn handle_key(&mut self, event: KeyEvent) -> bool {
        self.gui.lock().unwrap().form.handle_key(&event);
        self.process_widget_events();
        true
    }

    fn handle_scroll(&mut self, delta: f32, x: f32, y: f32) -> bool {
        self.gui.lock().unwrap().form.handle_scroll(delta, x, y)
    }

    fn cursor_icon(&self) -> vybe_widgets::CursorIcon {
        vybe_widgets::CursorIcon::Default
    }
}

// ── VM glue ────────────────────────────────────────────────────────────

impl FormApp {
    fn fire_load_event(&mut self) {
        let callback = {
            let g = self.gui.lock().unwrap();
            g.get_event_handler("form1", "Load").cloned()
                .or_else(|| g.get_event_handler("me", "Load").cloned())
        };
        if let Some(cb) = callback {
            let mut vm = self.vm.borrow_mut();
            let me = vm.globals.get("__f").cloned()
                .unwrap_or(vybe_bytecode::Value::Null);
            let arity = fn_arity(&cb);
            let result = match arity {
                0 => vm.invoke(&cb, &[]),
                1 => vm.invoke(&cb, &[me]),
                _ => vm.invoke(&cb, &[me, vybe_bytecode::Value::Null, vybe_bytecode::Value::Null]),
            };
            if let Err(e) = result {
                eprintln!("[LOAD] Error: {e}");
            }
            drop(vm);
            self.drain_side_effects();
            self.sync_widgets_from_vm();
        }
    }

    fn fire_click(&mut self, control_name: &str) {
        let callback = {
            let g = self.gui.lock().unwrap();
            g.get_event_handler(&control_name.to_lowercase(), "Click").cloned()
        };
        if let Some(cb) = callback {
            self.invoke_callback(&cb, control_name);
        }
    }

    fn invoke_callback(&mut self, cb: &vybe_bytecode::Value, control_name: &str) {
        let mut vm = self.vm.borrow_mut();
        let me = vm.globals.get("__f").cloned()
            .unwrap_or(vybe_bytecode::Value::Null);
        let arity = fn_arity(cb);
        let sender = vybe_bytecode::Value::String(Arc::from(control_name));
        let result = match arity {
            0 => vm.invoke(cb, &[]),
            1 => vm.invoke(cb, &[me]),
            2 => vm.invoke(cb, &[me, sender]),
            _ => vm.invoke(cb, &[me, sender, vybe_bytecode::Value::Null]),
        };
        if let Err(e) = result {
            eprintln!("Event handler error: {e}");
        }
        drop(vm);
        self.drain_side_effects();
        self.sync_widgets_from_vm();
    }

    fn drain_side_effects(&mut self) {
        let dialogs: Vec<(String, String)> = self.gui.lock().unwrap().pending_dialogs.drain(..).collect();
        for (text, title) in dialogs {
            rfd::MessageDialog::new()
                .set_title(&title).set_description(&text)
                .set_level(rfd::MessageLevel::Info).show();
        }
    }

    /// Push VM object properties into form widgets via `send_command`.
    /// This handles the C# pattern where compiled code sets properties on
    /// VM objects (e.g. `this.txtDisplay.Text = "hello"`) and we need to
    /// push those values to the actual widgets.
    fn sync_widgets_from_vm(&mut self) {
        let updates = {
            let vm = self.vm.borrow();
            let g = self.gui.lock().unwrap();
            let mut ups: Vec<(String, String)> = Vec::new();
            if let Some(vybe_bytecode::Value::Object(form_obj)) = vm.globals.get("__f") {
                let fo = form_obj.lock().unwrap();
                for ctrl_name in &g.control_names {
                    if let Some(vybe_bytecode::Value::Object(co)) = fo.properties.get(ctrl_name) {
                        let c = co.lock().unwrap();
                        if let Some(text) = c.properties.get("text") {
                            ups.push((ctrl_name.clone(), format!("{}", text)));
                        }
                    }
                }
            }
            ups
        };
        if !updates.is_empty() {
            let mut g = self.gui.lock().unwrap();
            for (name, text) in updates {
                g.form.send_command(&name, &WidgetCommand::SetText(text));
            }
        }
    }

    /// Drain all widget events and map them to VM callbacks.
    fn process_widget_events(&mut self) {
        let events = self.gui.lock().unwrap().form.drain_events();
        for event in events {
            match &event {
                WidgetEvent::ButtonClicked(name) |
                WidgetEvent::CheckboxToggled(name, _) |
                WidgetEvent::RadioSelected(name, _) |
                WidgetEvent::TextChanged(name, _) |
                WidgetEvent::LinkClicked(name) => {
                    self.fire_click(name);
                }
                WidgetEvent::SelectChanged(name, _) |
                WidgetEvent::ListBoxSelected(name, _) => {
                    let callback = {
                        let g = self.gui.lock().unwrap();
                        g.get_event_handler(&name.to_lowercase(), "SelectedIndexChanged").cloned()
                            .or_else(|| g.get_event_handler(&name.to_lowercase(), "Click").cloned())
                    };
                    if let Some(cb) = callback {
                        self.invoke_callback(&cb, name);
                    }
                }
                WidgetEvent::Action(action_str) => {
                    if action_str.starts_with("nav:") {
                        let parts: Vec<&str> = action_str.splitn(3, ':').collect();
                        if parts.len() == 3 {
                            let nav_name = parts[1];
                            let action = parts[2];
                            if let Some(nav_info) = self.navigators.iter()
                                .find(|n| n.navigator_name.eq_ignore_ascii_case(nav_name))
                            {
                                let bs_name = nav_info.binding_source_name.clone();
                                self.navigate_binding_source(&bs_name, action);
                            }
                        }
                    }
                }
                _ => {}
            }
        }
    }

    // ── Data binding ───────────────────────────────────────────────────

    fn init_data_bindings(&mut self) {
        if self.binding_sources.is_empty() { return; }

        let bs_infos: Vec<_> = self.binding_sources.clone();
        for bs_info in &bs_infos {
            let conn_str = self.get_connection_string(&bs_info.name, &bs_info.data_adapter_name);
            if conn_str.is_empty() { continue; }

            let sql = format!("SELECT * FROM {}", bs_info.data_member);
            match vybe_host::modules::database::query_rows(&conn_str, &sql) {
                Ok((columns, rows)) => {
                    let store = DataStore {
                        columns: columns.clone(),
                        rows: rows.clone(),
                        position: if rows.is_empty() { -1 } else { 0 },
                    };
                    self.data_store.insert(bs_info.name.to_lowercase(), store);
                    self.sync_bound_controls(&bs_info.name);
                }
                Err(e) => {
                    eprintln!("[DATA] Query error for '{}': {}", bs_info.name, e);
                    self.data_store.insert(bs_info.name.to_lowercase(), DataStore {
                        columns: Vec::new(),
                        rows: Vec::new(),
                        position: -1,
                    });
                }
            }
        }
        self.update_navigator_positions();
    }

    fn get_connection_string(&self, bs_name: &str, adapter_name: &str) -> String {
        let vm = self.vm.borrow();
        if let Some(vybe_bytecode::Value::Object(form_obj)) = vm.globals.get("__f") {
            let fo = form_obj.lock().unwrap();
            if let Some(vybe_bytecode::Value::Object(bs_obj)) = fo.properties.get(&bs_name.to_lowercase()) {
                let bs = bs_obj.lock().unwrap();
                if let Some(vybe_bytecode::Value::Object(da_obj)) = bs.properties.get("datasource") {
                    let da = da_obj.lock().unwrap();
                    if let Some(v) = da.properties.get("connectionstring") {
                        return format!("{}", v);
                    }
                }
            }
            if let Some(vybe_bytecode::Value::Object(da_obj)) = fo.properties.get(&adapter_name.to_lowercase()) {
                let da = da_obj.lock().unwrap();
                if let Some(v) = da.properties.get("connectionstring") {
                    return format!("{}", v);
                }
            }
        }
        String::new()
    }

    fn sync_bound_controls(&mut self, bs_name: &str) {
        let bs_lower = bs_name.to_lowercase();
        let store = match self.data_store.get(&bs_lower) {
            Some(s) => s,
            None => return,
        };
        if store.position < 0 || store.position as usize >= store.rows.len() { return; }
        let row = &store.rows[store.position as usize];

        let vm = self.vm.borrow_mut();
        if let Some(vybe_bytecode::Value::Object(form_obj)) = vm.globals.get("__f") {
            let fo = form_obj.lock().unwrap();
            for binding in &self.data_bindings {
                if !binding.source_name.eq_ignore_ascii_case(bs_name) { continue; }
                let col_key = row.keys()
                    .find(|k| k.eq_ignore_ascii_case(&binding.column))
                    .cloned();
                let value = col_key.and_then(|k| row.get(&k)).cloned().unwrap_or_default();
                let ctrl_lower = binding.control_name.to_lowercase();
                if let Some(vybe_bytecode::Value::Object(ctrl_obj)) = fo.properties.get(&ctrl_lower) {
                    ctrl_obj.lock().unwrap().properties.insert(
                        binding.property.to_lowercase(),
                        vybe_bytecode::Value::String(Arc::from(value.as_str())),
                    );
                }
            }
        }
        drop(vm);
        self.sync_widgets_from_vm();
    }

    fn update_navigator_positions(&mut self) {
        let mut g = self.gui.lock().unwrap();
        for nav_info in &self.navigators {
            if let Some(store) = self.data_store.get(&nav_info.binding_source_name.to_lowercase()) {
                let pos_count = format!("{},{}", store.position, store.rows.len());
                g.form.send_command(
                    &nav_info.navigator_name.to_lowercase(),
                    &WidgetCommand::Custom("set_position_and_count".into(), CommandValue::Text(pos_count)),
                );
            }
        }
    }

    fn navigate_binding_source(&mut self, bs_name: &str, action: &str) {
        let bs_lower = bs_name.to_lowercase();
        let new_pos = {
            let store = match self.data_store.get(&bs_lower) {
                Some(s) => s,
                None => return,
            };
            let count = store.rows.len() as i32;
            if count == 0 { return; }
            match action {
                "first" => 0,
                "prev" => (store.position - 1).max(0),
                "next" => (store.position + 1).min(count - 1),
                "last" => count - 1,
                _ => store.position,
            }
        };
        if let Some(store) = self.data_store.get_mut(&bs_lower) {
            store.position = new_pos;
        }
        self.sync_bound_controls(bs_name);
        self.update_navigator_positions();
    }
}

// ── Helpers ────────────────────────────────────────────────────────────

fn fn_arity(val: &vybe_bytecode::Value) -> usize {
    match val {
        vybe_bytecode::Value::Object(obj) => {
            match &obj.lock().unwrap().kind {
                vybe_bytecode::value::ObjectKind::Function(f) => f.arity as usize,
                _ => 0,
            }
        }
        _ => 0,
    }
}

// ── Dialog registration ────────────────────────────────────────────────

fn register_dialog_fns(vm: &mut vybe_bytecode::VM) {
    use vybe_bytecode::Value;
    use vybe_bytecode::value::{Object, ObjectKind};
    use std::sync::{Arc, Mutex};

    vm.register_host_fn("vybe:gui", "__dlg_show", Box::new(|_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
        let dialog_type = if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            o.properties.get("__control_type").map(|v| format!("{}", v)).unwrap_or_default()
        } else { String::new() };

        match dialog_type.as_str() {
            "OpenFileDialog" => {
                let result = rfd::FileDialog::new().set_title("Open File").pick_file();
                if let Some(path) = result {
                    if let Some(Value::Object(obj)) = args.first() {
                        obj.lock().unwrap().properties.insert("filename".into(),
                            Value::String(Arc::from(path.to_string_lossy().as_ref())));
                    }
                    Value::I32(1)
                } else { Value::I32(0) }
            }
            "SaveFileDialog" => {
                let result = rfd::FileDialog::new().set_title("Save File").save_file();
                if let Some(path) = result {
                    if let Some(Value::Object(obj)) = args.first() {
                        obj.lock().unwrap().properties.insert("filename".into(),
                            Value::String(Arc::from(path.to_string_lossy().as_ref())));
                    }
                    Value::I32(1)
                } else { Value::I32(0) }
            }
            "FolderBrowserDialog" => {
                let result = rfd::FileDialog::new().set_title("Select Folder").pick_folder();
                if let Some(path) = result {
                    if let Some(Value::Object(obj)) = args.first() {
                        obj.lock().unwrap().properties.insert("selectedpath".into(),
                            Value::String(Arc::from(path.to_string_lossy().as_ref())));
                    }
                    Value::I32(1)
                } else { Value::I32(0) }
            }
            "ColorDialog" | "FontDialog" => Value::I32(1),
            _ => Value::I32(0),
        }
    }));

    vm.register_host_fn("vybe:gui", "msgBox", Box::new(|_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
        let text = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let title = args.get(1).map(|v| format!("{}", v)).unwrap_or_else(|| "Message".into());
        rfd::MessageDialog::new()
            .set_title(&title).set_description(&text)
            .set_level(rfd::MessageLevel::Info).show();
        Value::Null
    }));

    vm.register_host_fn("vybe:gui", "inputBox", Box::new(|_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
        let default = args.get(2).map(|v| format!("{}", v)).unwrap_or_default();
        Value::String(Arc::from(default.as_str()))
    }));

    let dlg_show_idx = *vm.host_registry.get(&("vybe:gui".into(), "__dlg_show".into())).unwrap();
    let dlg_show_ref = {
        let mut o = Object::new();
        o.kind = ObjectKind::HostFunction(dlg_show_idx);
        Value::Object(Arc::new(Mutex::new(o)))
    };
    vm.globals.insert("__dlg_show_ref".into(), dlg_show_ref);
}

// ── Extract binding info from form definition ──────────────────────────

fn extract_binding_info(form: &vybe_forms::Form) -> (Vec<DataBindingEntry>, Vec<BindingSourceInfo>, Vec<NavigatorInfo>) {
    let mut data_bindings = Vec::new();
    let mut binding_sources = Vec::new();
    let mut navigators = Vec::new();

    for ctrl in &form.controls {
        let type_name = format!("{:?}", ctrl.control_type);

        if type_name.contains("BindingSource") {
            let data_source = ctrl.properties.get_string("DataSource").unwrap_or_default().to_string();
            let data_member = ctrl.properties.get_string("DataMember").unwrap_or_default().to_string();
            if !data_source.is_empty() && !data_member.is_empty() {
                binding_sources.push(BindingSourceInfo {
                    name: ctrl.name.clone(),
                    data_adapter_name: data_source,
                    data_member,
                });
            }
        }

        if type_name.contains("BindingNavigator") {
            let bs = ctrl.properties.get_string("BindingSource").unwrap_or_default().to_string();
            if !bs.is_empty() {
                navigators.push(NavigatorInfo {
                    navigator_name: ctrl.name.clone(),
                    binding_source_name: bs,
                });
            }
        }

        let binding_source = ctrl.properties.get_string("DataBindings.Source").map(|s| s.to_string());
        if let Some(ref bs_name) = binding_source {
            if !bs_name.is_empty() {
                for (key, val) in ctrl.properties.iter() {
                    let k = key.as_str();
                    if k.starts_with("DataBindings.") && k != "DataBindings.Source" {
                        let prop = &k["DataBindings.".len()..];
                        if let Some(column) = val.as_string() {
                            if !column.is_empty() {
                                data_bindings.push(DataBindingEntry {
                                    control_name: ctrl.name.clone(),
                                    property: prop.to_string(),
                                    source_name: bs_name.clone(),
                                    column: column.to_string(),
                                });
                            }
                        }
                    }
                }
            }
        }
    }

    (data_bindings, binding_sources, navigators)
}

// ── Public launch function ─────────────────────────────────────────────

pub fn launch_vybewidget_form(
    mut vm: vybe_bytecode::VM,
    gui: Arc<Mutex<GuiState>>,
    form: &vybe_forms::Form,
) {
    register_dialog_fns(&mut vm);

    // Set __f so event handlers have a `this`/`Me` reference
    if let Some(form_obj) = gui.lock().unwrap().form_object.clone() {
        vm.globals.insert("__f".into(), form_obj);
    }

    // Always rebuild widgets from the form model when it has controls.
    // The form model is extracted from the same AST and has correct positions,
    // sizes, text, nesting, colors. The VM's controlsAdd creates widgets too
    // early (before properties are set in .NET designer code order), producing
    // widgets with wrong names, zero positions, and default sizes.
    //
    // vybe_widgets is a flat arena — all controls live at the root level with
    // absolute positions. The form model has parent_id for nesting, so we
    // convert relative child positions to absolute by walking up the parent chain.
    let model_control_count = form.controls.iter().filter(|c| !c.control_type.is_non_visual()).count();
    if model_control_count > 0 {
        // Build a lookup: control id → (x, y) for absolute position computation
        let id_to_bounds: std::collections::HashMap<_, _> =
            form.controls.iter().map(|c| (c.id, &c.bounds)).collect();

        let mut g = gui.lock().unwrap();
        g.form = WidgetForm::new(&form.text);
        g.control_names.clear();
        for ctrl in &form.controls {
            if ctrl.control_type.is_non_visual() { continue; }
            let widget = make_widget(ctrl);

            // Compute absolute position by walking up parent chain
            let mut abs_x = ctrl.bounds.x;
            let mut abs_y = ctrl.bounds.y;
            let mut parent = ctrl.parent_id;
            while let Some(pid) = parent {
                if let Some(pb) = id_to_bounds.get(&pid) {
                    abs_x += pb.x;
                    abs_y += pb.y;
                }
                // Walk further up
                parent = form.controls.iter().find(|c| c.id == pid).and_then(|c| c.parent_id);
            }

            g.form.add_boxed_control(widget, abs_x as f32, abs_y as f32, ctrl.bounds.width as f32, ctrl.bounds.height as f32);
            g.control_names.push(ctrl.name.to_lowercase());
        }
    }

    // Debug dump all widget state before rendering
    gui.lock().unwrap().form.debug_dump();

    let (data_bindings, binding_sources, navigators) = extract_binding_info(form);

    let app = FormApp {
        font_system: FontSystem::new(),
        swash_cache: SwashCache::new(),
        vm: Rc::new(RefCell::new(vm)),
        gui,
        data_bindings,
        binding_sources,
        navigators,
        data_store: std::collections::HashMap::new(),
        initialised: false,
    };

    run_app(
        &form.text,
        form.width as u32,
        form.height as u32,
        1.0,
        app,
    );
}

/// Launch a programmatic form — GuiState already has all widgets and event handlers.
pub fn launch_gui(
    mut vm: vybe_bytecode::VM,
    gui: Arc<Mutex<GuiState>>,
) {
    register_dialog_fns(&mut vm);

    let (title, width, height) = {
        let g = gui.lock().unwrap();
        ("Form1".to_string(), g.width, g.height)
    };

    // Set __f so event handlers have a `this`/`Me` reference
    if let Some(form_obj) = gui.lock().unwrap().form_object.clone() {
        vm.globals.insert("__f".into(), form_obj);
    }

    // GuiState keeps its form — no mem::replace, no stealing.
    let app = FormApp {
        font_system: FontSystem::new(),
        swash_cache: SwashCache::new(),
        vm: Rc::new(RefCell::new(vm)),
        gui,
        data_bindings: Vec::new(),
        binding_sources: Vec::new(),
        navigators: Vec::new(),
        data_store: std::collections::HashMap::new(),
        initialised: false,
    };

    run_app(&title, width, height, 1.0, app);
}
