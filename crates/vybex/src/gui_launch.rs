//! GUI launch layer — owns the winit event loop, VM glue, and dialog registration.
//!
//! Uses `vybe_widgets::Form` as the container for all controls, and
//! `vybe_widgets::Application` + `run_app()` for the window/event loop.
//! All graphics, focus management, hover states, and keyboard routing live
//! in vybe_widgets.
//!
//! Two entry points:
//! - `launch_gui` — programmatic forms (GuiState already has widgets)
//! - `launch_vybewidget_form` — designer forms (builds widgets from `vybe_compiler::projects::vbforms::Form`)
//!   (requires `gui_forms` feature)

use std::cell::RefCell;
use std::rc::Rc;
use std::sync::{Arc, Mutex};
use vybe_host::GuiState;

use vybe_widgets::{
    // Application framework
    Application,
    CommandValue,
    FontSystem,
    KeyEvent,
    // Layout types
    LayoutRect,
    MouseEvent,
    // Widget trait + events/commands
    PanelWidget,
    Pixmap,
    RenderContext,
    SwashCache,
    WidgetCommand,
    WidgetEvent,
    fill_background,
    run_app,
};

#[cfg(feature = "gui_forms")]
use vybe_widgets::{
    BindingNavigator, Button, Checkbox, ContextMenu, DataGrid, DateTimePicker, FlowLayoutPanel,
    Form as WidgetForm, GroupBox, Label, ListBox, ListView, MaskedTextBox, MenuStrip,
    MonthCalendar, NumericUpDown, Panel, PictureBox, ProgressBar, Radio, ScrollBar, Select, Slider,
    SplitContainer, StatusStrip, TableLayoutPanel, Tabs, TextInput, ToolStrip, TreeView,
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

#[allow(dead_code)]
struct DataStore {
    columns: Vec<String>,
    rows: Vec<std::collections::HashMap<String, String>>,
    position: i32,
}

// ── Control type → Widget mapping (designer forms only) ────────────────

#[cfg(feature = "gui_forms")]
fn make_widget(ctrl: &vybe_compiler::projects::vbforms::Control) -> Box<dyn PanelWidget> {
    let text = ctrl
        .properties
        .get_string("Text")
        .unwrap_or_default()
        .to_string();
    // Preserve original case — VB compiler will lowercase identifiers as needed
    let name = &ctrl.name;

    match ctrl.control_type {
        vybe_compiler::projects::vbforms::ControlType::Button => {
            let mut w = Button::new(&text).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_compiler::projects::vbforms::ControlType::Label
        | vybe_compiler::projects::vbforms::ControlType::LinkLabel => {
            let mut w = Label::new(&text).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_compiler::projects::vbforms::ControlType::TextBox
        | vybe_compiler::projects::vbforms::ControlType::RichTextBox => {
            let mut w = TextInput::new().with_name(&name);
            w.value = text;
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_compiler::projects::vbforms::ControlType::MaskedTextBox => {
            let mut w = MaskedTextBox::new().with_name(&name);
            w.value = text;
            Box::new(w)
        }
        vybe_compiler::projects::vbforms::ControlType::CheckBox => {
            Box::new(Checkbox::new(&text).with_name(&name))
        }
        vybe_compiler::projects::vbforms::ControlType::RadioButton => {
            Box::new(Radio::new(&text).with_name(&name))
        }
        vybe_compiler::projects::vbforms::ControlType::ComboBox => {
            let items = ctrl
                .properties
                .get_string_array("Items")
                .cloned()
                .unwrap_or_default();
            let mut w = Select::new(items).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_compiler::projects::vbforms::ControlType::ListBox
        | vybe_compiler::projects::vbforms::ControlType::CheckedListBox => {
            let items = ctrl
                .properties
                .get_string_array("Items")
                .cloned()
                .unwrap_or_default();
            let mut w = ListBox::new().with_name(&name);
            w.items = items;
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_compiler::projects::vbforms::ControlType::Panel
        | vybe_compiler::projects::vbforms::ControlType::UserControl => {
            let mut w = Panel::new().with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_compiler::projects::vbforms::ControlType::Frame => {
            let mut w = GroupBox::new(&text).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_compiler::projects::vbforms::ControlType::PictureBox => {
            let mut w = PictureBox::new().with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_compiler::projects::vbforms::ControlType::ProgressBar => {
            let mut w = ProgressBar::new().with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_compiler::projects::vbforms::ControlType::TrackBar => {
            Box::new(Slider::new(0.0, 100.0, 50.0).with_name(&name))
        }
        vybe_compiler::projects::vbforms::ControlType::NumericUpDown => {
            Box::new(NumericUpDown::new().with_name(&name))
        }
        vybe_compiler::projects::vbforms::ControlType::DateTimePicker => {
            Box::new(DateTimePicker::new().with_name(&name))
        }
        vybe_compiler::projects::vbforms::ControlType::TreeView => {
            Box::new(TreeView::new("", 1.0).with_name(&name))
        }
        vybe_compiler::projects::vbforms::ControlType::DataGridView
        | vybe_compiler::projects::vbforms::ControlType::DataGrid => {
            Box::new(DataGrid::new(&[]).with_name(&name))
        }
        vybe_compiler::projects::vbforms::ControlType::ListView => {
            Box::new(ListView::new().with_name(&name))
        }
        vybe_compiler::projects::vbforms::ControlType::TabControl => {
            let mut w = Tabs::new(&["Tab1"]).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_compiler::projects::vbforms::ControlType::MonthCalendar => {
            Box::new(MonthCalendar::new().with_name(&name))
        }
        vybe_compiler::projects::vbforms::ControlType::HScrollBar => {
            let mut w = ScrollBar::new(false).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_compiler::projects::vbforms::ControlType::VScrollBar => {
            let mut w = ScrollBar::new(true).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_compiler::projects::vbforms::ControlType::MenuStrip => {
            Box::new(MenuStrip::new().with_name(&name))
        }
        vybe_compiler::projects::vbforms::ControlType::ToolStrip => {
            Box::new(ToolStrip::new().with_name(&name))
        }
        vybe_compiler::projects::vbforms::ControlType::StatusStrip => {
            Box::new(StatusStrip::new().with_name(&name))
        }
        vybe_compiler::projects::vbforms::ControlType::ContextMenuStrip => {
            Box::new(ContextMenu::new().with_name(&name))
        }
        vybe_compiler::projects::vbforms::ControlType::SplitContainer => {
            Box::new(SplitContainer::new(false).with_name(&name))
        }
        vybe_compiler::projects::vbforms::ControlType::FlowLayoutPanel => {
            Box::new(FlowLayoutPanel::new().with_name(&name))
        }
        vybe_compiler::projects::vbforms::ControlType::TableLayoutPanel => {
            Box::new(TableLayoutPanel::new(2, 2).with_name(&name))
        }
        vybe_compiler::projects::vbforms::ControlType::BindingNavigator => {
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
        self.gui
            .lock()
            .unwrap()
            .form
            .set_rect(LayoutRect::new(0.0, 0.0, width, height));
        if !self.initialised {
            self.initialised = true;
            self.fire_load_event();
            self.init_data_bindings();
        }
    }

    fn on_resize(&mut self, width: f32, height: f32) {
        self.gui
            .lock()
            .unwrap()
            .form
            .set_rect(LayoutRect::new(0.0, 0.0, width, height));
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
        if Self::gui_trace_enabled() {
            eprintln!("[gui] formapp.handle_mouse event={:?}", event);
        }
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
    fn gui_trace_enabled() -> bool {
        std::env::var("VYBE_GUI_TRACE")
            .map(|value| !matches!(value.as_str(), "" | "0" | "false" | "False"))
            .unwrap_or(false)
    }

    fn fire_load_event(&mut self) {
        let callback = {
            let g = self.gui.lock().unwrap();
            g.get_event_handler("form1", "Load")
                .cloned()
                .or_else(|| g.get_event_handler("me", "Load").cloned())
        };
        if let Some(cb) = callback {
            let mut vm = self.vm.borrow_mut();
            let me = vm
                .globals
                .get("__f")
                .cloned()
                .unwrap_or(vybe_bytecode::Value::Null);
            let arity = fn_arity(&cb);
            let result = match arity {
                0 => vm.invoke(&cb, &[]),
                1 => vm.invoke(&cb, &[me]),
                _ => vm.invoke(
                    &cb,
                    &[me, vybe_bytecode::Value::Null, vybe_bytecode::Value::Null],
                ),
            };
            if let Err(e) = result {
                eprintln!("[LOAD] Error: {e}");
            }
        }
    }

    fn fire_click(&mut self, control_name: &str) {
        let callback = {
            let g = self.gui.lock().unwrap();
            if Self::gui_trace_enabled() {
                eprintln!(
                    "[gui] fire_click control={} keys={:?}",
                    control_name,
                    g.event_keys()
                );
            }
            g.get_event_handler(control_name, "Click").cloned()
        };
        if Self::gui_trace_enabled() {
            eprintln!("[gui] fire_click found={}", callback.is_some());
        }
        if let Some(cb) = callback {
            self.invoke_callback(&cb, control_name);
        }
    }

    fn invoke_callback(&mut self, cb: &vybe_bytecode::Value, control_name: &str) {
        let mut vm = self.vm.borrow_mut();
        let me = vm
            .globals
            .get("__f")
            .cloned()
            .unwrap_or(vybe_bytecode::Value::Null);
        let arity = fn_arity(cb);
        let sender = vybe_bytecode::Value::String(Arc::from(control_name));
        if Self::gui_trace_enabled() {
            eprintln!(
                "[gui] invoke_callback control={} sender={} arity={} me_type={}",
                control_name,
                control_name,
                arity,
                me.type_tag()
            );
        }
        let result = match arity {
            0 => vm.invoke(cb, &[]),
            1 => vm.invoke(cb, &[me]),
            2 => vm.invoke(cb, &[me, sender]),
            _ => vm.invoke(cb, &[me, sender, vybe_bytecode::Value::Null]),
        };
        if let Err(e) = result {
            eprintln!("Event handler error: {e}");
        } else if Self::gui_trace_enabled() {
            eprintln!("[gui] invoke_callback ok");
            if let Some(vybe_bytecode::Value::Object(form_obj)) = vm.globals.get("__f") {
                let form = form_obj.lock().unwrap();
                let keys: Vec<String> = form.properties.keys().cloned().collect();
                let txtcalc_text = form.properties.get("txtcalc").and_then(|value| {
                    if let vybe_bytecode::Value::Object(control_obj) = value {
                        let control = control_obj.lock().unwrap();
                        control
                            .properties
                            .get("text")
                            .map(|text| format!("{}", text))
                    } else {
                        None
                    }
                });
                let txtdisplay_text = form.properties.get("txtdisplay").and_then(|value| {
                    if let vybe_bytecode::Value::Object(control_obj) = value {
                        let control = control_obj.lock().unwrap();
                        control
                            .properties
                            .get("text")
                            .map(|text| format!("{}", text))
                    } else {
                        None
                    }
                });
                eprintln!(
                    "[gui] post_callback form_keys={:?} txtcalc.text={:?} txtdisplay.text={:?}",
                    keys, txtcalc_text, txtdisplay_text,
                );
            }
        }
        drop(vm);
    }

    fn sync_widgets_from_vm(&mut self) {
        let updates = {
            let vm = self.vm.borrow();
            let mut ups: Vec<(String, String)> = Vec::new();
            if let Some(vybe_bytecode::Value::Object(form_obj)) = vm.globals.get("__f") {
                let fo = form_obj.lock().unwrap();
                for (field_name, value) in &fo.properties {
                    if let vybe_bytecode::Value::Object(control_obj) = value {
                        let control = control_obj.lock().unwrap();
                        let control_name = control
                            .properties
                            .get("__control_name")
                            .or_else(|| control.properties.get("name"))
                            .map(|v| format!("{}", v).to_lowercase())
                            .filter(|name| !name.is_empty())
                            .unwrap_or_else(|| field_name.to_lowercase());
                        if let Some(text) = control.properties.get("text") {
                            ups.push((control_name, format!("{}", text)));
                        }
                    }
                }
            }
            ups
        };
        if Self::gui_trace_enabled() && !updates.is_empty() {
            eprintln!("[gui] sync_widgets_from_vm updates={:?}", updates);
        }
        if !updates.is_empty() {
            let mut g = self.gui.lock().unwrap();
            for (name, text) in updates {
                g.form.send_command(&name, &WidgetCommand::SetText(text));
            }
        }
    }

    fn process_widget_events(&mut self) {
        let events = self.gui.lock().unwrap().form.drain_events();
        if Self::gui_trace_enabled() && !events.is_empty() {
            eprintln!("[gui] process_widget_events events={:?}", events);
        }
        for event in events {
            match &event {
                WidgetEvent::ButtonClicked(name)
                | WidgetEvent::CheckboxToggled(name, _)
                | WidgetEvent::RadioSelected(name, _)
                | WidgetEvent::TextChanged(name, _)
                | WidgetEvent::LinkClicked(name) => {
                    self.fire_click(name);
                }
                WidgetEvent::SelectChanged(name, _) | WidgetEvent::ListBoxSelected(name, _) => {
                    let callback = {
                        let g = self.gui.lock().unwrap();
                        g.get_event_handler(&name, "SelectedIndexChanged")
                            .cloned()
                            .or_else(|| g.get_event_handler(&name, "Click").cloned())
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
                            if let Some(nav_info) = self
                                .navigators
                                .iter()
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
        if self.binding_sources.is_empty() {
            return;
        }

        let bs_infos: Vec<_> = self.binding_sources.clone();
        for bs_info in &bs_infos {
            let conn_str = self.get_connection_string(&bs_info.name, &bs_info.data_adapter_name);
            if conn_str.is_empty() {
                continue;
            }

            let sql = format!("SELECT * FROM {}", bs_info.data_member);
            match vybe_host::wasi::sql::query_rows(&conn_str, &sql) {
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
                    self.data_store.insert(
                        bs_info.name.to_lowercase(),
                        DataStore {
                            columns: Vec::new(),
                            rows: Vec::new(),
                            position: -1,
                        },
                    );
                }
            }
        }
        self.update_navigator_positions();
    }

    fn get_connection_string(&self, bs_name: &str, adapter_name: &str) -> String {
        let vm = self.vm.borrow();
        if let Some(vybe_bytecode::Value::Object(form_obj)) = vm.globals.get("__f") {
            let fo = form_obj.lock().unwrap();
            if let Some(vybe_bytecode::Value::Object(bs_obj)) =
                fo.properties.get(&bs_name.to_lowercase())
            {
                let bs = bs_obj.lock().unwrap();
                if let Some(vybe_bytecode::Value::Object(da_obj)) = bs.properties.get("datasource")
                {
                    let da = da_obj.lock().unwrap();
                    if let Some(v) = da.properties.get("connectionstring") {
                        return format!("{}", v);
                    }
                }
            }
            if let Some(vybe_bytecode::Value::Object(da_obj)) =
                fo.properties.get(&adapter_name.to_lowercase())
            {
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
        if store.position < 0 || store.position as usize >= store.rows.len() {
            return;
        }
        let row = &store.rows[store.position as usize];

        let vm = self.vm.borrow_mut();
        if let Some(vybe_bytecode::Value::Object(form_obj)) = vm.globals.get("__f") {
            let fo = form_obj.lock().unwrap();
            for binding in &self.data_bindings {
                if !binding.source_name.eq_ignore_ascii_case(bs_name) {
                    continue;
                }
                let col_key = row
                    .keys()
                    .find(|k| k.eq_ignore_ascii_case(&binding.column))
                    .cloned();
                let value = col_key
                    .and_then(|k| row.get(&k))
                    .cloned()
                    .unwrap_or_default();
                let ctrl_lower = binding.control_name.to_lowercase();
                if let Some(vybe_bytecode::Value::Object(ctrl_obj)) = fo.properties.get(&ctrl_lower)
                {
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
            if let Some(store) = self
                .data_store
                .get(&nav_info.binding_source_name.to_lowercase())
            {
                let pos_count = format!("{},{}", store.position, store.rows.len());
                g.form.send_command(
                    &nav_info.navigator_name.to_lowercase(),
                    &WidgetCommand::Custom(
                        "set_position_and_count".into(),
                        CommandValue::Text(pos_count),
                    ),
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
            if count == 0 {
                return;
            }
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
        vybe_bytecode::Value::Object(obj) => match &obj.lock().unwrap().kind {
            vybe_bytecode::value::ObjectKind::Function(f) => f.arity as usize,
            _ => 0,
        },
        _ => 0,
    }
}

// ── Dialog registration ────────────────────────────────────────────────

fn register_dialog_fns(vm: &mut vybe_bytecode::VM) {
    use std::sync::{Arc, Mutex};
    use vybe_bytecode::Value;
    use vybe_bytecode::value::{Object, ObjectKind};
    use vybe_widgets::dialogs::{FileDialog, FolderDialog};

    vm.register_host_fn(
        "vybe:gui",
        "__dlg_show",
        Box::new(|_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
            let (dialog_type, title) = if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                let dt = o
                    .properties
                    .get("__control_type")
                    .map(|v| format!("{}", v))
                    .unwrap_or_default();
                let t = o
                    .properties
                    .get("title")
                    .or_else(|| o.properties.get("text"))
                    .map(|v| format!("{}", v))
                    .unwrap_or_default();
                (dt, t)
            } else {
                (String::new(), String::new())
            };

            match dialog_type.as_str() {
                "OpenFileDialog" => {
                    let dlg_title = if title.is_empty() {
                        "Open File".into()
                    } else {
                        title
                    };
                    if let Some(path) = FileDialog::new(dlg_title).open() {
                        if let Some(Value::Object(obj)) = args.first() {
                            obj.lock().unwrap().properties.insert(
                                "filename".into(),
                                Value::String(Arc::from(path.to_string_lossy().as_ref())),
                            );
                        }
                        Value::I32(1)
                    } else {
                        Value::I32(0)
                    }
                }
                "SaveFileDialog" => {
                    let dlg_title = if title.is_empty() {
                        "Save File".into()
                    } else {
                        title
                    };
                    if let Some(path) = FileDialog::new(dlg_title).save() {
                        if let Some(Value::Object(obj)) = args.first() {
                            obj.lock().unwrap().properties.insert(
                                "filename".into(),
                                Value::String(Arc::from(path.to_string_lossy().as_ref())),
                            );
                        }
                        Value::I32(1)
                    } else {
                        Value::I32(0)
                    }
                }
                "FolderBrowserDialog" => {
                    let dlg_title = if title.is_empty() {
                        "Select Folder".into()
                    } else {
                        title
                    };
                    if let Some(path) = FolderDialog::new(dlg_title).pick() {
                        if let Some(Value::Object(obj)) = args.first() {
                            obj.lock().unwrap().properties.insert(
                                "selectedpath".into(),
                                Value::String(Arc::from(path.to_string_lossy().as_ref())),
                            );
                        }
                        Value::I32(1)
                    } else {
                        Value::I32(0)
                    }
                }
                "ColorDialog" | "FontDialog" => Value::I32(1),
                _ => Value::I32(0),
            }
        }),
    );

    vm.register_host_fn(
        "vybe:gui",
        "inputBox",
        Box::new(|_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
            let default = args.get(2).map(|v| format!("{}", v)).unwrap_or_default();
            Value::String(Arc::from(default.as_str()))
        }),
    );

    let dlg_show_idx = *vm
        .host_registry
        .get(&("vybe:gui".into(), "__dlg_show".into()))
        .unwrap();
    let dlg_show_ref = {
        let mut o = Object::new();
        o.kind = ObjectKind::HostFunction(dlg_show_idx);
        Value::Object(Arc::new(Mutex::new(o)))
    };
    vm.globals.insert("__dlg_show_ref".into(), dlg_show_ref);
}

// ── Extract binding info from form definition (designer forms only) ────

#[cfg(feature = "gui_forms")]
fn extract_binding_info(
    form: &vybe_compiler::projects::vbforms::Form,
) -> (
    Vec<DataBindingEntry>,
    Vec<BindingSourceInfo>,
    Vec<NavigatorInfo>,
) {
    let mut data_bindings = Vec::new();
    let mut binding_sources = Vec::new();
    let mut navigators = Vec::new();

    for ctrl in &form.controls {
        let type_name = format!("{:?}", ctrl.control_type);

        if type_name.contains("BindingSource") {
            let data_source = ctrl
                .properties
                .get_string("DataSource")
                .unwrap_or_default()
                .to_string();
            let data_member = ctrl
                .properties
                .get_string("DataMember")
                .unwrap_or_default()
                .to_string();
            if !data_source.is_empty() && !data_member.is_empty() {
                binding_sources.push(BindingSourceInfo {
                    name: ctrl.name.clone(),
                    data_adapter_name: data_source,
                    data_member,
                });
            }
        }

        if type_name.contains("BindingNavigator") {
            let bs = ctrl
                .properties
                .get_string("BindingSource")
                .unwrap_or_default()
                .to_string();
            if !bs.is_empty() {
                navigators.push(NavigatorInfo {
                    navigator_name: ctrl.name.clone(),
                    binding_source_name: bs,
                });
            }
        }

        let binding_source = ctrl
            .properties
            .get_string("DataBindings.Source")
            .map(|s| s.to_string());
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

// ── Public launch functions ────────────────────────────────────────────

/// Launch a designer form — builds widgets from a `vybe_compiler::projects::vbforms::Form` model.
/// Requires the `gui_forms` feature.
#[cfg(feature = "gui_forms")]
pub fn launch_vybewidget_form(
    mut vm: vybe_bytecode::VM,
    gui: Arc<Mutex<GuiState>>,
    form: &vybe_compiler::projects::vbforms::Form,
) {
    register_dialog_fns(&mut vm);

    if let Some(form_obj) = gui.lock().unwrap().form_object.clone() {
        vm.globals.insert("__f".into(), form_obj);
    }

    let model_control_count = form
        .controls
        .iter()
        .filter(|c| !c.control_type.is_non_visual())
        .count();
    if model_control_count > 0 {
        let id_to_bounds: std::collections::HashMap<_, _> =
            form.controls.iter().map(|c| (c.id, &c.bounds)).collect();

        let mut g = gui.lock().unwrap();
        g.form = WidgetForm::new(&form.text);
        g.control_names.clear();
        for ctrl in &form.controls {
            if ctrl.control_type.is_non_visual() {
                continue;
            }
            let widget = make_widget(ctrl);

            let mut abs_x = ctrl.bounds.x;
            let mut abs_y = ctrl.bounds.y;
            let mut parent = ctrl.parent_id;
            while let Some(pid) = parent {
                if let Some(pb) = id_to_bounds.get(&pid) {
                    abs_x += pb.x;
                    abs_y += pb.y;
                }
                parent = form
                    .controls
                    .iter()
                    .find(|c| c.id == pid)
                    .and_then(|c| c.parent_id);
            }

            g.form.add_boxed_control(
                widget,
                abs_x as f32,
                abs_y as f32,
                ctrl.bounds.width as f32,
                ctrl.bounds.height as f32,
            );
            g.control_names.push(ctrl.name.clone());
        }
    }

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

    run_app(&form.text, form.width as u32, form.height as u32, 1.0, app);
}

/// Launch a programmatic form — GuiState already has all widgets and event handlers.
pub fn launch_gui(mut vm: vybe_bytecode::VM, gui: Arc<Mutex<GuiState>>) {
    register_dialog_fns(&mut vm);

    let (title, width, height) = {
        let g = gui.lock().unwrap();
        ("Form1".to_string(), g.width, g.height)
    };

    if let Some(form_obj) = gui.lock().unwrap().form_object.clone() {
        vm.globals.insert("__f".into(), form_obj);
    }

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

/// Wrapper that dispatches to `launch_gui` or `launch_vybewidget_form`.
/// Requires the `gui_forms` feature for the designer form path.
#[cfg(feature = "gui_forms")]
pub fn launch_vm_form(
    vm: vybe_bytecode::VM,
    gui: Arc<Mutex<GuiState>>,
    initial_form: Option<vybe_compiler::projects::vbforms::Form>,
) {
    let should_launch = gui.lock().unwrap().should_run || initial_form.is_some();

    if should_launch {
        if let Some(form) = initial_form {
            launch_vybewidget_form(vm, gui, &form);
        } else {
            launch_gui(vm, gui);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::{Arc, Mutex};
    use vybe_bytecode::value::ObjectKind;
    use vybe_bytecode::{HostContext, VM, Value};
    use vybe_compiler::compiler::Compiler;
    use vybe_language_vb as vb;
    use vybe_compiler::profile::parse_profile;
    use vybe_compiler::projects;
    use vybe_host::gui_state::GuiState;
    use vybe_widgets::layout::{MouseButton, MouseEvent, MouseEventKind};

    fn run_vb_gui(src: &str) -> (VM, Arc<Mutex<GuiState>>) {
        let module = vb::parse(src).expect("VB parse failed");
        let profile = parse_profile(vb::profile_source()).expect("Failed to parse VB profile");
        let chunks = Compiler::with_profile(profile)
            .compile(&module)
            .expect("VB compile failed");

        let mut vm = VM::new();
        let gui = vybe_host::register_all_with_gui(&mut vm);
        vm.register_host_fn(
            "wasi:logging/logging",
            "log",
            Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::Null),
        );
        vybe_host::setup_namespaces(&mut vm);
        vm.run(chunks).expect("VB run failed");
        (vm, gui)
    }

    fn run_bundle_gui(path: &str) -> (VM, Arc<Mutex<GuiState>>) {
        let bundle = projects::load(std::path::Path::new(path)).expect("project load failed");
        let chunks = bundle.compile().expect("project compile failed");

        let mut vm = VM::new();
        let gui = vybe_host::register_all_with_gui(&mut vm);
        vm.register_host_fn(
            "wasi:logging/logging",
            "log",
            Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::Null),
        );
        vybe_host::setup_namespaces(&mut vm);
        vm.run(chunks).expect("project run failed");
        (vm, gui)
    }

    fn control_widget_name(form: &Value, field_name: &str) -> String {
        match form {
            Value::Object(form_obj) => {
                let form_guard = form_obj.lock().unwrap();
                match form_guard.properties.get(field_name) {
                    Some(Value::Object(control_obj)) => {
                        let control_guard = control_obj.lock().unwrap();
                        control_guard
                            .properties
                            .get("__control_name")
                            .or_else(|| control_guard.properties.get("name"))
                            .map(|value| format!("{}", value).to_lowercase())
                            .unwrap_or_else(|| field_name.to_string())
                    }
                    _ => field_name.to_string(),
                }
            }
            _ => field_name.to_string(),
        }
    }

    fn collection_count(value: &Value) -> usize {
        match value {
            Value::Object(obj) => {
                let obj = obj.lock().unwrap();
                match &obj.kind {
                    ObjectKind::Array(items) => items.len(),
                    _ => 0,
                }
            }
            _ => 0,
        }
    }

    fn collection_contains(collection: &Value, needle: &Value) -> bool {
        match collection {
            Value::Object(obj) => {
                let obj = obj.lock().unwrap();
                match &obj.kind {
                    ObjectKind::Array(items) => items.iter().any(|item| item.eq(needle)),
                    _ => false,
                }
            }
            _ => false,
        }
    }

    #[test]
    fn simulated_button_click_updates_display() {
        let source = std::fs::read_to_string(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../examples/vb/calculator.vb"
        ))
        .expect("calculator source");

        let (mut vm, gui) = run_vb_gui(&source);
        if let Some(form_obj) = gui.lock().unwrap().form_object.clone() {
            vm.globals.insert("__f".into(), form_obj);
        }

        let mut app = FormApp {
            font_system: FontSystem::new(),
            swash_cache: SwashCache::new(),
            vm: Rc::new(RefCell::new(vm)),
            gui: gui.clone(),
            data_bindings: Vec::new(),
            binding_sources: Vec::new(),
            navigators: Vec::new(),
            data_store: std::collections::HashMap::new(),
            initialised: false,
        };

        app.on_init(300.0, 400.0, 1.0);

        let press = MouseEvent {
            x: 20.0,
            y: 70.0,
            kind: MouseEventKind::Press(MouseButton::Left),
            cmd: false,
            shift: false,
            alt: false,
        };
        let release = MouseEvent {
            x: 20.0,
            y: 70.0,
            kind: MouseEventKind::Release(MouseButton::Left),
            cmd: false,
            shift: false,
            alt: false,
        };

        assert!(app.handle_mouse(press));
        assert!(app.handle_mouse(release));

        let form = app
            .vm
            .borrow()
            .globals
            .get("__f")
            .cloned()
            .expect("__f global");
        let display_name = control_widget_name(&form, "txtdisplay");
        let display_text = {
            let mut guard = gui.lock().unwrap();
            match guard
                .form
                .send_command(&display_name, &WidgetCommand::GetText)
            {
                CommandValue::Text(text) => text,
                other => panic!("Expected txtdisplay widget text, got {:?}", other),
            }
        };

        assert_eq!(display_text, "7");
    }

    #[test]
    fn simulated_project_button_click_updates_widget_text() {
        let (mut vm, gui) = run_bundle_gui(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../examples/vb/calc/Calculator.vbproj"
        ));
        if let Some(form_obj) = gui.lock().unwrap().form_object.clone() {
            vm.globals.insert("__f".into(), form_obj);
        }

        let mut app = FormApp {
            font_system: FontSystem::new(),
            swash_cache: SwashCache::new(),
            vm: Rc::new(RefCell::new(vm)),
            gui: gui.clone(),
            data_bindings: Vec::new(),
            binding_sources: Vec::new(),
            navigators: Vec::new(),
            data_store: std::collections::HashMap::new(),
            initialised: false,
        };

        app.on_init(340.0, 280.0, 1.0);

        app.fire_click("btn8");
        app.fire_click("btn5");

        let form = app
            .vm
            .borrow()
            .globals
            .get("__f")
            .cloned()
            .expect("__f global");
        let txtcalc_name = control_widget_name(&form, "txtcalc");
        let text = {
            let mut guard = gui.lock().unwrap();
            match guard
                .form
                .send_command(&txtcalc_name, &WidgetCommand::GetText)
            {
                CommandValue::Text(text) => text,
                other => panic!("Expected txtcalc widget text, got {:?}", other),
            }
        };

        assert_eq!(text, "85");
    }

    #[test]
    fn simulated_project_textbox_starts_empty_and_first_click_has_no_null_prefix() {
        let (mut vm, gui) = run_bundle_gui(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../examples/vb/calc/Calculator.vbproj"
        ));
        if let Some(form_obj) = gui.lock().unwrap().form_object.clone() {
            vm.globals.insert("__f".into(), form_obj);
        }

        let mut app = FormApp {
            font_system: FontSystem::new(),
            swash_cache: SwashCache::new(),
            vm: Rc::new(RefCell::new(vm)),
            gui: gui.clone(),
            data_bindings: Vec::new(),
            binding_sources: Vec::new(),
            navigators: Vec::new(),
            data_store: std::collections::HashMap::new(),
            initialised: false,
        };

        app.on_init(340.0, 280.0, 1.0);

        let form = app
            .vm
            .borrow()
            .globals
            .get("__f")
            .cloned()
            .expect("__f global");
        let txtcalc_name = control_widget_name(&form, "txtcalc");

        let initial_text = {
            let mut guard = gui.lock().unwrap();
            match guard
                .form
                .send_command(&txtcalc_name, &WidgetCommand::GetText)
            {
                CommandValue::Text(text) => text,
                other => panic!("Expected txtcalc widget text, got {:?}", other),
            }
        };

        assert_eq!(initial_text, "");

        app.fire_click("btn8");

        let updated_text = {
            let mut guard = gui.lock().unwrap();
            match guard
                .form
                .send_command(&txtcalc_name, &WidgetCommand::GetText)
            {
                CommandValue::Text(text) => text,
                other => panic!("Expected txtcalc widget text, got {:?}", other),
            }
        };

        assert_eq!(updated_text, "8");
    }

    #[test]
    fn project_form_exposes_controls_components_and_openforms_collections() {
        let source = r#"
Imports System.Windows.Forms

Public Class Form1
    Inherits Form

    Friend WithEvents txt1 As TextBox
    Friend WithEvents bs1 As BindingSource

    Public Sub New()
        MyBase.New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.txt1 = New TextBox()
        Me.bs1 = New BindingSource()
        Me.txt1.Name = "txt1"
        Me.bs1.Name = "bs1"
        Me.Controls.Add(Me.txt1)
    End Sub
End Class

Module Program
    Sub Main()
        Dim f As New Form1()
        Application.Run(f)
    End Sub
End Module
"#;

        let (mut vm, gui) = run_vb_gui(source);
        if let Some(form_obj) = gui.lock().unwrap().form_object.clone() {
            vm.globals.insert("__f".into(), form_obj);
        }

        let form = vm.globals.get("__f").cloned().expect("__f global");
        let open_forms = vm
            .globals
            .get("__openforms")
            .cloned()
            .expect("__openforms global");

        let (controls, components, txt1, bs1) = match &form {
            Value::Object(form_obj) => {
                let form = form_obj.lock().unwrap();
                (
                    form.properties
                        .get("controls")
                        .cloned()
                        .expect("controls collection"),
                    form.properties
                        .get("components")
                        .cloned()
                        .expect("components collection"),
                    form.properties.get("txt1").cloned().expect("txt1 field"),
                    form.properties.get("bs1").cloned().expect("bs1 field"),
                )
            }
            _ => panic!("expected form object"),
        };

        assert!(collection_count(&controls) > 0);
        assert!(collection_count(&components) > 0);
        assert!(collection_contains(&controls, &txt1));
        assert!(!collection_contains(&controls, &bs1));
        assert!(collection_contains(&components, &bs1));
        assert!(collection_contains(&open_forms, &form));
    }
}
