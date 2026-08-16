//! GUI launch layer — owns the winit event loop, VM glue, and dialog registration.
//!
//! Uses `vybe_widgets::Form` as the container for all controls, and
//! `vybe_widgets::Application` + `run_app()` for the window/event loop.
//! All graphics, focus management, hover states, and keyboard routing live
//! in vybe_widgets.
//!
//! Two entry points:
//! - `launch_gui` — programmatic forms (GuiState already has widgets)
//! - `launch_vybewidget_form` — designer forms (builds widgets from `vybe_platform_dotnet::winforms::designer::Form`)
//!   (requires `gui_forms` feature)

use std::cell::RefCell;
use std::rc::Rc;
use std::sync::{Arc, Mutex};
use vybe_platform_vybe::gui_state::GuiState;

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
    SwashCache,
    WidgetCommand,
    WidgetEvent,
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
fn make_widget(ctrl: &vybe_platform_dotnet::winforms::designer::Control) -> Box<dyn PanelWidget> {
    let text = ctrl
        .properties
        .get_string("Text")
        .unwrap_or_default()
        .to_string();
    // Preserve original case — VB compiler will lowercase identifiers as needed
    let name = &ctrl.name;

    match ctrl.control_type {
        vybe_platform_dotnet::winforms::designer::ControlType::Button => {
            let mut w = Button::new(&text).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_platform_dotnet::winforms::designer::ControlType::Label
        | vybe_platform_dotnet::winforms::designer::ControlType::LinkLabel => {
            let mut w = Label::new(&text).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_platform_dotnet::winforms::designer::ControlType::TextBox
        | vybe_platform_dotnet::winforms::designer::ControlType::RichTextBox => {
            let mut w = TextInput::new().with_name(&name);
            w.value = text;
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_platform_dotnet::winforms::designer::ControlType::MaskedTextBox => {
            let mut w = MaskedTextBox::new().with_name(&name);
            w.value = text;
            Box::new(w)
        }
        vybe_platform_dotnet::winforms::designer::ControlType::CheckBox => {
            Box::new(Checkbox::new(&text).with_name(&name))
        }
        vybe_platform_dotnet::winforms::designer::ControlType::RadioButton => {
            Box::new(Radio::new(&text).with_name(&name))
        }
        vybe_platform_dotnet::winforms::designer::ControlType::ComboBox => {
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
        vybe_platform_dotnet::winforms::designer::ControlType::ListBox
        | vybe_platform_dotnet::winforms::designer::ControlType::CheckedListBox => {
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
        vybe_platform_dotnet::winforms::designer::ControlType::Panel
        | vybe_platform_dotnet::winforms::designer::ControlType::UserControl => {
            let mut w = Panel::new().with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_platform_dotnet::winforms::designer::ControlType::Frame => {
            let mut w = GroupBox::new(&text).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_platform_dotnet::winforms::designer::ControlType::PictureBox => {
            let mut w = PictureBox::new().with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_platform_dotnet::winforms::designer::ControlType::ProgressBar => {
            let mut w = ProgressBar::new().with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_platform_dotnet::winforms::designer::ControlType::TrackBar => {
            Box::new(Slider::new(0.0, 100.0, 50.0).with_name(&name))
        }
        vybe_platform_dotnet::winforms::designer::ControlType::NumericUpDown => {
            Box::new(NumericUpDown::new().with_name(&name))
        }
        vybe_platform_dotnet::winforms::designer::ControlType::DateTimePicker => {
            Box::new(DateTimePicker::new().with_name(&name))
        }
        vybe_platform_dotnet::winforms::designer::ControlType::TreeView => {
            Box::new(TreeView::new("", 1.0).with_name(&name))
        }
        vybe_platform_dotnet::winforms::designer::ControlType::DataGridView
        | vybe_platform_dotnet::winforms::designer::ControlType::DataGrid => {
            Box::new(DataGrid::new(&[]).with_name(&name))
        }
        vybe_platform_dotnet::winforms::designer::ControlType::ListView => {
            Box::new(ListView::new().with_name(&name))
        }
        vybe_platform_dotnet::winforms::designer::ControlType::TabControl => {
            let mut w = Tabs::new(&["Tab1"]).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_platform_dotnet::winforms::designer::ControlType::MonthCalendar => {
            Box::new(MonthCalendar::new().with_name(&name))
        }
        vybe_platform_dotnet::winforms::designer::ControlType::HScrollBar => {
            let mut w = ScrollBar::new(false).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_platform_dotnet::winforms::designer::ControlType::VScrollBar => {
            let mut w = ScrollBar::new(true).with_name(&name);
            w.width = ctrl.bounds.width as f32;
            w.height = ctrl.bounds.height as f32;
            Box::new(w)
        }
        vybe_platform_dotnet::winforms::designer::ControlType::MenuStrip => {
            Box::new(MenuStrip::new().with_name(&name))
        }
        vybe_platform_dotnet::winforms::designer::ControlType::ToolStrip => {
            Box::new(ToolStrip::new().with_name(&name))
        }
        vybe_platform_dotnet::winforms::designer::ControlType::StatusStrip => {
            Box::new(StatusStrip::new().with_name(&name))
        }
        vybe_platform_dotnet::winforms::designer::ControlType::ContextMenuStrip => {
            Box::new(ContextMenu::new().with_name(&name))
        }
        vybe_platform_dotnet::winforms::designer::ControlType::SplitContainer => {
            Box::new(SplitContainer::new(false).with_name(&name))
        }
        vybe_platform_dotnet::winforms::designer::ControlType::FlowLayoutPanel => {
            Box::new(FlowLayoutPanel::new().with_name(&name))
        }
        vybe_platform_dotnet::winforms::designer::ControlType::TableLayoutPanel => {
            Box::new(TableLayoutPanel::new(2, 2).with_name(&name))
        }
        vybe_platform_dotnet::winforms::designer::ControlType::BindingNavigator => {
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
    vm: Rc<RefCell<vybe_runtime::VM>>,
    gui: Arc<Mutex<GuiState>>,
    // Data binding state
    data_bindings: Vec<DataBindingEntry>,
    binding_sources: Vec<BindingSourceInfo>,
    navigators: Vec<NavigatorInfo>,
    data_store: std::collections::HashMap<String, DataStore>,
    initialised: bool,
    /// Last time each GUI timer fired, keyed by control name — used by `on_tick`
    /// to decide when a timer's interval has elapsed.
    timer_last_fire: std::collections::HashMap<String, std::time::Instant>,
}

impl Application for FormApp {
    fn on_init(&mut self, width: f32, height: f32, _scale: f32) {
        Self::lay_out(&self.gui, width, height);
        if !self.initialised {
            self.initialised = true;
            self.fire_load_event();
            self.init_data_bindings();
        }
    }

    fn on_resize(&mut self, width: f32, height: f32) {
        Self::lay_out(&self.gui, width, height);
    }

    /// **The window's title is the document's title.**
    ///
    /// `document.title = "Calculator"` reached `Document::set_title` and stopped
    /// there: `FormApp` never implemented this, so it returned the trait's empty
    /// default, `app_window` skips an empty title, and the window kept the name
    /// the launcher gave it at startup — `Form1`, whatever the page said.
    ///
    /// Read live rather than cached at launch, which is what makes a title set
    /// LATER — or changed while running — reach the chrome. `app_window` polls
    /// this after each render and only touches the window when the string
    /// actually changes, so re-reading per frame costs a comparison.
    fn title(&self) -> String {
        crate::gui_document::with_live(|doc| doc.title()).unwrap_or_default()
    }

    fn render(&mut self, pixmap: &mut Pixmap, scale: f32) {
        // One renderer, shared with `--capture` and the debugger's `capture`,
        // so a captured frame is byte-for-byte what the window shows.
        crate::gui_capture::render_into(
            &self.gui,
            pixmap,
            &mut self.font_system,
            &mut self.swash_cache,
            scale,
        );
    }

    fn handle_mouse(&mut self, event: MouseEvent) -> bool {
        if Self::gui_trace_enabled() {
            eprintln!("[gui] formapp.handle_mouse event={:?}", event);
        }
        // Window events become W3C UI Events in the `web:ui-events` queue —
        // the same queue a browser host fills from the real DOM. SDL's
        // vocabulary is applied later, by SDL's own adapter, not here.
        {
            use vybe_widgets::layout::{MouseButton, MouseEventKind};
            use vybe_widgets::ui_events::{UiEvent, queue};
            // DOM `button`: 0 left, 1 middle, 2 right.
            let dom_button = |b: &MouseButton| match b {
                MouseButton::Left => 0i32,
                MouseButton::Middle => 1,
                MouseButton::Right => 2,
            };
            // DOM `buttons` mask: 1 left, 2 right, 4 middle.
            let dom_mask = |b: &MouseButton| match b {
                MouseButton::Left => 1i32,
                MouseButton::Right => 2,
                MouseButton::Middle => 4,
            };
            let (kind, button, buttons) = match &event.kind {
                MouseEventKind::Press(b) => ("mousedown", dom_button(b), dom_mask(b)),
                MouseEventKind::Release(b) => ("mouseup", dom_button(b), 0),
                MouseEventKind::Move | MouseEventKind::Scroll(_) => ("mousemove", 0, 0),
            };
            queue().push(UiEvent {
                kind: kind.to_string(),
                client_x: event.x as i32,
                client_y: event.y as i32,
                button,
                buttons,
                ..UiEvent::default()
            });
        }
        if crate::gui_document::with_live(|d| d.form_mut().handle_mouse(&event)).is_none() {
            self.gui.lock().unwrap().form.handle_mouse(&event);
        }
        self.process_widget_events();
        self.dispatch_document_events();
        true
    }

    fn handle_key(&mut self, event: KeyEvent) -> bool {
        // Both edges as `keydown`/`keyup`, in W3C shape.
        {
            use vybe_widgets::ui_events::{UiEvent, queue};
            let pressed = event.state == vybe_widgets::winit::event::ElementState::Pressed;
            let (key, code, key_code) = dom_key_fields(&event.key_without_modifiers);
            queue().push(UiEvent {
                kind: if pressed { "keydown" } else { "keyup" }.to_string(),
                key,
                code,
                key_code,
                ctrl_key: event.cmd,
                shift_key: event.shift,
                alt_key: event.alt,
                ..UiEvent::default()
            });
            // Widget dispatch is pressed-only — releases exist for the event
            // queue alone, so typing/focus behavior is unchanged.
            if !pressed {
                return false;
            }
        }
        if crate::gui_document::with_live(|d| d.form_mut().handle_key(&event)).is_none() {
            self.gui.lock().unwrap().form.handle_key(&event);
        }
        self.process_widget_events();
        self.dispatch_document_events();
        true
    }

    fn handle_scroll(&mut self, delta: f32, x: f32, y: f32) -> bool {
        {
            use vybe_widgets::ui_events::{UiEvent, queue};
            // DOM `deltaY` is positive DOWN — the opposite of a scroll delta
            // that reports "up" as positive.
            queue().push(UiEvent {
                kind: "wheel".to_string(),
                client_x: x as i32,
                client_y: y as i32,
                delta_y: -(delta as f64),
                ..UiEvent::default()
            });
        }
        let handled =
            match crate::gui_document::with_live(|d| d.form_mut().handle_scroll(delta, x, y)) {
                Some(handled) => handled,
                None => self.gui.lock().unwrap().form.handle_scroll(delta, x, y),
            };
        self.dispatch_document_events();
        handled
    }

    fn cursor_icon(&self) -> vybe_widgets::CursorIcon {
        vybe_widgets::CursorIcon::Default
    }

    /// Fire any GUI timers whose interval has elapsed. Called ~60 Hz from the
    /// event loop's `about_to_wait`. Each due timer's `OnTimer`/`Tick` handler
    /// runs through the VM (like a click); the ~60 Hz repaint then reflects any
    /// state it changed. This is what makes `TTimer`/`WinForms.Timer` actually
    /// tick — nothing drove them before.
    fn on_tick(&mut self) {
        // ── The frame boundary ────────────────────────────────────────────
        //
        // `requestAnimationFrame` callbacks run HERE, before the ~60 Hz
        // repaint, which is precisely the browser's contract: a page is
        // called back before the next paint and draws then. This is what
        // replaces "present the buffer" — the window is already redrawing,
        // so a guest that re-registers each frame gets a real animation
        // loop, and one that stops asking costs nothing.
        {
            use vybe_runtime::scheduler::DeferredSource;
            let clock = vybe_platform_web::animation::callbacks();
            let stamp = vybe_runtime::Value::F64(vybe_runtime::event_loop::monotonic_now_ms());
            // Drain only what this frame owes: `pop_due` stops handing out
            // callbacks once the frame advances, so a callback that
            // re-registers runs NEXT frame instead of spinning here.
            while let Some(cb) = clock.pop_due() {
                let mut vm = self.vm.borrow_mut();
                let _ = match fn_arity(&cb) {
                    0 => vm.invoke(&cb, &[]),
                    _ => vm.invoke(&cb, &[stamp.clone()]),
                };
            }
        }

        // Anything the document queued that input handling did not already
        // drain — a widget that reports on a later frame, or an event raised
        // by a timer/animation callback rather than by a click.
        self.dispatch_document_events();

        let now = std::time::Instant::now();
        let due: Vec<vybe_runtime::Value> = {
            let timers = self.gui.lock().unwrap().active_timers();
            let mut due = Vec::new();
            for (name, interval_ms, handler) in timers {
                let last = self.timer_last_fire.entry(name).or_insert(now);
                if now.duration_since(*last) >= std::time::Duration::from_millis(interval_ms) {
                    *last = now;
                    due.push(handler);
                }
            }
            due
        };
        if due.is_empty() {
            return;
        }
        // Receiver for instance-method handlers (`procedure Tick(Sender)`), same
        // rule as click handlers: the form, from `__f` (falling back to the live
        // GuiState form object). Args by arity.
        let me = {
            let vm = self.vm.borrow();
            vm.globals.get("__f").cloned()
        }
        .or_else(|| self.gui.lock().unwrap().form_object.clone())
        .unwrap_or(vybe_runtime::Value::Null);
        for handler in due {
            let mut vm = self.vm.borrow_mut();
            let _ = match fn_arity(&handler) {
                0 => vm.invoke(&handler, &[]),
                1 => vm.invoke(&handler, &[me.clone()]),
                _ => vm.invoke(&handler, &[me.clone(), vybe_runtime::Value::Null]),
            };
        }
    }
}

// ── VM glue ────────────────────────────────────────────────────────────

impl FormApp {
    fn gui_trace_enabled() -> bool {
        std::env::var("VYBE_GUI_TRACE")
            .map(|value| !matches!(value.as_str(), "" | "0" | "false" | "False"))
            .unwrap_or(false)
    }

    /// Give the tree the window's size. The document's viewport IS the body's
    /// containing block, so a document that never gets one lays every control
    /// out against its 800×600 default instead of the window.
    ///
    /// Both forms are sized, not one: `GuiState.form` is still the tree for a
    /// designer form, and neither is authoritative for every program.
    fn lay_out(gui: &Arc<Mutex<GuiState>>, width: f32, height: f32) {
        crate::gui_document::with_live(|d| d.set_viewport(width, height));
        gui.lock()
            .unwrap()
            .form
            .set_rect(LayoutRect::new(0.0, 0.0, width, height));
    }

    /// Turn what the user did in the document into VM calls.
    ///
    /// `web:dom` listeners are where EVERY frontend's event wiring ends up —
    /// `primitives/gui.rs` lowers VCL's `OnClick := h`, WinForms' `Click += h`
    /// and Flutter's `onPressed: h` to the same `addEventListener` — so this
    /// single drain serves all of them and none of it is framework-specific.
    ///
    /// A listener is called with the Event and nothing else, which is what a
    /// document does and the only thing it CAN do — the receiver is already
    /// bound into the handler by `primitives/gui.rs`, so there is no form
    /// object to look up and no arity table to keep in step with four
    /// frontends.
    fn dispatch_document_events(&mut self) {
        for dispatch in crate::gui_document::drain() {
            if Self::gui_trace_enabled() {
                eprintln!(
                    "[gui] document dispatch kind={} sender={}",
                    dispatch.kind, dispatch.sender
                );
            }
            let mut vm = self.vm.borrow_mut();
            if let Err(e) = vm.invoke(&dispatch.callback, &[dispatch.event]) {
                eprintln!("Event handler error: {e}");
            }
        }
    }

    fn fire_load_event(&mut self) {
        // `Handles Me.Load` is a subscription on the FORM, and a form IS the
        // document's body — so the listener lives on the body node, not in
        // `GuiState`'s name-keyed table. Reading only the table meant a
        // designer form's `Form1_Load` never ran: `TicTacToe` left `turn`
        // unset, so every cell click hit `If turn <> "X" Then Exit Sub` and the
        // window looked completely dead while every handler was correctly
        // wired. The table stays as the fallback for a form built without a
        // document, which is the same rule the rest of this file paints by.
        let document_listeners =
            crate::gui_document::listeners_for(vybe_widgets::dom::DOCUMENT, "load");
        if !document_listeners.is_empty() {
            let event = crate::gui_document::event_object("load", vybe_widgets::dom::DOCUMENT);
            for listener in document_listeners {
                let mut vm = self.vm.borrow_mut();
                if let Err(e) = vm.invoke(&listener, &[event.clone()]) {
                    eprintln!("[LOAD] Error: {e}");
                }
            }
            return;
        }
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
                .unwrap_or(vybe_runtime::Value::Null);
            let arity = fn_arity(&cb);
            let result = match arity {
                0 => vm.invoke(&cb, &[]),
                1 => vm.invoke(&cb, &[me]),
                _ => vm.invoke(
                    &cb,
                    &[me, vybe_runtime::Value::Null, vybe_runtime::Value::Null],
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

    /// Fire a value-bearing input event (checkbox toggle, radio select, text
    /// change, slider change). The control's NEW value becomes the handler's
    /// FIRST argument for an arity-1 handler — matching Flutter's
    /// `onChanged(v)`. Arity-2+ handlers keep the framework-agnostic
    /// `(sender, e)` shape (`.NET`/VB), so no language is special-cased.
    fn fire_value_event(&mut self, control_name: &str, value: vybe_runtime::Value) {
        let callback = {
            let g = self.gui.lock().unwrap();
            g.get_event_handler(control_name, "Click").cloned()
        };
        if let Some(cb) = callback {
            self.invoke_callback_with_value(&cb, control_name, value);
        }
    }

    fn invoke_callback_with_value(
        &mut self,
        cb: &vybe_runtime::Value,
        control_name: &str,
        value: vybe_runtime::Value,
    ) {
        let mut vm = self.vm.borrow_mut();
        let me = vm
            .globals
            .get("__f")
            .cloned()
            .unwrap_or(vybe_runtime::Value::Null);
        let arity = fn_arity(cb);
        let sender = vybe_runtime::Value::String(Arc::from(control_name));
        let result = match arity {
            0 => vm.invoke(cb, &[]),
            // Flutter `onChanged(value)` — the new value is the sole argument.
            1 => vm.invoke(cb, &[value]),
            // .NET/VB `(sender, e)` — unchanged.
            2 => vm.invoke(cb, &[me, sender]),
            _ => vm.invoke(cb, &[me, sender, value]),
        };
        if let Err(e) = result {
            eprintln!("Event handler error: {e}");
        }
        drop(vm);
    }

    fn invoke_callback(&mut self, cb: &vybe_runtime::Value, control_name: &str) {
        let mut vm = self.vm.borrow_mut();
        let me = vm
            .globals
            .get("__f")
            .cloned()
            .unwrap_or(vybe_runtime::Value::Null);
        let arity = fn_arity(cb);
        let sender = vybe_runtime::Value::String(Arc::from(control_name));
        if Self::gui_trace_enabled() {
            eprintln!(
                "[gui] invoke_callback control={} sender={} arity={} me_type={}",
                control_name,
                control_name,
                arity,
                me.type_tag()
            );
        }
        let _t0 = std::time::Instant::now();
        let result = match arity {
            0 => vm.invoke(cb, &[]),
            1 => vm.invoke(cb, &[me]),
            2 => vm.invoke(cb, &[me, sender]),
            _ => vm.invoke(cb, &[me, sender, vybe_runtime::Value::Null]),
        };
        if Self::gui_trace_enabled() {
            eprintln!(
                "[gui] callback elapsed={:.1}ms control={}",
                _t0.elapsed().as_secs_f64() * 1000.0,
                control_name
            );
        }
        if let Err(e) = result {
            eprintln!("Event handler error: {e}");
        } else if Self::gui_trace_enabled() {
            eprintln!("[gui] invoke_callback ok");
            if let Some(vybe_runtime::Value::Object(form_obj)) = vm.globals.get("__f") {
                let form = form_obj.lock().unwrap();
                let keys: Vec<String> = form.properties.keys().cloned().collect();
                let txtcalc_text = form.properties.get("txtcalc").and_then(|value| {
                    if let vybe_runtime::Value::Object(control_obj) = value {
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
                    if let vybe_runtime::Value::Object(control_obj) = value {
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
            if let Some(vybe_runtime::Value::Object(form_obj)) = vm.globals.get("__f") {
                let fo = form_obj.lock().unwrap();
                for (field_name, value) in &fo.properties {
                    if let vybe_runtime::Value::Object(control_obj) = value {
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
                WidgetEvent::ButtonClicked(name) | WidgetEvent::LinkClicked(name) => {
                    self.fire_click(name);
                }
                // Value-bearing input events carry the control's NEW value
                // (checkbox/radio bool, text string, slider number). Deliver
                // that value as the handler's first argument so Flutter's
                // `onChanged(v)` receives it (see `fire_value_event`).
                WidgetEvent::CheckboxToggled(name, v) | WidgetEvent::RadioSelected(name, v) => {
                    self.fire_value_event(name, vybe_runtime::Value::Bool(*v));
                }
                WidgetEvent::TextChanged(name, s) => {
                    self.fire_value_event(name, vybe_runtime::Value::String(Arc::from(s.as_str())));
                }
                WidgetEvent::SliderChanged(name, v) => {
                    self.fire_value_event(name, vybe_runtime::Value::F64(*v as f64));
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
            match vybe_platform_wasi::sql::query_rows(&conn_str, &sql) {
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
        if let Some(vybe_runtime::Value::Object(form_obj)) = vm.globals.get("__f") {
            let fo = form_obj.lock().unwrap();
            if let Some(vybe_runtime::Value::Object(bs_obj)) =
                fo.properties.get(&bs_name.to_lowercase())
            {
                let bs = bs_obj.lock().unwrap();
                if let Some(vybe_runtime::Value::Object(da_obj)) = bs.properties.get("datasource") {
                    let da = da_obj.lock().unwrap();
                    if let Some(v) = da.properties.get("connectionstring") {
                        return format!("{}", v);
                    }
                }
            }
            if let Some(vybe_runtime::Value::Object(da_obj)) =
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
        if let Some(vybe_runtime::Value::Object(form_obj)) = vm.globals.get("__f") {
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
                // A bound control IS an element, so the binding writes to the
                // DOCUMENT — through `set_text`, the same entry point the guest
                // reaches for the `text` role, which is what makes an `<input>`
                // take its `value` and a `<select>` take its options without
                // this knowing the difference.
                //
                // Writing `ctrl_obj.properties["text"]` was the old model: a
                // plain property on the control OBJECT, which nothing paints
                // from once the control is a node. Same shape as the `Load`
                // event and `vybe.gui.setProperty` — an axis converted to the
                // document with one reader left on the old registry.
                //
                // The BOUND PROPERTY decides which write, the same two-way
                // split the guest's own property path makes: a text-ish role
                // goes through `set_text`, and anything else becomes an
                // attribute — which is where an unmapped property belongs on
                // the web, and exactly what `emit_gui_property_set` does with
                // one it has no operation for.
                let bound_to_document = crate::gui_document::node_by_id(&binding.control_name)
                    .and_then(|node| {
                        let property = binding.property.to_ascii_lowercase();
                        match property.as_str() {
                            "text" | "value" | "caption" => {
                                crate::gui_document::inspect::set_text(node, &value)
                            }
                            _ => crate::gui_document::inspect::set_attribute(
                                node, &property, &value,
                            ),
                        }
                    })
                    .is_some();
                if let Some(vybe_runtime::Value::Object(ctrl_obj)) = fo.properties.get(&ctrl_lower)
                {
                    // The object keeps the value too: a guest reading
                    // `txt.Text` before the next paint asks the object, and a
                    // form built with no document has only this.
                    ctrl_obj.lock().unwrap().properties.insert(
                        binding.property.to_lowercase(),
                        vybe_runtime::Value::String(Arc::from(value.as_str())),
                    );
                }
                let _ = bound_to_document;
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

pub(crate) fn fn_arity(val: &vybe_runtime::Value) -> usize {
    match val {
        vybe_runtime::Value::Object(obj) => match &obj.lock().unwrap().kind {
            vybe_runtime::value::ObjectKind::Function(f) => f.arity as usize,
            _ => 0,
        },
        _ => 0,
    }
}

// ── Dialog registration ────────────────────────────────────────────────

/// winit key → W3C `KeyboardEvent` fields: `(key, code, keyCode)`.
///
/// `key` is what the keypress MEANS ("a", "Enter", "ArrowLeft"), `code` is
/// the physical key ("KeyA", "Digit1", "ArrowLeft") and stays layout-
/// independent, `keyCode` is the legacy numeric identity browsers still
/// ship. No SDL here — SDL's keysyms are derived from these by its own
/// adapter, so a browser host producing real DOM events needs no changes.
fn dom_key_fields(key: &vybe_widgets::winit::keyboard::Key) -> (String, String, i32) {
    use vybe_widgets::winit::keyboard::{Key, NamedKey};
    match key {
        Key::Character(text) => {
            let Some(c) = text.chars().next() else {
                return (String::new(), String::new(), 0);
            };
            let lower = c.to_ascii_lowercase();
            let code = match lower {
                'a'..='z' => format!("Key{}", lower.to_ascii_uppercase()),
                '0'..='9' => format!("Digit{}", lower),
                '-' => "Minus".to_string(),
                '=' => "Equal".to_string(),
                ' ' => "Space".to_string(),
                _ => String::new(),
            };
            // `keyCode` is the uppercase code point for letters, the digit
            // for digits — the browser's legacy convention.
            let key_code = if lower.is_ascii_alphabetic() {
                lower.to_ascii_uppercase() as i32
            } else {
                lower as i32
            };
            (c.to_string(), code, key_code)
        }
        Key::Named(named) => {
            let f = |k: &str, c: &str, kc: i32| (k.to_string(), c.to_string(), kc);
            match named {
                NamedKey::Enter => f("Enter", "Enter", 13),
                NamedKey::Escape => f("Escape", "Escape", 27),
                NamedKey::Backspace => f("Backspace", "Backspace", 8),
                NamedKey::Tab => f("Tab", "Tab", 9),
                NamedKey::Space => f(" ", "Space", 32),
                NamedKey::Delete => f("Delete", "Delete", 46),
                NamedKey::ArrowRight => f("ArrowRight", "ArrowRight", 39),
                NamedKey::ArrowLeft => f("ArrowLeft", "ArrowLeft", 37),
                NamedKey::ArrowDown => f("ArrowDown", "ArrowDown", 40),
                NamedKey::ArrowUp => f("ArrowUp", "ArrowUp", 38),
                NamedKey::Home => f("Home", "Home", 36),
                NamedKey::End => f("End", "End", 35),
                NamedKey::PageUp => f("PageUp", "PageUp", 33),
                NamedKey::PageDown => f("PageDown", "PageDown", 34),
                NamedKey::Insert => f("Insert", "Insert", 45),
                NamedKey::CapsLock => f("CapsLock", "CapsLock", 20),
                NamedKey::F1 => f("F1", "F1", 112),
                NamedKey::F2 => f("F2", "F2", 113),
                NamedKey::F3 => f("F3", "F3", 114),
                NamedKey::F4 => f("F4", "F4", 115),
                NamedKey::F5 => f("F5", "F5", 116),
                NamedKey::F6 => f("F6", "F6", 117),
                NamedKey::F7 => f("F7", "F7", 118),
                NamedKey::F8 => f("F8", "F8", 119),
                NamedKey::F9 => f("F9", "F9", 120),
                NamedKey::F10 => f("F10", "F10", 121),
                NamedKey::F11 => f("F11", "F11", 122),
                NamedKey::F12 => f("F12", "F12", 123),
                NamedKey::Control => f("Control", "ControlLeft", 17),
                NamedKey::Shift => f("Shift", "ShiftLeft", 16),
                NamedKey::Alt => f("Alt", "AltLeft", 18),
                _ => (String::new(), String::new(), 0),
            }
        }
        _ => (String::new(), String::new(), 0),
    }
}

fn register_dialog_fns(vm: &mut vybe_runtime::VM) {
    use std::sync::Arc;
    use vybe_runtime::Value;
    use vybe_runtime::value::{Object, ObjectKind};
    use vybe_widgets::dialogs::{FileDialog, FolderDialog};

    vm.register_host_fn(
        "vybe:gui",
        "__dlg_show",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
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
        Box::new(|_ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
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
        Value::Object(vybe_runtime::heap::alloc(o))
    };
    vm.globals.insert("__dlg_show_ref".into(), dlg_show_ref);
}

// ── Extract binding info from form definition (designer forms only) ────

#[cfg(feature = "gui_forms")]
fn extract_binding_info(
    form: &vybe_platform_dotnet::winforms::designer::Form,
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

/// Launch a designer form — builds widgets from a `vybe_platform_dotnet::winforms::designer::Form` model.
/// Requires the `gui_forms` feature.
#[cfg(feature = "gui_forms")]
pub fn launch_vybewidget_form(
    mut vm: vybe_runtime::VM,
    gui: Arc<Mutex<GuiState>>,
    form: &vybe_platform_dotnet::winforms::designer::Form,
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
        timer_last_fire: std::collections::HashMap::new(),
    };

    run_app(&form.text, form.width as u32, form.height as u32, 1.0, app);
}

/// Launch a programmatic form — GuiState already has all widgets and event handlers.
pub fn launch_gui(mut vm: vybe_runtime::VM, gui: Arc<Mutex<GuiState>>) {
    register_dialog_fns(&mut vm);

    // The document's viewport is where a form's `Width`/`Height` land — they
    // are CSS on the body. `GuiState`'s pair is the fallback for a form built
    // without a document.
    let (title, width, height) = {
        let g = gui.lock().unwrap();
        let (width, height) = crate::gui_document::viewport().unwrap_or((g.width, g.height));
        ("Form1".to_string(), width, height)
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
        timer_last_fire: std::collections::HashMap::new(),
    };

    run_app(&title, width, height, 1.0, app);
}

/// Headless counterpart to `launch_gui`: lay the form out, render ONE frame
/// offscreen, write it as a PNG, and return — no window, no event loop.
///
/// Everything a GUI program drew during its run is already in `GuiState`, so
/// this replays it exactly as the window would. Used by `--capture`; it makes a
/// GUI program's output readable without a screenshot, which also means a frame
/// can be diffed as a regression check.
pub fn capture_gui(
    mut vm: vybe_runtime::VM,
    gui: Arc<Mutex<GuiState>>,
    path: &str,
    control: Option<&str>,
) -> Result<(u32, u32), String> {
    register_dialog_fns(&mut vm);

    if let Some(form_obj) = gui.lock().unwrap().form_object.clone() {
        vm.globals.insert("__f".into(), form_obj);
    }

    // Same size question as `launch_gui`, same answer — a capture must be the
    // frame the window would show.
    let (width, height) = {
        let g = gui.lock().unwrap();
        crate::gui_document::viewport().unwrap_or((g.width, g.height))
    };

    let mut app = FormApp {
        font_system: FontSystem::new(),
        swash_cache: SwashCache::new(),
        vm: Rc::new(RefCell::new(vm)),
        gui: Arc::clone(&gui),
        data_bindings: Vec::new(),
        binding_sources: Vec::new(),
        navigators: Vec::new(),
        data_store: std::collections::HashMap::new(),
        initialised: false,
        timer_last_fire: std::collections::HashMap::new(),
    };

    // Lays the form out and fires Load + data bindings, exactly as the window
    // path does — without this, controls have no rect and the frame is blank.
    app.on_init(width as f32, height as f32, 1.0);

    crate::gui_capture::capture_to_png(&gui, path, control, 1.0)
}

/// Wrapper that dispatches to `launch_gui` or `launch_vybewidget_form`.
/// Requires the `gui_forms` feature for the designer form path.
#[cfg(feature = "gui_forms")]
pub fn launch_vm_form(
    vm: vybe_runtime::VM,
    gui: Arc<Mutex<GuiState>>,
    initial_form: Option<vybe_platform_dotnet::winforms::designer::Form>,
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
    use vybe_compiler::primitives::Compiler;
    use vybe_compiler::profile::parse_profile;
    use vybe_compiler::projects;
    use vybe_language_vb as vb;
    use vybe_platform_vybe::gui_state::GuiState;
    use vybe_runtime::value::ObjectKind;
    use vybe_runtime::{HostContext, VM, Value};
    use vybe_widgets::layout::{MouseButton, MouseEvent, MouseEventKind};

    fn run_vb_gui(src: &str) -> (VM, Arc<Mutex<GuiState>>) {
        let module = vb::parse(src).expect("VB parse failed");
        let profile = parse_profile(vb::profile_source()).expect("Failed to parse VB profile");
        let chunks = Compiler::with_profile(profile)
            .compile(&module)
            .expect("VB compile failed");

        let mut vm = VM::new();
        let gui = crate::cli::register_plugins_with_gui(
            &mut vm,
            &vybe_runtime::capabilities::Capabilities::all(),
        );
        vm.register_host_fn(
            "web:console",
            "log",
            Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::Null),
        );

        vm.run(chunks).expect("VB run failed");
        (vm, gui)
    }

    fn run_bundle_gui(path: &str) -> (VM, Arc<Mutex<GuiState>>) {
        let bundle = projects::load(std::path::Path::new(path)).expect("project load failed");
        let chunks = bundle.compile().expect("project compile failed");

        let mut vm = VM::new();
        let gui = crate::cli::register_plugins_with_gui(
            &mut vm,
            &vybe_runtime::capabilities::Capabilities::all(),
        );
        vm.register_host_fn(
            "web:console",
            "log",
            Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::Null),
        );

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
            timer_last_fire: std::collections::HashMap::new(),
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
            timer_last_fire: std::collections::HashMap::new(),
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
            timer_last_fire: std::collections::HashMap::new(),
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
