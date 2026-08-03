//! GUI state — owns the widget form and event handlers.
//!
//! This is the bridge between the `vybe:gui` host module and `vybe_widgets`.
//! Host functions directly create widgets and register events here.
//! The CLI just takes the finished form and runs it in a window.

use std::collections::HashMap;
use vybe_runtime::Value;
use vybe_widgets::canvas::RecordingCanvas;
use vybe_widgets::{
    BindingNavigator, Button, Canvas as CanvasWidget, Checkbox, CommandValue, ContextMenu,
    DataGrid, DateTimePicker, FlowLayoutPanel, Form as WidgetForm, GroupBox, Label, ListBox,
    ListView, MaskedTextBox, MenuStrip, MonthCalendar, NumericUpDown, Panel, PanelWidget,
    PictureBox, ProgressBar, Radio, ScrollBar, Select, Slider, SplitContainer, StatusStrip,
    TableLayoutPanel, Tabs, TextInput, ToolStrip, TreeView, WidgetCommand };

/// Holds the live widget form + event callbacks.
/// Created before VM runs, shared with host fns via `Arc<Mutex<>>`.
pub struct GuiState {
    /// The widget form being built by host functions.
    pub form: WidgetForm,
    /// Event handlers: "controlname.eventname" → VM callback Value.
    pub event_handlers: HashMap<String, Value>,
    /// Control names in insertion order.
    pub control_names: Vec<String>,
    /// Public control names (for example .NET `Name`) mapped back to the
    /// physical widget name used by `vybe_widgets`.
    control_aliases: HashMap<String, String>,
    /// Current public/logical name for each physical widget name.
    control_public_names: HashMap<String, String>,
    /// Form dimensions (logical pixels, before scaling).
    pub width: u32,
    pub height: u32,
    /// Whether runApplication/showForm was called — signals the CLI to open a window.
    pub should_run: bool,
    /// The form object reference for `this`/`Me` dispatch.
    pub form_object: Option<Value>,
    /// Pending close request.
    pub close_requested: bool,
    /// Property store: every `set_property` write is mirrored here keyed by
    /// `(control name, lowercased property)`. Used both as the
    /// authoritative source for `get_property` (so callers see what was
    /// written even when the control isn't a child widget on the form, e.g.
    /// the form itself) and to handle form-level properties without needing
    /// to hand-craft a separate widget for the form.
    pub properties: HashMap<(String, String), String>,
    /// Set by `Control.Refresh()` / `Invalidate()` / `Update()`. The form
    /// runner clears this each frame and triggers a repaint when set. The
    /// underlying tiny-skia renderer doesn't need explicit invalidation
    /// (it repaints every frame), but tracking the flag lets us skip
    /// redundant work in headless tests and mirrors the .NET semantics
    /// for callers that check repaint state.
    pub needs_repaint: bool,
    /// Set by `Form.Activate()`. The CLI window driver checks this on
    /// each frame and brings the OS window to the foreground.
    pub front_requested: bool,
    /// Per-control overlay recordings — keyed by control
    /// name. Populated by the `vybe:gui::canvas*` host fns when user
    /// code calls `Graphics.DrawLine` etc. against a `Graphics` handle
    /// created from a non-Canvas-widget control (e.g.
    /// `Me.CreateGraphics()` on a Form, or `btn.CreateGraphics()` on a
    /// Button). The form's render loop replays each entry through the
    /// matching widget's `paint_overlay` hook each frame.
    ///
    /// Canvas WIDGETS (`vybe_widgets::Canvas`) carry their own
    /// `RecordingCanvas` inside the widget itself — those don't need an
    /// entry here. The host bridge looks at the form's child widgets
    /// FIRST, falls back to this overlay map only when no Canvas widget
    /// matches the requested control name.
    pub overlay_canvases: HashMap<String, RecordingCanvas> }

impl GuiState {
    pub fn new() -> Self {
        Self {
            form: WidgetForm::new("Form1"),
            event_handlers: HashMap::new(),
            control_names: Vec::new(),
            control_aliases: HashMap::new(),
            control_public_names: HashMap::new(),
            width: 800,
            height: 600,
            should_run: false,
            form_object: None,
            close_requested: false,
            properties: Default::default(),
            needs_repaint: false,
            front_requested: false,
            overlay_canvases: HashMap::new() }
    }

    /// VM hot-reset (bucket D): drop all script-created GUI state — controls,
    /// event handlers, the property store, the form object, canvases — back to
    /// a pristine post-boot form. The runner calls this on its shared
    /// `Arc<Mutex<GuiState>>` as part of `reset_to`, so a reused VM never leaks
    /// the previous run's window/controls. See `vmhotresetplan.md` bucket D.
    pub fn reset(&mut self) {
        *self = GuiState::new();
    }

    /// Resolve a control name to the canonical spelling stored by the live form.
    ///
    /// Exact match wins. If not found, we do a case-insensitive match so VB-style
    /// callers can still reach controls regardless of source casing.
    fn resolve_control_name(&self, control: &str) -> String {
        if let Some(target) = self.control_aliases.get(control) {
            return target.clone();
        }
        if let Some((_, target)) = self
            .control_aliases
            .iter()
            .find(|(alias, _)| alias.eq_ignore_ascii_case(control))
        {
            return target.clone();
        }
        if self.control_public_names.contains_key(control) {
            return control.to_string();
        }
        if let Some((physical, _)) = self
            .control_public_names
            .iter()
            .find(|(physical, _)| physical.eq_ignore_ascii_case(control))
        {
            return physical.clone();
        }
        if self.control_names.iter().any(|n| n == control) {
            return control.to_string();
        }
        if let Some(found) = self
            .control_names
            .iter()
            .find(|n| n.eq_ignore_ascii_case(control))
        {
            return found.clone();
        }
        control.to_string()
    }

    fn public_control_name(&self, control: &str) -> String {
        let physical = self.resolve_control_name(control);
        self.control_public_names
            .get(&physical)
            .cloned()
            .unwrap_or(physical)
    }

    pub fn is_live_control_name(&self, control: &str) -> bool {
        let physical = self.resolve_control_name(control);
        self.control_names
            .iter()
            .any(|name| name.eq_ignore_ascii_case(&physical))
    }

    pub fn rename_control(&mut self, control: &str, new_name: &str) {
        let physical = self.resolve_control_name(control);
        let new_public = new_name.trim().to_lowercase();
        if physical.is_empty() || new_public.is_empty() {
            return;
        }

        let previous_public = self.public_control_name(&physical);
        if previous_public.eq_ignore_ascii_case(&new_public) {
            self.control_public_names
                .insert(physical.clone(), new_public.clone());
            self.control_aliases.insert(new_public.clone(), physical);
            return;
        }

        self.control_public_names
            .insert(physical.clone(), new_public.clone());
        self.control_aliases.retain(|alias, target| {
            !(target.eq_ignore_ascii_case(&physical)
                && alias.eq_ignore_ascii_case(&previous_public))
        });
        self.control_aliases
            .insert(new_public.clone(), physical.clone());
        if !self
            .control_names
            .iter()
            .any(|name| name.eq_ignore_ascii_case(&new_public))
        {
            self.control_names.push(new_public.clone());
        }

        let property_moves: Vec<(String, String)> = self
            .properties
            .iter()
            .filter_map(|((ctrl, prop), value)| {
                if ctrl.eq_ignore_ascii_case(&previous_public) {
                    Some((prop.clone(), value.clone()))
                } else {
                    None
                }
            })
            .collect();
        self.properties
            .retain(|(ctrl, _), _| !ctrl.eq_ignore_ascii_case(&previous_public));
        for (prop, value) in property_moves {
            self.properties.insert((new_public.clone(), prop), value);
        }

        let event_moves: Vec<(String, Value)> = self
            .event_handlers
            .iter()
            .filter_map(|(key, callback)| {
                let (ctrl, event) = key.rsplit_once('.')?;
                if ctrl.eq_ignore_ascii_case(&previous_public) {
                    Some((event.to_string(), callback.clone()))
                } else {
                    None
                }
            })
            .collect();
        self.event_handlers.retain(|key, _| {
            key.rsplit_once('.')
                .map(|(ctrl, _)| !ctrl.eq_ignore_ascii_case(&previous_public))
                .unwrap_or(true)
        });
        for (event, callback) in event_moves {
            self.event_handlers
                .insert(format!("{}.{}", new_public, event), callback);
        }
    }

    /// Find a `RecordingCanvas` for `control`. Resolution order:
    ///
    /// 1. **Canvas widget on the form** — if a `vybe_widgets::Canvas`
    ///    widget with the requested name exists, return its own
    ///    recording. This is the canonical case for explicit Canvas
    ///    widgets the user added to the form.
    ///
    /// 2. **Overlay recording** — for any other control (Button, Label,
    ///    Form, …), an entry is created in `overlay_canvases` on first
    ///    access. The form's render loop replays this overlay through
    ///    the widget's `paint_overlay` hook each frame, drawing on top
    ///    of the standard widget chrome.
    ///
    /// Returns a `&mut RecordingCanvas` borrowed from one of the two
    /// sources. Always succeeds — for an unknown control name, the
    /// overlay map gains a new entry.
    pub fn find_canvas_mut(&mut self, control: &str) -> &mut RecordingCanvas {
        let name = self.resolve_control_name(control);
        // Step 1: search child widgets for a Canvas widget with this name.
        // We use a raw-pointer trick to avoid the borrow-checker complaining
        // about returning a borrow that depends on a temporary closure
        // result. The lifetimes are sound — the &mut comes from
        // `self.form.controls`, which `self` owns, and we hand it back to
        // the caller as `&mut self.???`-flavoured.
        //
        // Walk the form's controls and downcast each one. If we find a
        // matching Canvas widget, return its inner recording.
        let canvas_ptr: Option<*mut RecordingCanvas> = {
            let widgets = self.form.controls_mut();
            let mut found: Option<*mut RecordingCanvas> = None;
            for w in widgets.iter_mut() {
                if w.name() == name {
                    if let Some(any) = w.as_any_mut() {
                        if let Some(c) = any.downcast_mut::<CanvasWidget>() {
                            let p: *mut RecordingCanvas = c.canvas_mut();
                            found = Some(p);
                            break;
                        }
                    }
                }
            }
            found
        };
        if let Some(p) = canvas_ptr {
            // Safety: the pointer borrows from `self.form.controls`. The
            // borrow is valid for as long as `self` lives because no
            // subsequent code in this function modifies the controls vec.
            return unsafe { &mut *p };
        }
        // Step 2: fall through to the overlay map.
        self.overlay_canvases.entry(name).or_default()
    }

    /// Register an event handler: key = "controlname.eventname".
    /// Control name is stored as-is — the language compiler is responsible for
    /// case normalisation (VB lowercases before calling, C# passes original case).
    /// Event name is lowercased since it is always a fixed word ("click", etc.).
    pub fn register_event(&mut self, control: &str, event: &str, callback: Value) {
        let public = self.public_control_name(control);
        let key = format!("{}.{}", public, event.to_lowercase());
        if std::env::var("VYBE_GUI_TRACE")
            .map(|v| matches!(v.as_str(), "1" | "true" | "TRUE" | "yes" | "YES"))
            .unwrap_or(false)
        {
            eprintln!("[gui-host] register_event key={}", key);
        }
        self.event_handlers.insert(key, callback);
    }

    /// Look up an event handler. Exact control match wins; then case-insensitive
    /// fallback is used for VB-style callers.
    pub fn get_event_handler(&self, control: &str, event: &str) -> Option<&Value> {
        let event_lower = event.to_lowercase();
        let public = self.public_control_name(control);
        let exact = format!("{}.{}", public, event_lower);
        if let Some(cb) = self.event_handlers.get(&exact) {
            return Some(cb);
        }

        self.event_handlers.iter().find_map(|(k, v)| {
            let (ctrl, ev) = k.rsplit_once('.')?;
            if ev == event_lower && ctrl.eq_ignore_ascii_case(control) {
                Some(v)
            } else {
                None
            }
        })
    }

    /// List all registered event keys (for debugging).
    pub fn event_keys(&self) -> Vec<String> {
        self.event_handlers.keys().cloned().collect()
    }

    pub fn track_live_control_name(&mut self, physical_name: &str, public_name: &str) {
        let public = public_name.trim().to_lowercase();
        if public.is_empty() {
            return;
        }
        if !self
            .control_names
            .iter()
            .any(|name| name.eq_ignore_ascii_case(&public))
        {
            self.control_names.push(public.clone());
        }
        self.control_aliases
            .insert(public.clone(), physical_name.to_string());
        self.control_public_names
            .insert(physical_name.to_string(), public);
    }

    fn hide_root_form_entries(&mut self) {
        self.control_names.retain(|name| {
            !self.control_public_names.iter().any(|(physical, public)| {
                physical.starts_with("Form_") && public.eq_ignore_ascii_case(name)
            })
        });
    }

    /// Create a widget from a control type name and add it to the form.
    pub fn add_widget(
        &mut self,
        type_name: &str,
        name: &str,
        text: &str,
        x: i32,
        y: i32,
        w: i32,
        h: i32,
    ) {
        // Pass original-case name to the widget so ButtonClicked events carry the original spelling.
        // Lookup keys (event handlers, control_names membership check) already lowercase at call time.
        let widget = make_widget(type_name, name, text, w as f32, h as f32);
        self.form
            .add_boxed_control(widget, x as f32, y as f32, w as f32, h as f32);
        self.track_live_control_name(name, name);
        self.hide_root_form_entries();
    }

    /// Declarative (Flutter) path: create a control widget and stage it into
    /// the widget tree under `parent` (a layout panel, or the form itself),
    /// letting vybe_widgets own nesting + flow layout. Contrast `add_widget`,
    /// which flat-adds at absolute coords (WinForms/VCL adapters).
    pub fn stage_control(
        &mut self,
        type_name: &str,
        name: &str,
        text: &str,
        w: i32,
        h: i32,
        parent: &str,
        parent_is_form: bool,
    ) {
        let widget = make_widget(type_name, name, text, w as f32, h as f32);
        self.form.stage_control(name, widget, parent, parent_is_form);
        self.track_live_control_name(name, name);
    }

    pub fn seed_form_identity(&mut self, name: &str, title: &str) {
        self.properties
            .insert((name.to_string(), "name".into()), name.to_string());
        self.properties
            .insert((name.to_string(), "text".into()), title.to_string());
        self.control_public_names
            .insert(name.to_string(), name.to_string());
    }

    /// Set a property on a control by name — directly updates the widget
    /// AND mirrors to the property store. The mirror lets callers query
    /// `get_property` for any control (including the form itself, which
    /// isn't represented as a child widget).
    /// Enabled GUI timers as `(control_name, interval_ms, handler)`. A timer is a
    /// control with a registered `Timer`/`Tick` event handler; `Enabled` defaults
    /// to true unless explicitly false, `Interval` defaults to 1000 ms. Drives
    /// `TTimer.OnTimer` / `System.Windows.Forms.Timer.Tick` from the event loop.
    pub fn active_timers(&self) -> Vec<(String, u64, Value)> {
        let mut out = Vec::new();
        for (key, handler) in &self.event_handlers {
            let Some((name, ev)) = key.rsplit_once('.') else { continue };
            if ev != "timer" && ev != "tick" {
                continue;
            }
            let prop = |p: &str| {
                self.properties
                    .iter()
                    .find(|((c, k), _)| k == p && c.eq_ignore_ascii_case(name))
                    .map(|(_, v)| v.clone())
            };
            let enabled = prop("enabled")
                .map(|v| !matches!(v.as_str(), "false" | "False" | "0" | ""))
                .unwrap_or(true);
            if !enabled {
                continue;
            }
            let interval = prop("interval")
                .and_then(|v| v.trim().parse::<u64>().ok())
                .filter(|ms| *ms > 0)
                .unwrap_or(1000);
            out.push((name.to_string(), interval, handler.clone()));
        }
        out
    }

    pub fn set_property(&mut self, control: &str, property: &str, value: &str) {
        let name = self.resolve_control_name(control);
        let public = self.public_control_name(control);
        let prop_lower = property.to_lowercase();
        // Always mirror to the property store first.
        self.properties
            .insert((public, prop_lower.clone()), value.to_string());
        match prop_lower.as_str() {
            "text" => {
                self.form
                    .send_command(&name, &WidgetCommand::SetText(value.to_string()));
            }
            "flex" => {
                let f = value.parse::<f32>().unwrap_or(1.0);
                self.form.send_command(&name, &WidgetCommand::SetFlex(f));
            }
            "enabled" => {
                let enabled = !matches!(value, "false" | "False" | "0" | "");
                self.form
                    .send_command(&name, &WidgetCommand::SetEnabled(enabled));
            }
            "visible" => {
                let visible = !matches!(value, "false" | "False" | "0" | "");
                self.form
                    .send_command(&name, &WidgetCommand::SetVisible(visible));
            }
            "readonly" => {
                let ro = matches!(value, "true" | "True" | "1");
                self.form.send_command(
                    &name,
                    &WidgetCommand::Custom("SetReadOnly".into(), CommandValue::Bool(ro)),
                );
            }
            // Semantic property names route to the EXISTING typed commands the
            // controls already handle (Slider/ProgressBar SetValue, Checkbox/
            // Radio SetChecked, Combo/List/Tabs SetSelectedIndex) — the Flutter
            // adapter forwards `value`/`checked`/`selectedindex` here.
            "value" => {
                // `value` is numeric on a Slider/ProgressBar, boolean on a
                // Checkbox/Switch — the control owns the state either way.
                if let Ok(n) = value.parse::<f64>() {
                    self.form.send_command(&name, &WidgetCommand::SetValue(n));
                } else {
                    let c = matches!(value, "true" | "True" | "1");
                    self.form
                        .send_command(&name, &WidgetCommand::SetChecked(c));
                }
            }
            "checked" | "ischecked" | "selected" => {
                let c = matches!(value, "true" | "True" | "1");
                self.form
                    .send_command(&name, &WidgetCommand::SetChecked(c));
            }
            "selectedindex" => {
                if let Ok(i) = value.parse::<usize>() {
                    self.form
                        .send_command(&name, &WidgetCommand::SetSelectedIndex(i));
                }
            }
            // Item-list population for combobox/listbox/tabcontrol/datagrid —
            // routes to the AddItem/ClearItems commands the widgets already
            // implement. The Flutter adapter clears then adds each item's
            // caption when it realizes a DropdownButton/ListView/TabBar.
            "clearitems" => {
                self.form.send_command(&name, &WidgetCommand::ClearItems);
            }
            "additem" => {
                self.form
                    .send_command(&name, &WidgetCommand::AddItem(value.to_string()));
            }
            // Visual/box properties. These carry NUMBERS (a size, a padding, a
            // packed ARGB colour), so they go over as `CommandValue::Number`
            // when they parse — a `Text` payload would force every widget to
            // re-parse. Colours stay textual so names ("red") and `#RRGGBB`
            // still work; the widgets accept either form.
            "width" | "height" | "padding" | "spacing" | "fontsize" => {
                let cmd = format!("Set{}", capitalize_first(&prop_lower));
                let payload = match value.trim().parse::<f64>() {
                    Ok(n) => CommandValue::Number(n),
                    Err(_) => CommandValue::Text(value.to_string()) };
                self.form
                    .send_command(&name, &WidgetCommand::Custom(cmd, payload));
            }
            "color" | "backcolor" | "backgroundcolor" => {
                self.form.send_command(
                    &name,
                    &WidgetCommand::Custom(
                        "SetBackColor".into(),
                        CommandValue::Text(value.to_string()),
                    ),
                );
            }
            "forecolor" | "textcolor" => {
                self.form.send_command(
                    &name,
                    &WidgetCommand::Custom(
                        "SetForeColor".into(),
                        CommandValue::Text(value.to_string()),
                    ),
                );
            }
            other => {
                self.form.send_command(
                    &name,
                    &WidgetCommand::Custom(
                        format!("Set{}", capitalize_first(other)),
                        CommandValue::Text(value.to_string()),
                    ),
                );
            }
        }
    }

    /// Get a property from a control by name. Reads from the property store
    /// first (covers form-level properties and any prior set_property write),
    /// falling back to a widget query for properties that are managed by
    /// the live widget rather than the store.
    pub fn get_property(&mut self, control: &str, property: &str) -> String {
        let name = self.resolve_control_name(control);
        let public = self.public_control_name(control);
        let prop_lower = property.to_lowercase();
        if let Some(v) = self.properties.get(&(public, prop_lower.clone())) {
            return v.clone();
        }
        match prop_lower.as_str() {
            "text" => {
                let result = self.form.send_command(&name, &WidgetCommand::GetText);
                match result {
                    CommandValue::Text(s) => s,
                    _ => String::new() }
            }
            _ => {
                let result = self.form.send_command(
                    &name,
                    &WidgetCommand::Custom(
                        format!("Get{}", capitalize_first(property)),
                        CommandValue::None,
                    ),
                );
                match result {
                    CommandValue::Text(s) => s,
                    _ => String::new() }
            }
        }
    }
}

fn capitalize_first(s: &str) -> String {
    let mut c = s.chars();
    match c.next() {
        None => String::new(),
        Some(f) => f.to_uppercase().collect::<String>() + c.as_str() }
}

/// Create a boxed PanelWidget from a type name string.
fn make_widget(type_name: &str, name: &str, text: &str, w: f32, h: f32) -> Box<dyn PanelWidget> {
    match type_name.to_lowercase().as_str() {
        "canvas" | "paintbox" => {
            // The Canvas widget is the bare drawable surface. PaintBox
            // is the .NET BCL/FCL alias the dotnet wrapper uses.
            use vybe_widgets::layout::LayoutRect;
            let mut c = vybe_widgets::Canvas::new().with_name(name);
            <vybe_widgets::Canvas as PanelWidget>::set_rect(
                &mut c,
                LayoutRect::new(0.0, 0.0, w, h),
            );
            Box::new(c)
        }
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
            t.cursor = t.value.len();
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
        // Horizontal flow — Flutter `Row`. (Vertical `FlowLayoutPanel` default
        // serves Column/Scaffold.)
        "hflowlayoutpanel" => Box::new(
            FlowLayoutPanel::new()
                .with_name(name)
                .with_direction(vybe_widgets::flow_layout::FlowDirection::LeftToRight),
        ),
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

#[cfg(test)]
mod adapter_tests {
    //! Per-widget adapter verification: build the control the way the Flutter
    //! realizer does (`stage_control`), forward a declared property the way the
    //! adapter does (`set_property`), then assert the REAL control's state — not
    //! the mirror store — reflects it. State lives in the widget.
    use super::*;
    use vybe_widgets::layout::{LayoutRect, MouseButton, MouseEvent, MouseEventKind};
    use vybe_widgets::{WidgetCommand, WidgetEvent};

    /// Checkbox, fully wired: `Checkbox(value: true/false)` reflects onto the
    /// real checkbox control's checked state.
    #[test]
    fn checkbox_value_applies_to_control() {
        let mut gui = GuiState::new();
        gui.stage_control("checkbox", "cb1", "", 20, 20, "", true);

        // Flutter `Checkbox(value: true)` → adapter forwards value=true.
        gui.set_property("cb1", "value", "true");
        assert!(
            matches!(
                gui.form.send_command("cb1", &WidgetCommand::GetValue),
                CommandValue::Bool(true)
            ),
            "value:true must check the real control"
        );

        gui.set_property("cb1", "value", "false");
        assert!(
            matches!(
                gui.form.send_command("cb1", &WidgetCommand::GetValue),
                CommandValue::Bool(false)
            ),
            "value:false must uncheck the real control"
        );
    }

    /// Checkbox `onChanged`: a click on the real control emits its toggle event
    /// (which the host routes to the Dart handler → `setState`). State stays in
    /// the widget — the click flips the control's own checked state.
    #[test]
    fn checkbox_click_emits_toggle_event() {
        let mut gui = GuiState::new();
        gui.stage_control("checkbox", "cb1", "", 20, 20, "", true);
        // Lay the form out so the staged checkbox gets a hit-testable rect.
        gui.form.set_rect(LayoutRect::new(0.0, 0.0, 200.0, 200.0));

        let click = MouseEvent {
            x: 10.0,
            y: 10.0,
            kind: MouseEventKind::Press(MouseButton::Left),
            cmd: false,
            shift: false,
            alt: false };
        gui.form.handle_mouse(&click);
        let events = gui.form.drain_events();
        assert!(
            events
                .iter()
                .any(|e| matches!(e, WidgetEvent::CheckboxToggled(name, true) if name == "cb1")),
            "clicking the checkbox must emit CheckboxToggled(checked=true); got {events:?}"
        );
    }

    fn slider_actual(gui: &mut GuiState, name: &str) -> f64 {
        match gui.form.send_command(name, &WidgetCommand::GetValue) {
            CommandValue::Number(n) => n,
            other => panic!("slider GetValue returned {other:?}") }
    }

    /// Slider, fully wired: `Slider(value:, min:, max:)` positions the real
    /// trackbar at the actual value, regardless of the order fields arrive in
    /// (the control stores the actual value and derives the fraction).
    #[test]
    fn slider_value_and_bounds_apply_to_control() {
        // Standard 0..100 range.
        let mut gui = GuiState::new();
        gui.stage_control("trackbar", "sl1", "", 200, 20, "", true);
        gui.set_property("sl1", "value", "40");
        gui.set_property("sl1", "min", "0");
        gui.set_property("sl1", "max", "100");
        assert!(
            (slider_actual(&mut gui, "sl1") - 40.0).abs() < 0.001,
            "Slider(value:40) must sit at 40, not the 0..1 fraction"
        );

        // Custom range with the value set BEFORE the bounds — must still land at
        // the actual value (order-independence is the whole point).
        let mut gui2 = GuiState::new();
        gui2.stage_control("trackbar", "sl2", "", 200, 20, "", true);
        gui2.set_property("sl2", "value", "50");
        gui2.set_property("sl2", "min", "10");
        gui2.set_property("sl2", "max", "90");
        assert!(
            (slider_actual(&mut gui2, "sl2") - 50.0).abs() < 0.001,
            "Slider(value:50, min:10, max:90) must sit at 50"
        );
    }

    /// Progress indicator, fully wired: `LinearProgressIndicator(value: 0.6)`
    /// fills the real progress bar to 0.6 (Flutter's value is already 0..1).
    #[test]
    fn progress_value_applies_to_control() {
        let mut gui = GuiState::new();
        gui.stage_control("progressbar", "pb1", "", 200, 8, "", true);
        gui.set_property("pb1", "value", "0.6");
        assert!(
            (slider_actual(&mut gui, "pb1") - 0.6).abs() < 0.001,
            "LinearProgressIndicator(value:0.6) must fill to 0.6"
        );
    }

    /// TextField, control side: forwarded text lands on the real text box.
    /// (The `controller.text` → `text` extraction is Dart-side in the realizer,
    /// verified end-to-end via the debugger on the gallery.)
    #[test]
    fn textfield_text_applies_to_control() {
        let mut gui = GuiState::new();
        gui.stage_control("textbox", "tf1", "", 200, 24, "", true);
        gui.set_property("tf1", "text", "hello");
        match gui.form.send_command("tf1", &WidgetCommand::GetText) {
            CommandValue::Text(s) => assert_eq!(s, "hello"),
            other => panic!("TextField GetText returned {other:?}") }
    }

    /// Radio, control side: the adapter reflects `value == groupValue` as a
    /// programmatic `checked` — the radio shows selected and does NOT queue a
    /// spurious `onChanged` (only a user click emits `RadioSelected`).
    #[test]
    fn radio_programmatic_select_is_silent() {
        let mut gui = GuiState::new();
        gui.stage_control("radiobutton", "rb1", "", 20, 20, "", true);
        gui.set_property("rb1", "checked", "true");
        assert!(
            matches!(
                gui.form.send_command("rb1", &WidgetCommand::GetValue),
                CommandValue::Bool(true)
            ),
            "value==groupValue must select the radio"
        );
        let events = gui.form.drain_events();
        assert!(
            events.is_empty(),
            "programmatic radio select must not emit onChanged; got {events:?}"
        );
    }
}
