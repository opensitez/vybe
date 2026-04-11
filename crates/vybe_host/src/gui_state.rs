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
    Canvas as CanvasWidget,
};
use vybe_widgets::canvas::RecordingCanvas;

/// Holds the live widget form + event callbacks.
/// Created before VM runs, shared with host fns via `Arc<Mutex<>>`.
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
    /// Property store: every `set_property` write is mirrored here keyed by
    /// `(lowercased control name, lowercased property)`. Used both as the
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
    /// Per-control overlay recordings — keyed by lowercased control
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
    pub overlay_canvases: HashMap<String, RecordingCanvas>,
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
            properties: HashMap::new(),
            needs_repaint: false,
            front_requested: false,
            overlay_canvases: HashMap::new(),
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
        let name = control.to_lowercase();
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

    /// Register an event handler: key = "controlname.eventname" (both lowercased).
    /// Lowercasing both ensures VB (case-insensitive) and JS (case-sensitive)
    /// both match — events are always "Click", "TextChanged" etc.
    pub fn register_event(&mut self, control: &str, event: &str, callback: Value) {
        let key = format!("{}.{}", control.to_lowercase(), event.to_lowercase());
        self.event_handlers.insert(key, callback);
    }

    /// Look up an event handler. Both control and event are lowercased.
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

    /// Set a property on a control by name — directly updates the widget
    /// AND mirrors to the property store. The mirror lets callers query
    /// `get_property` for any control (including the form itself, which
    /// isn't represented as a child widget).
    pub fn set_property(&mut self, control: &str, property: &str, value: &str) {
        let name = control.to_lowercase();
        let prop_lower = property.to_lowercase();
        // Always mirror to the property store first.
        self.properties.insert((name.clone(), prop_lower.clone()), value.to_string());
        match prop_lower.as_str() {
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

    /// Get a property from a control by name. Reads from the property store
    /// first (covers form-level properties and any prior set_property write),
    /// falling back to a widget query for properties that are managed by
    /// the live widget rather than the store.
    pub fn get_property(&mut self, control: &str, property: &str) -> String {
        let name = control.to_lowercase();
        let prop_lower = property.to_lowercase();
        if let Some(v) = self.properties.get(&(name.clone(), prop_lower.clone())) {
            return v.clone();
        }
        match prop_lower.as_str() {
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
