//! `vybe:gui` host module — registers GUI host functions on the VM.
//!
//! When the `gui` feature is enabled, host functions directly create
//! `vybe_widgets` widgets and store them in a shared `GuiState`.

#[cfg(feature = "gui")]
mod gui_impl {

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use crate::gui_state::GuiState;

pub fn register(
    vm: &mut VM,
    gui: Arc<Mutex<GuiState>>,
) {
    // Form creation
    vm.register_host_fn("vybe:gui", "createForm", {
        let gui = gui.clone();
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let title = str_arg(args, 0, "Form1");
            let name = title.clone();
            let mut g = gui.lock().unwrap();
            g.form = vybe_widgets::Form::new(&title);
            Value::String(Arc::from(name.as_str()))
        })
    });

    vm.register_host_fn("vybe:gui", "newForm", {
        let gui = gui.clone();
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let title = str_arg(args, 0, "Form1");
            let name = title.clone();
            { let mut g = gui.lock().unwrap(); g.form = vybe_widgets::Form::new(&title); }
            let mut obj = vybe_bytecode::value::Object::new();
            obj.properties.insert("__control_type".into(), Value::String(Arc::from("Form")));
            obj.properties.insert("__control_name".into(), Value::String(Arc::from(name.as_str())));
            obj.properties.insert("name".into(), Value::String(Arc::from(name.as_str())));
            obj.properties.insert("text".into(), Value::String(Arc::from(title.as_str())));
            obj.properties.insert("width".into(), Value::F64(800.0));
            obj.properties.insert("height".into(), Value::F64(600.0));
            Value::Object(Arc::new(Mutex::new(obj)))
        })
    });

    // Add control to form
    vm.register_host_fn("vybe:gui", "controlsAdd", {
        let gui = gui.clone();
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            // args[0] = parent container, args[1] = child control
            // Compute parent's absolute offset so child is positioned inside parent.
            let (parent_abs_x, parent_abs_y) = if let Some(Value::Object(parent_obj)) = args.first() {
                let po = parent_obj.lock().unwrap();
                let (px, py) = if let Some(Value::Object(loc)) = po.properties.get("location") {
                    let loc = loc.lock().unwrap();
                    (loc.properties.get("x").map(|v| v.as_f64() as i32).unwrap_or(0),
                     loc.properties.get("y").map(|v| v.as_f64() as i32).unwrap_or(0))
                } else {
                    (po.properties.get("left").map(|v| v.as_f64() as i32).unwrap_or(0),
                     po.properties.get("top").map(|v| v.as_f64() as i32).unwrap_or(0))
                };
                // Walk up __parent chain to accumulate offsets for deeply nested containers
                let mut abs_x = px;
                let mut abs_y = py;
                let mut cur = po.properties.get("__parent").cloned();
                drop(po);
                while let Some(Value::Object(ancestor)) = cur {
                    let anc = ancestor.lock().unwrap();
                    let (ax, ay) = if let Some(Value::Object(loc)) = anc.properties.get("location") {
                        let loc = loc.lock().unwrap();
                        (loc.properties.get("x").map(|v| v.as_f64() as i32).unwrap_or(0),
                         loc.properties.get("y").map(|v| v.as_f64() as i32).unwrap_or(0))
                    } else {
                        (anc.properties.get("left").map(|v| v.as_f64() as i32).unwrap_or(0),
                         anc.properties.get("top").map(|v| v.as_f64() as i32).unwrap_or(0))
                    };
                    abs_x += ax;
                    abs_y += ay;
                    cur = anc.properties.get("__parent").cloned();
                }
                (abs_x, abs_y)
            } else {
                (0, 0)
            };

            if let Some(Value::Object(obj)) = args.get(1) {
                // Record parent reference on the child for deep nesting
                if let Some(parent_val) = args.first() {
                    obj.lock().unwrap().properties.insert("__parent".into(), parent_val.clone());
                }
                let o = obj.lock().unwrap();
                let control_type = o.properties.get("__control_type")
                    .map(|v| format!("{}", v)).unwrap_or_else(|| "Button".into());
                let control_name = o.properties.get("name")
                    .or_else(|| o.properties.get("__control_name"))
                    .map(|v| format!("{}", v)).unwrap_or_else(|| "ctrl".into());
                let text = o.properties.get("text")
                    .map(|v| format!("{}", v)).unwrap_or_default();
                let left = o.properties.get("left").map(|v| v.as_f64() as i32).unwrap_or(0);
                let top = o.properties.get("top").map(|v| v.as_f64() as i32).unwrap_or(0);
                let width = o.properties.get("width").map(|v| v.as_f64() as i32).unwrap_or(100);
                let height = o.properties.get("height").map(|v| v.as_f64() as i32).unwrap_or(30);
                let (left, top) = if let Some(Value::Object(loc)) = o.properties.get("location") {
                    let loc = loc.lock().unwrap();
                    (loc.properties.get("x").map(|v| v.as_f64() as i32).unwrap_or(left),
                     loc.properties.get("y").map(|v| v.as_f64() as i32).unwrap_or(top))
                } else { (left, top) };
                let (width, height) = if let Some(Value::Object(sz)) = o.properties.get("size") {
                    let sz = sz.lock().unwrap();
                    (sz.properties.get("width").map(|v| v.as_f64() as i32).unwrap_or(width),
                     sz.properties.get("height").map(|v| v.as_f64() as i32).unwrap_or(height))
                } else { (width, height) };
                let props: Vec<(String, String)> = o.properties.iter()
                    .filter(|(k, _)| !k.starts_with("__") && !matches!(k.as_str(),
                        "name" | "left" | "top" | "width" | "height" | "text"
                        | "location" | "size" | "show" | "close" | "focus" | "hide" | "showdialog"))
                    .filter_map(|(k, v)| {
                        let val_str = value_to_property_string(v)?;
                        Some((capitalize_first(k), val_str))
                    })
                    .collect();
                drop(o);
                // Add widget at absolute position (child local + parent absolute)
                let abs_left = left + parent_abs_x;
                let abs_top = top + parent_abs_y;
                let mut g = gui.lock().unwrap();
                g.add_widget(&control_type, &control_name, &text, abs_left, abs_top, width, height);
                let name_lower = control_name.to_lowercase();
                for (prop, val) in props {
                    apply_property(&mut g.form, &name_lower, &prop, &val);
                }
            }
            Value::Null
        })
    });

    vm.register_host_fn("vybe:gui", "addControl", {
        let gui = gui.clone();
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let _form_name = str_arg(args, 0, "Form1");
            let control_type = str_arg(args, 1, "Button");
            let control_name = str_arg(args, 2, "control1");
            let left = i32_arg(args, 3, 0);
            let top = i32_arg(args, 4, 0);
            let width = i32_arg(args, 5, 100);
            let height = i32_arg(args, 6, 30);
            gui.lock().unwrap().add_widget(&control_type, &control_name, "", left, top, width, height);
            Value::String(Arc::from(control_name.as_str()))
        })
    });

    // Property set/get — directly update the widget
    vm.register_host_fn("vybe:gui", "setProperty", {
        let gui = gui.clone();
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let control = str_arg(args, 0, "");
            let property = str_arg(args, 1, "");
            let val_str = args.get(2).map(|v| format!("{}", v)).unwrap_or_default();
            gui.lock().unwrap().set_property(&control, &property, &val_str);
            Value::Null
        })
    });

    vm.register_host_fn("vybe:gui", "controlSetProperty", {
        let gui = gui.clone();
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let control_name = {
                    let o = obj.lock().unwrap();
                    o.properties.get("__control_name")
                        .map(|v| format!("{}", v)).unwrap_or_default()
                };
                let property = str_arg(args, 1, "");
                let val = args.get(2).cloned().unwrap_or(Value::Null);
                let val_str = format!("{}", val);
                let prop_lower = property.to_lowercase();
                obj.lock().unwrap().properties.insert(prop_lower.clone(), val.clone());
                if prop_lower == "name" {
                    obj.lock().unwrap().properties.insert("__control_name".into(), val);
                }
                gui.lock().unwrap().set_property(&control_name, &property, &val_str);
            }
            Value::Null
        })
    });

    vm.register_host_fn("vybe:gui", "getProperty", {
        let gui = gui.clone();
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let control = str_arg(args, 0, "");
            let property = str_arg(args, 1, "");
            let val = gui.lock().unwrap().get_property(&control, &property);
            if val.is_empty() { Value::Null } else { Value::String(Arc::from(val.as_str())) }
        })
    });

    // Event registration
    vm.register_host_fn("vybe:gui", "onEvent", {
        let gui = gui.clone();
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let control = str_arg(args, 0, "");
            let event = str_arg(args, 1, "");
            let callback = args.get(2).cloned().unwrap_or(Value::Null);
            gui.lock().unwrap().register_event(&control, &event, callback);
            Value::Null
        })
    });

    vm.register_host_fn("vybe:gui", "addHandler", {
        let gui = gui.clone();
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let ctrl = str_arg(args, 0, "");
            let event = str_arg(args, 1, "");
            if let Some(callback) = args.get(2) {
                gui.lock().unwrap().register_event(&ctrl, &event, callback.clone());
            }
            Value::Null
        })
    });

    vm.register_host_fn("vybe:gui", "removeHandler", Box::new(|_ctx, _| Value::Null));

    // Form lifecycle
    vm.register_host_fn("vybe:gui", "showForm", {
        let gui = gui.clone();
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let mut g = gui.lock().unwrap();
            g.should_run = true;
            if let Some(obj) = args.first().cloned() { g.form_object = Some(obj); }
            Value::Null
        })
    });

    vm.register_host_fn("vybe:gui", "runApplication", {
        let gui = gui.clone();
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let mut g = gui.lock().unwrap();
            g.should_run = true;
            if let Some(obj) = args.first().cloned() {
                if let Value::Object(o) = &obj {
                    let o = o.lock().unwrap();
                    if let Some(w) = o.properties.get("width") { g.width = w.as_f64() as u32; }
                    if let Some(h) = o.properties.get("height") { g.height = h.as_f64() as u32; }
                }
                g.form_object = Some(obj);
            }
            Value::Null
        })
    });

    vm.register_host_fn("vybe:gui", "closeForm", {
        let gui = gui.clone();
        Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
            gui.lock().unwrap().close_requested = true;
            Value::Null
        })
    });

    // MsgBox — push to GuiState pending_dialogs; runner drains and shows dialog
    vm.register_host_fn("vybe:gui", "msgBox", {
        let gui = gui.clone();
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let text = str_arg(args, 0, "");
            let title = str_arg(args, 1, "");
            gui.lock().unwrap().pending_dialogs.push((text, title));
            Value::Null
        })
    });

    vm.register_host_fn("vybe:gui", "noop", Box::new(|_ctx, _| Value::Null));

    // Control constructors
    let control_types = [
        "Button", "Label", "TextBox", "CheckBox", "RadioButton", "ComboBox",
        "ListBox", "Panel", "GroupBox", "TabControl", "TabPage", "DataGridView",
        "ProgressBar", "TrackBar", "NumericUpDown", "DateTimePicker", "RichTextBox",
        "PictureBox", "MenuStrip", "ToolStrip", "StatusStrip", "SplitContainer",
        "FlowLayoutPanel", "TableLayoutPanel", "LinkLabel", "MaskedTextBox",
        "HScrollBar", "VScrollBar", "MonthCalendar", "BindingNavigator",
        "BindingSource", "DataSet", "DataTable", "DataAdapter",
        "OpenFileDialog", "SaveFileDialog", "FontDialog", "ColorDialog",
        "FolderBrowserDialog", "PrintDialog", "PrintPreviewDialog",
        "ListView", "WebBrowser", "ContextMenuStrip",
        "Timer", "ImageList", "ToolTip",
        "NotifyIcon", "ErrorProvider", "HelpProvider", "BackgroundWorker",
        "Form", "TreeView",
    ];

    // ── Control lifecycle host fns ─────────────────────────────────────────
    //
    // These are bound on every Control via the dotnet class wrapper layer
    // (see `compiler_common::dotnet::classes::control::CONTROL_METHODS`).
    // The method thunk passes `this` as the first arg, so we read
    // `this.__control_name` to find the target control and route the call
    // through the existing `WidgetCommand` interface.
    //
    // Form-specific shortcuts: when `this.__type == "Form"` we ALSO trigger
    // `should_run = true` for `Show()` (real .NET semantics: opening a Form
    // for the first time enters its message loop) and `close_requested = true`
    // for `Close()`. Child controls don't get these side effects.

    let gui_show = gui.clone();
    vm.register_host_fn("vybe:gui", "__ctrl_show", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let (name, is_form) = read_this_identity(args);
        let mut g = gui_show.lock().unwrap();
        if is_form {
            g.should_run = true;
        }
        if !name.is_empty() {
            g.set_property(&name, "Visible", "true");
        }
        Value::Null
    }));
    let gui_close = gui.clone();
    vm.register_host_fn("vybe:gui", "__ctrl_close", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let (_name, is_form) = read_this_identity(args);
        let mut g = gui_close.lock().unwrap();
        if is_form {
            g.close_requested = true;
        }
        // Non-form `Close` (rare — most controls don't expose Close in real
        // .NET; only forms do) is a no-op.
        Value::Null
    }));
    let gui_hide = gui.clone();
    vm.register_host_fn("vybe:gui", "__ctrl_hide", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let (name, _is_form) = read_this_identity(args);
        if !name.is_empty() {
            gui_hide.lock().unwrap().set_property(&name, "Visible", "false");
        }
        Value::Null
    }));
    let gui_focus = gui.clone();
    vm.register_host_fn("vybe:gui", "__ctrl_focus", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let (name, _) = read_this_identity(args);
        if !name.is_empty() {
            let mut g = gui_focus.lock().unwrap();
            g.form.send_command(&name, &vybe_widgets::WidgetCommand::Focus);
        }
        Value::Null
    }));
    // Refresh / Invalidate / Update all map to "request a repaint". Real
    // .NET distinguishes them (Refresh = Invalidate + Update), but our
    // renderer is fully repainting every frame, so the distinction
    // collapses to a single needs_repaint flag.
    let gui_refresh = gui.clone();
    vm.register_host_fn("vybe:gui", "__ctrl_refresh", Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
        gui_refresh.lock().unwrap().needs_repaint = true;
        Value::Null
    }));
    let gui_invalidate = gui.clone();
    vm.register_host_fn("vybe:gui", "__ctrl_invalidate", Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
        gui_invalidate.lock().unwrap().needs_repaint = true;
        Value::Null
    }));
    let gui_update = gui.clone();
    vm.register_host_fn("vybe:gui", "__ctrl_update", Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
        gui_update.lock().unwrap().needs_repaint = true;
        Value::Null
    }));
    // BringToFront / SendToBack — relayed to the widget via Custom command.
    // Form delegates to the OS window manager (out of scope for the
    // headless renderer); child controls reorder within the form. We
    // capture the intent on the property store too so tests can verify
    // the call landed.
    let gui_btf = gui.clone();
    vm.register_host_fn("vybe:gui", "__ctrl_bring_to_front", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let (name, _) = read_this_identity(args);
        if !name.is_empty() {
            let mut g = gui_btf.lock().unwrap();
            g.form.send_command(&name, &vybe_widgets::WidgetCommand::Custom(
                "BringToFront".into(),
                vybe_widgets::CommandValue::None,
            ));
            g.properties.insert((name, "__zorder".into()), "front".into());
            g.needs_repaint = true;
        }
        Value::Null
    }));
    let gui_stb = gui.clone();
    vm.register_host_fn("vybe:gui", "__ctrl_send_to_back", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let (name, _) = read_this_identity(args);
        if !name.is_empty() {
            let mut g = gui_stb.lock().unwrap();
            g.form.send_command(&name, &vybe_widgets::WidgetCommand::Custom(
                "SendToBack".into(),
                vybe_widgets::CommandValue::None,
            ));
            g.properties.insert((name, "__zorder".into()), "back".into());
            g.needs_repaint = true;
        }
        Value::Null
    }));
    // Dispose: hide the control + drop its event handlers + clear any
    // pending draw commands targeting it. Real .NET also frees native
    // GDI handles, which we don't have.
    let gui_dispose = gui.clone();
    vm.register_host_fn("vybe:gui", "__ctrl_dispose", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let (name, _) = read_this_identity(args);
        if !name.is_empty() {
            let mut g = gui_dispose.lock().unwrap();
            g.set_property(&name, "Visible", "false");
            // Drop any event handlers keyed under "<name>.*"
            let prefix = format!("{}.", name);
            g.event_handlers.retain(|k, _| !k.starts_with(&prefix));
            // Drop any overlay canvas recording for this control
            g.overlay_canvases.remove(&name);
            g.needs_repaint = true;
        }
        Value::Null
    }));
    // Form.Activate — request the OS window be brought to the foreground.
    let gui_activate = gui.clone();
    vm.register_host_fn("vybe:gui", "__form_activate", Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
        gui_activate.lock().unwrap().front_requested = true;
        Value::Null
    }));
    // Form.CenterToScreen — compute the centered position from screen size
    // and set it on the form. Stored on the property store; the window
    // driver reads it on next show.
    let gui_center = gui.clone();
    vm.register_host_fn("vybe:gui", "__form_center_to_screen", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let (name, _) = read_this_identity(args);
        if !name.is_empty() {
            // Use a sensible default for screen size — the GUI backend can
            // override this when it knows the real resolution. 1920x1080 is
            // common enough for the centered-position calculation to land
            // on-screen for most users.
            const SCREEN_W: u32 = 1920;
            const SCREEN_H: u32 = 1080;
            let mut g = gui_center.lock().unwrap();
            let w = g.width;
            let h = g.height;
            let left = (SCREEN_W.saturating_sub(w) / 2) as i32;
            let top = (SCREEN_H.saturating_sub(h) / 2) as i32;
            g.set_property(&name, "Left", &left.to_string());
            g.set_property(&name, "Top", &top.to_string());
            g.front_requested = true;
        }
        Value::Null
    }));
    // ShowDialog — modal show. Sets should_run + close_requested
    // semantics is up to the runner; the return value is `DialogResult.OK = 1`
    // by convention, the same as the legacy stub. Real modal handling is a
    // separate workstream (it requires nested message loops).
    let gui_show_dlg = gui.clone();
    vm.register_host_fn("vybe:gui", "__dlg_showdialog", Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
        gui_show_dlg.lock().unwrap().should_run = true;
        Value::I32(1) // DialogResult.OK
    }));
    vm.register_host_fn("vybe:gui", "__dlg_show", Box::new(|_ctx, _| Value::I32(0)));

    let show_ref = host_fn_ref(vm, "__ctrl_show");
    let close_ref = host_fn_ref(vm, "__ctrl_close");
    let focus_ref = host_fn_ref(vm, "__ctrl_focus");
    let hide_ref = host_fn_ref(vm, "__ctrl_hide");
    let dlg_ref = host_fn_ref(vm, "__dlg_showdialog");

    for ct in control_types {
        let type_name = ct.to_string();
        let show = show_ref.clone();
        let close = close_ref.clone();
        let focus = focus_ref.clone();
        let hide = hide_ref.clone();
        let dlg = dlg_ref.clone();
        vm.register_host_fn("vybe:gui", &format!("new_{}", ct), Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
            // The vybe_widgets backing object. The dotnet class wrappers in
            // `compiler_common::dotnet::classes` consume this from inside the
            // ctor of `Form`/`Button`/etc. and copy `__control_name`,
            // `__control_type`, `show`, `close`, `focus`, `hide` onto the
            // class instance. Property setters (`__set_text`, `__set_name`,
            // `__set_width`, etc.) live in compiled bytecode chunks generated
            // by the dotnet class layer — they are NOT installed here.
            use std::sync::atomic::{AtomicU32, Ordering};
            static COUNTER: AtomicU32 = AtomicU32::new(1);
            let id = COUNTER.fetch_add(1, Ordering::Relaxed);
            let name = format!("{}_{}", type_name, id);
            let mut obj = vybe_bytecode::value::Object::new();
            obj.properties.insert("__control_type".into(), Value::String(Arc::from(type_name.as_str())));
            obj.properties.insert("__control_name".into(), Value::String(Arc::from(name.as_str())));
            obj.properties.insert("__type".into(), Value::String(Arc::from(type_name.as_str())));
            obj.properties.insert("name".into(), Value::String(Arc::from(name.as_str())));
            obj.properties.insert("width".into(), Value::F64(100.0));
            obj.properties.insert("height".into(), Value::F64(30.0));
            obj.properties.insert("left".into(), Value::F64(0.0));
            obj.properties.insert("top".into(), Value::F64(0.0));
            obj.properties.insert("show".into(), show.clone());
            obj.properties.insert("close".into(), close.clone());
            obj.properties.insert("focus".into(), focus.clone());
            obj.properties.insert("hide".into(), hide.clone());
            if matches!(type_name.as_str(),
                "OpenFileDialog" | "SaveFileDialog" | "FontDialog" | "ColorDialog"
                | "FolderBrowserDialog" | "PrintDialog" | "PrintPreviewDialog"
            ) {
                obj.properties.insert("showdialog".into(), dlg.clone());
            }
            Value::Object(Arc::new(Mutex::new(obj)))
        }));
    }
}

fn str_arg(args: &[Value], idx: usize, default: &str) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_else(|| default.into())
}

/// Extract `(__control_name, is_form)` from `args[0]`. Used by the
/// control lifecycle host fns to find the target control. `is_form` is
/// `true` when the object's `__type` is `"Form"` (or starts with
/// `"Form"` for user subclasses) — used to gate form-specific side
/// effects like `should_run = true` on `Show()`.
fn read_this_identity(args: &[Value]) -> (String, bool) {
    if let Some(Value::Object(obj)) = args.first() {
        let o = obj.lock().unwrap();
        let name = o.properties
            .get("__control_name")
            .map(|v| format!("{}", v).to_lowercase())
            .unwrap_or_default();
        let type_str = o.properties
            .get("__type")
            .map(|v| format!("{}", v))
            .unwrap_or_default();
        // Match `Form` exactly, AND user subclasses whose chain re-stamps
        // `__type` with the subclass name. Real .NET would test `is Form`,
        // but the type registry doesn't ship its full hierarchy down to
        // host fns yet — string match against the `Form` keyword is
        // good enough for the cases that matter (Show/Close).
        let is_form = type_str == "Form"
            || matches!(o.properties.get("__control_type"),
                Some(Value::String(s)) if s.as_ref() == "Form");
        (name, is_form)
    } else {
        (String::new(), false)
    }
}

fn i32_arg(args: &[Value], idx: usize, default: i32) -> i32 {
    args.get(idx).map(|v| v.as_f64() as i32).unwrap_or(default)
}

fn host_fn_ref(vm: &VM, name: &str) -> Value {
    let idx = *vm.host_registry.get(&("vybe:gui".into(), name.into())).unwrap();
    let mut o = vybe_bytecode::value::Object::new();
    o.kind = vybe_bytecode::value::ObjectKind::HostFunction(idx);
    Value::Object(Arc::new(Mutex::new(o)))
}

fn capitalize_first(s: &str) -> String {
    let mut c = s.chars();
    match c.next() {
        None => String::new(),
        Some(f) => f.to_uppercase().collect::<String>() + c.as_str(),
    }
}

/// Convert a VM Value into a string suitable for apply_property.
/// Returns None for values that shouldn't be passed as properties (functions, etc.).
fn value_to_property_string(v: &Value) -> Option<String> {
    match v {
        Value::String(s) => Some(s.to_string()),
        Value::F64(n) => Some(n.to_string()),
        Value::I32(n) => Some(n.to_string()),
        Value::Bool(b) => Some(if *b { "True".into() } else { "False".into() }),
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            // Color objects → extract "name" which holds "#RRGGBB" or named color
            if let Some(Value::String(t)) = o.properties.get("__type") {
                if t.as_ref() == "Color" {
                    if let Some(Value::String(name)) = o.properties.get("name") {
                        return Some(name.to_string());
                    }
                    // Fallback: reconstruct from r,g,b
                    let r = o.properties.get("r").map(|v| v.as_f64() as u8).unwrap_or(0);
                    let g = o.properties.get("g").map(|v| v.as_f64() as u8).unwrap_or(0);
                    let b = o.properties.get("b").map(|v| v.as_f64() as u8).unwrap_or(0);
                    return Some(format!("#{:02X}{:02X}{:02X}", r, g, b));
                }
                if t.as_ref() == "BorderStyle" {
                    if let Some(Value::String(name)) = o.properties.get("name") {
                        return Some(name.to_string());
                    }
                }
            }
            // Skip complex objects (Point, Size, functions, etc.)
            None
        }
        Value::Null => None,
        _ => None,
    }
}

fn apply_property(form: &mut vybe_widgets::Form, control_name: &str, property: &str, value: &str) {
    use vybe_widgets::{WidgetCommand, CommandValue};
    match property {
        "Text" | "text" => {
            form.send_command(control_name, &WidgetCommand::SetText(value.to_string()));
        }
        "Enabled" | "enabled" => {
            let enabled = !matches!(value, "false" | "False" | "0" | "");
            form.send_command(control_name, &WidgetCommand::SetEnabled(enabled));
        }
        "Visible" | "visible" => {
            let visible = !matches!(value, "false" | "False" | "0" | "");
            form.send_command(control_name, &WidgetCommand::SetVisible(visible));
        }
        "ReadOnly" | "readonly" => {
            let ro = matches!(value, "true" | "True" | "1");
            form.send_command(control_name, &WidgetCommand::Custom("SetReadOnly".into(), CommandValue::Bool(ro)));
        }
        _ => {
            form.send_command(control_name, &WidgetCommand::Custom(
                format!("Set{}", capitalize_first(property)),
                CommandValue::Text(value.to_string()),
            ));
        }
    }
}

} // mod gui_impl

// Public re-export when gui feature is on
#[cfg(feature = "gui")]
pub use gui_impl::register;

// Non-GUI fallback: register stubs so compiled code does not crash.
#[cfg(not(feature = "gui"))]
pub fn register(
    vm: &mut vybe_bytecode::VM,
) {
    use vybe_bytecode::Value;
    let stubs = [
        "createForm", "addControl", "setProperty", "getProperty",
        "onEvent", "showForm", "runApplication", "msgBox", "closeForm",
        "newControl", "controlSetProperty", "controlsAdd", "newForm",
        "noop", "addHandler", "removeHandler",
        "__ctrl_show", "__ctrl_close", "__ctrl_focus", "__ctrl_hide",
        "__dlg_showdialog", "__dlg_show",
    ];
    for name in stubs {
        vm.register_host_fn("vybe:gui", name, Box::new(|_ctx, _| Value::Null));
    }
}
