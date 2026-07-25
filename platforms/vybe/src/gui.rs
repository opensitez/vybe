//! `vybe:gui` host module — registers GUI host functions on the VM.
//!
//! When the `gui` feature is enabled, host functions directly create
//! `vybe_widgets` widgets and store them in a shared `GuiState`.

#[cfg(feature = "gui")]
mod gui_impl {

    use crate::gui_state::GuiState;
    use std::sync::{Arc, Mutex};
    use vybe_bytecode::value::{Object, ObjectKind};
    use vybe_bytecode::{HostContext, VM, Value};

    fn gui_trace_enabled() -> bool {
        std::env::var("VYBE_GUI_TRACE")
            .map(|value| !matches!(value.as_str(), "" | "0" | "false" | "False"))
            .unwrap_or(false)
    }

    fn control_property_value_from_live_or_fallback(
        gui: &mut GuiState,
        control_name: &str,
        property: &str,
        fallback: Option<Value>,
    ) -> Value {
        let prop_lower = property.to_lowercase();

        if !control_name.is_empty() {
            let value = gui.get_property(control_name, property);
            if !value.is_empty() || prop_lower == "text" {
                return match prop_lower.as_str() {
                    "enabled" | "visible" | "readonly" | "tabstop" | "autosize" | "multiline" => {
                        Value::Bool(matches!(value.as_str(), "true" | "True" | "1"))
                    }
                    "left" | "top" | "width" | "height" | "tabindex" | "maxlength"
                    | "selectionlength" | "selectionstart" | "textlength" | "opacity" => value
                        .parse::<f64>()
                        .map(Value::F64)
                        .unwrap_or_else(|_| Value::String(Arc::from(value.as_str()))),
                    _ => Value::String(Arc::from(value.as_str())),
                };
            }
        }

        fallback.unwrap_or(Value::Null)
    }

    fn collection_type_name(kind: &str) -> &'static str {
        match kind {
            "controls" => "ControlCollection",
            "components" => "ComponentCollection",
            "forms" => "FormCollection",
            _ => "Collection",
        }
    }

    fn sync_collection_metadata(collection: &mut Object) {
        let len = match &collection.kind {
            ObjectKind::Array(items) => items.len(),
            _ => 0,
        };
        collection
            .properties
            .insert("count".into(), Value::F64(len as f64));
        collection
            .properties
            .insert("length".into(), Value::F64(len as f64));
    }

    fn create_collection_object(
        kind: &str,
        owner: Option<Value>,
        add_ref: &Value,
        clear_ref: &Value,
        contains_ref: &Value,
    ) -> Value {
        let mut collection = Object::new_array(Vec::new());
        collection.properties.insert(
            "__type".into(),
            Value::String(Arc::from(collection_type_name(kind))),
        );
        collection
            .properties
            .insert("__collection_kind".into(), Value::String(Arc::from(kind)));
        if let Some(owner) = owner {
            collection.properties.insert("__owner".into(), owner);
        }
        collection.properties.insert("add".into(), add_ref.clone());
        collection
            .properties
            .insert("clear".into(), clear_ref.clone());
        collection
            .properties
            .insert("contains".into(), contains_ref.clone());
        sync_collection_metadata(&mut collection);
        Value::Object(vybe_bytecode::heap::alloc(collection))
    }

    fn push_collection_value(collection_obj: &Arc<Mutex<Object>>, value: Value) {
        let mut collection = collection_obj.lock().unwrap();
        if let ObjectKind::Array(items) = &mut collection.kind {
            if !items.iter().any(|existing| existing.eq(&value)) {
                items.push(value);
            }
        }
        sync_collection_metadata(&mut collection);
    }

    fn clear_collection_value(collection_obj: &Arc<Mutex<Object>>) {
        let mut collection = collection_obj.lock().unwrap();
        if let ObjectKind::Array(items) = &mut collection.kind {
            items.clear();
        }
        sync_collection_metadata(&mut collection);
    }

    fn remove_collection_value(collection_obj: &Arc<Mutex<Object>>, value: &Value) {
        let mut collection = collection_obj.lock().unwrap();
        if let ObjectKind::Array(items) = &mut collection.kind {
            items.retain(|existing| !existing.eq(value));
        }
        sync_collection_metadata(&mut collection);
    }

    fn append_to_owner_collection(owner_obj: &Arc<Mutex<Object>>, property: &str, value: Value) {
        let collection = { owner_obj.lock().unwrap().properties.get(property).cloned() };
        if let Some(Value::Object(collection_obj)) = collection {
            push_collection_value(&collection_obj, value);
        }
    }

    fn clear_owner_collection(owner_obj: &Arc<Mutex<Object>>, property: &str) {
        let collection = { owner_obj.lock().unwrap().properties.get(property).cloned() };
        if let Some(Value::Object(collection_obj)) = collection {
            clear_collection_value(&collection_obj);
        }
    }

    fn is_non_visual_component_type(type_name: &str) -> bool {
        matches!(
            type_name.to_lowercase().as_str(),
            "bindingsource"
                | "timer"
                | "imagelist"
                | "tooltip"
                | "notifyicon"
                | "errorprovider"
                | "helpprovider"
                | "backgroundworker"
                | "dataset"
                | "datatable"
                | "dataadapter"
                | "openfiledialog"
                | "savefiledialog"
                | "folderbrowserdialog"
                | "fontdialog"
                | "colordialog"
                | "printdialog"
                | "printpreviewdialog"
                | "pagesetupdialog"
                | "printdocument"
                | "sqlconnection"
                | "dataview"
        )
    }

    fn refresh_form_component_collection(form_obj: &Arc<Mutex<Object>>) {
        clear_owner_collection(form_obj, "components");
        let components: Vec<Value> = {
            let form = form_obj.lock().unwrap();
            form.properties
                .iter()
                .filter_map(|(key, value)| {
                    if key.starts_with("__") || matches!(key.as_str(), "controls" | "components") {
                        return None;
                    }
                    let Value::Object(candidate) = value else {
                        return None;
                    };
                    let type_name = {
                        let candidate_guard = candidate.lock().unwrap();
                        candidate_guard
                            .properties
                            .get("__type")
                            .or_else(|| candidate_guard.properties.get("__control_type"))
                            .map(|value| format!("{}", value))
                            .unwrap_or_default()
                    };
                    if is_non_visual_component_type(&type_name) {
                        Some(Value::Object(candidate.clone()))
                    } else {
                        None
                    }
                })
                .collect()
        };

        for component in components {
            append_to_owner_collection(form_obj, "components", component);
        }
    }

    fn controls_add_impl(
        gui: &Arc<Mutex<GuiState>>,
        parent: Option<&Value>,
        obj: &Arc<Mutex<Object>>,
    ) {
        let (parent_abs_x, parent_abs_y) = if let Some(Value::Object(parent_obj)) = parent {
            let po = parent_obj.lock().unwrap();
            let (px, py) = if let Some(Value::Object(loc)) = po.properties.get("location") {
                let loc = loc.lock().unwrap();
                (
                    loc.properties
                        .get("x")
                        .map(|v| v.as_f64() as i32)
                        .unwrap_or(0),
                    loc.properties
                        .get("y")
                        .map(|v| v.as_f64() as i32)
                        .unwrap_or(0),
                )
            } else {
                (
                    po.properties
                        .get("left")
                        .map(|v| v.as_f64() as i32)
                        .unwrap_or(0),
                    po.properties
                        .get("top")
                        .map(|v| v.as_f64() as i32)
                        .unwrap_or(0),
                )
            };
            let mut abs_x = px;
            let mut abs_y = py;
            let mut cur = po.properties.get("__parent").cloned();
            drop(po);
            while let Some(Value::Object(ancestor)) = cur {
                let anc = ancestor.lock().unwrap();
                let (ax, ay) = if let Some(Value::Object(loc)) = anc.properties.get("location") {
                    let loc = loc.lock().unwrap();
                    (
                        loc.properties
                            .get("x")
                            .map(|v| v.as_f64() as i32)
                            .unwrap_or(0),
                        loc.properties
                            .get("y")
                            .map(|v| v.as_f64() as i32)
                            .unwrap_or(0),
                    )
                } else {
                    (
                        anc.properties
                            .get("left")
                            .map(|v| v.as_f64() as i32)
                            .unwrap_or(0),
                        anc.properties
                            .get("top")
                            .map(|v| v.as_f64() as i32)
                            .unwrap_or(0),
                    )
                };
                abs_x += ax;
                abs_y += ay;
                cur = anc.properties.get("__parent").cloned();
            }
            (abs_x, abs_y)
        } else {
            (0, 0)
        };

        let parent_field_name = if let Some(Value::Object(parent_obj)) = parent {
            let parent = parent_obj.lock().unwrap();
            parent.properties.iter().find_map(|(key, value)| {
                if key.starts_with("__") {
                    return None;
                }
                match value {
                    Value::Object(candidate) if Arc::ptr_eq(candidate, obj) => Some(key.clone()),
                    _ => None,
                }
            })
        } else {
            None
        };

        if let Some(parent_val) = parent {
            obj.lock()
                .unwrap()
                .properties
                .insert("__parent".into(), parent_val.clone());
        }
        if let Some(ref field_name) = parent_field_name {
            let mut child = obj.lock().unwrap();
            let explicit_name = child
                .properties
                .get("name")
                .map(|v| format!("{}", v))
                .filter(|name| !name.is_empty());
            let effective_name = explicit_name.unwrap_or_else(|| field_name.clone());
            child.properties.insert(
                "name".into(),
                Value::String(Arc::from(effective_name.as_str())),
            );
            child.properties.insert(
                "__control_name".into(),
                Value::String(Arc::from(effective_name.as_str())),
            );
        }
        let o = obj.lock().unwrap();
        let control_type = o
            .properties
            .get("__control_type")
            .map(|v| format!("{}", v))
            .unwrap_or_else(|| "Button".into());
        let control_name = o
            .properties
            .get("name")
            .or_else(|| o.properties.get("__control_name"))
            .map(|v| format!("{}", v))
            .unwrap_or_else(|| "ctrl".into());
        let text = o
            .properties
            .get("text")
            .map(|v| format!("{}", v))
            .unwrap_or_default();
        let left = o
            .properties
            .get("left")
            .map(|v| v.as_f64() as i32)
            .unwrap_or(0);
        let top = o
            .properties
            .get("top")
            .map(|v| v.as_f64() as i32)
            .unwrap_or(0);
        let width = o
            .properties
            .get("width")
            .map(|v| v.as_f64() as i32)
            .unwrap_or(100);
        let height = o
            .properties
            .get("height")
            .map(|v| v.as_f64() as i32)
            .unwrap_or(30);
        let (left, top) = if let Some(Value::Object(loc)) = o.properties.get("location") {
            let loc = loc.lock().unwrap();
            (
                loc.properties
                    .get("x")
                    .map(|v| v.as_f64() as i32)
                    .unwrap_or(left),
                loc.properties
                    .get("y")
                    .map(|v| v.as_f64() as i32)
                    .unwrap_or(top),
            )
        } else {
            (left, top)
        };
        let (width, height) = if let Some(Value::Object(sz)) = o.properties.get("size") {
            let sz = sz.lock().unwrap();
            (
                sz.properties
                    .get("width")
                    .map(|v| v.as_f64() as i32)
                    .unwrap_or(width),
                sz.properties
                    .get("height")
                    .map(|v| v.as_f64() as i32)
                    .unwrap_or(height),
            )
        } else {
            (width, height)
        };
        let props: Vec<(String, String)> = o
            .properties
            .iter()
            .filter(|(k, _)| {
                !k.starts_with("__")
                    && !matches!(
                        k.as_str(),
                        "name"
                            | "left"
                            | "top"
                            | "width"
                            | "height"
                            | "text"
                            | "location"
                            | "size"
                            | "show"
                            | "close"
                            | "focus"
                            | "hide"
                            | "showdialog"
                            | "controls"
                            | "components"
                    )
            })
            .filter_map(|(k, v)| {
                let val_str = value_to_property_string(v)?;
                Some((capitalize_first(k), val_str))
            })
            .collect();
        drop(o);

        // Declarative (Flutter) path: when the parent is a layout container
        // (FlowLayoutPanel/StackPanel), or the `runApp` root (a layout panel
        // added under the form — `createForm` returns the form NAME string, so
        // a string/none parent IS the form), stage into the widget tree and let
        // vybe_widgets own nesting + flow layout. WinForms/VCL adapters use
        // absolute Panels/Forms with a Form *object* parent and non-layout
        // children, so they fall through to the flat placement below.
        let (parent_is_layout, parent_name) = match parent {
            Some(Value::Object(p)) => {
                let p = p.lock().unwrap();
                let pt = p
                    .properties
                    .get("__control_type")
                    .map(|v| format!("{}", v))
                    .unwrap_or_default();
                let pn = p
                    .properties
                    .get("__control_name")
                    .or_else(|| p.properties.get("name"))
                    .map(|v| format!("{}", v))
                    .unwrap_or_default();
                (matches!(pt.as_str(), "FlowLayoutPanel" | "HFlowLayoutPanel" | "StackPanel"), pn)
            }
            _ => (false, String::new()),
        };
        let parent_is_form = matches!(parent, Some(Value::String(_)) | None);
        let child_is_layout = matches!(control_type.as_str(), "FlowLayoutPanel" | "HFlowLayoutPanel" | "StackPanel");
        if parent_is_layout || (parent_is_form && child_is_layout) {
            {
                let mut g = gui.lock().unwrap();
                g.stage_control(
                    &control_type,
                    &control_name,
                    &text,
                    width,
                    height,
                    &parent_name,
                    parent_is_form && !parent_is_layout,
                );
            }
            if let Some(Value::Object(parent_obj)) = parent {
                append_to_owner_collection(parent_obj, "controls", Value::Object(obj.clone()));
            }
            return;
        }

        let abs_left = left + parent_abs_x;
        let abs_top = top + parent_abs_y;
        let mut g = gui.lock().unwrap();
        g.add_widget(
            &control_type,
            &control_name,
            &text,
            abs_left,
            abs_top,
            width,
            height,
        );
        if gui_trace_enabled() {
            eprintln!(
                "[gui-host] controlsAdd type={} parent_field={:?} widget_name={} text={}",
                control_type, parent_field_name, control_name, text,
            );
        }
        for (prop, val) in props {
            apply_property(&mut g.form, &control_name, &prop, &val);
        }
        drop(g);

        if let Some(Value::Object(parent_obj)) = parent {
            append_to_owner_collection(parent_obj, "controls", Value::Object(obj.clone()));
        }
    }

    pub fn register(vm: &mut VM, gui: Arc<Mutex<GuiState>>) {
        let gui_collection_add = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            "__collection_add",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let Some(Value::Object(collection_obj)) = args.first() else {
                    return Value::Null;
                };
                let value = args.get(1).cloned().unwrap_or(Value::Null);
                let (kind, owner) = {
                    let collection = collection_obj.lock().unwrap();
                    let kind = collection
                        .properties
                        .get("__collection_kind")
                        .map(|value| format!("{}", value))
                        .unwrap_or_default();
                    let owner = collection.properties.get("__owner").cloned();
                    (kind, owner)
                };

                match kind.as_str() {
                    "controls" => {
                        if let (Some(owner), Value::Object(child_obj)) = (owner.as_ref(), &value) {
                            controls_add_impl(&gui_collection_add, Some(owner), child_obj);
                        }
                    }
                    "components" | "forms" => {
                        push_collection_value(collection_obj, value);
                    }
                    _ => {}
                }
                Value::Null
            }),
        );
        vm.register_host_fn(
            "vybe:gui",
            "__collection_clear",
            Box::new(|_ctx: &mut HostContext, args: &[Value]| {
                if let Some(Value::Object(collection_obj)) = args.first() {
                    clear_collection_value(collection_obj);
                }
                Value::Null
            }),
        );
        vm.register_host_fn(
            "vybe:gui",
            "__collection_contains",
            Box::new(|_ctx: &mut HostContext, args: &[Value]| {
                let Some(Value::Object(collection_obj)) = args.first() else {
                    return Value::Bool(false);
                };
                let needle = args.get(1).cloned().unwrap_or(Value::Null);
                let contains = {
                    let collection = collection_obj.lock().unwrap();
                    if let ObjectKind::Array(items) = &collection.kind {
                        items.iter().any(|existing| existing.eq(&needle))
                    } else {
                        false
                    }
                };
                Value::Bool(contains)
            }),
        );
        let collection_add_ref = host_fn_ref(vm, "__collection_add");
        let collection_clear_ref = host_fn_ref(vm, "__collection_clear");
        let collection_contains_ref = host_fn_ref(vm, "__collection_contains");
        let open_forms = create_collection_object(
            "forms",
            None,
            &collection_add_ref,
            &collection_clear_ref,
            &collection_contains_ref,
        );
        vm.globals.insert("__openforms".into(), open_forms.clone());

        vm.register_host_fn("vybe:gui", "newControlsCollection", {
            let add_ref = collection_add_ref.clone();
            let clear_ref = collection_clear_ref.clone();
            let contains_ref = collection_contains_ref.clone();
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                create_collection_object(
                    "controls",
                    args.first().cloned(),
                    &add_ref,
                    &clear_ref,
                    &contains_ref,
                )
            })
        });
        vm.register_host_fn("vybe:gui", "newComponentsCollection", {
            let add_ref = collection_add_ref.clone();
            let clear_ref = collection_clear_ref.clone();
            let contains_ref = collection_contains_ref.clone();
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                create_collection_object(
                    "components",
                    args.first().cloned(),
                    &add_ref,
                    &clear_ref,
                    &contains_ref,
                )
            })
        });

        // Form creation
        vm.register_host_fn("vybe:gui", "createForm", {
            let gui = gui.clone();
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let title = str_arg(args, 0, "Form1");
                let name = title.clone();
                let mut g = gui.lock().unwrap();
                g.form = vybe_widgets::Form::new(&title);
                g.seed_form_identity(&name, &title);
                Value::String(Arc::from(name.as_str()))
            })
        });

        vm.register_host_fn("vybe:gui", "newForm", {
            let gui = gui.clone();
            let add_ref = collection_add_ref.clone();
            let clear_ref = collection_clear_ref.clone();
            let contains_ref = collection_contains_ref.clone();
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let title = str_arg(args, 0, "Form1");
                let name = title.clone();
                {
                    let mut g = gui.lock().unwrap();
                    g.form = vybe_widgets::Form::new(&title);
                    g.seed_form_identity(&name, &title);
                }
                let form_obj = vybe_bytecode::heap::alloc(Object::new());
                {
                    let mut obj = form_obj.lock().unwrap();
                    obj.properties
                        .insert("__control_type".into(), Value::String(Arc::from("Form")));
                    obj.properties.insert(
                        "__control_name".into(),
                        Value::String(Arc::from(name.as_str())),
                    );
                    obj.properties
                        .insert("name".into(), Value::String(Arc::from(name.as_str())));
                    obj.properties
                        .insert("text".into(), Value::String(Arc::from(title.as_str())));
                    obj.properties.insert("width".into(), Value::F64(800.0));
                    obj.properties.insert("height".into(), Value::F64(600.0));
                }
                let owner = Value::Object(form_obj.clone());
                let controls = create_collection_object(
                    "controls",
                    Some(owner.clone()),
                    &add_ref,
                    &clear_ref,
                    &contains_ref,
                );
                let components = create_collection_object(
                    "components",
                    Some(owner.clone()),
                    &add_ref,
                    &clear_ref,
                    &contains_ref,
                );
                {
                    let mut obj = form_obj.lock().unwrap();
                    obj.properties.insert("controls".into(), controls);
                    obj.properties.insert("components".into(), components);
                }
                owner
            })
        });

        // Add control to form
        vm.register_host_fn("vybe:gui", "controlsAdd", {
            let gui = gui.clone();
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                if let Some(Value::Object(obj)) = args.get(1) {
                    controls_add_impl(&gui, args.first(), obj);
                }
                Value::Null
            })
        });

        // Drop all controls so the Flutter `setState` rebuild can re-realize
        // the widget tree from scratch (State persists in the Dart runtime).
        vm.register_host_fn("vybe:gui", "clearControls", {
            let gui = gui.clone();
            Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
                gui.lock().unwrap().form.clear_controls();
                Value::Null
            })
        });

        // True when a control with this name already exists — lets the Flutter
        // realizer create-or-update by stable name (the control's name is its
        // cross-framework identity; setState updates in place, no rebuild).
        vm.register_host_fn("vybe:gui", "hasControl", {
            let gui = gui.clone();
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let name = str_arg(args, 0, "");
                Value::Bool(gui.lock().unwrap().is_live_control_name(&name))
            })
        });

        // Make a control-handle object of `type` named `name` (NOT added to any
        // form yet — the realizer nests it with `controlsAdd`). `type` is a
        // control type string (`Label`/`Button`/`FlowLayoutPanel`…).
        vm.register_host_fn(
            "vybe:gui",
            "newControl",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let type_name = str_arg(args, 0, "Panel");
                let name = str_arg(args, 1, "control");
                let obj = vybe_bytecode::heap::alloc(Object::new());
                {
                    let mut o = obj.lock().unwrap();
                    let s = |v: &str| Value::String(Arc::from(v));
                    o.properties.insert("__control_type".into(), s(&type_name));
                    o.properties.insert("__control_name".into(), s(&name));
                    o.properties.insert("__type".into(), s(&type_name));
                    o.properties.insert("name".into(), s(&name));
                    o.properties.insert("width".into(), Value::F64(100.0));
                    o.properties.insert("height".into(), Value::F64(30.0));
                    o.properties.insert("left".into(), Value::F64(0.0));
                    o.properties.insert("top".into(), Value::F64(0.0));
                }
                Value::Object(obj)
            }),
        );

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
                gui.lock().unwrap().add_widget(
                    &control_type,
                    &control_name,
                    "",
                    left,
                    top,
                    width,
                    height,
                );
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
                gui.lock()
                    .unwrap()
                    .set_property(&control, &property, &val_str);
                Value::Null
            })
        });

        vm.register_host_fn("vybe:gui", "controlSetProperty", {
        let gui = gui.clone();
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let property = str_arg(args, 1, "");
                let val = args.get(2).cloned().unwrap_or(Value::Null);
                let val_str = format!("{}", val);
                let prop_lower = property.to_lowercase();
                let (control_name, fallback_text, control_type) = {
                    let o = obj.lock().unwrap();
                    let text = o.properties.get("text").map(|v| format!("{}", v));
                    let control_type = o.properties
                        .get("__control_type")
                        .map(|v| format!("{}", v))
                        .unwrap_or_default();
                    let control_name = o.properties.get("__control_name")
                        .or_else(|| o.properties.get("name"))
                        .map(|v| format!("{}", v)).unwrap_or_default();
                    (control_name, text, control_type)
                };
                let live_widget = if control_name.is_empty() {
                    false
                } else {
                    gui.lock().unwrap().is_live_control_name(&control_name)
                };
                if !live_widget || prop_lower == "name" {
                    obj.lock().unwrap().properties.insert(prop_lower.clone(), val.clone());
                }
                if prop_lower == "name" {
                    {
                        let mut g = gui.lock().unwrap();
                        g.rename_control(&control_name, &val_str);
                        if let Some(text) = fallback_text.as_deref() {
                            if !text.is_empty() && g.get_property(&val_str, "Text").is_empty() {
                                g.set_property(&val_str, "Text", text);
                            }
                        } else if control_type.eq_ignore_ascii_case("Form")
                            && g.get_property(&val_str, "Text").is_empty()
                        {
                            g.set_property(&val_str, "Text", &val_str);
                        }
                    }
                    obj.lock().unwrap().properties.insert("__control_name".into(), val.clone());
                }
                if gui_trace_enabled() && matches!(prop_lower.as_str(), "name" | "text") {
                    eprintln!(
                        "[gui-host] controlSetProperty control={} property={} live_widget={} value={}",
                        control_name,
                        property,
                        live_widget,
                        val_str,
                    );
                }
                gui.lock().unwrap().set_property(&control_name, &property, &val_str);
            }
            Value::Null
        })
    });

        vm.register_host_fn("vybe:gui", "controlGetProperty", {
            let gui = gui.clone();
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let Some(Value::Object(obj)) = args.first() else {
                    return Value::Null;
                };

                let property = str_arg(args, 1, "");
                let prop_lower = property.to_lowercase();
                let (control_name, fallback) = {
                    let o = obj.lock().unwrap();
                    let fallback = o.properties.get(&prop_lower).cloned();
                    let control_name = o
                        .properties
                        .get("__control_name")
                        .or_else(|| o.properties.get("name"))
                        .map(|v| format!("{}", v))
                        .unwrap_or_default();
                    (control_name, fallback)
                };

                control_property_value_from_live_or_fallback(
                    &mut gui.lock().unwrap(),
                    &control_name,
                    &property,
                    fallback,
                )
            })
        });

        vm.register_host_fn("vybe:gui", "getProperty", {
            let gui = gui.clone();
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let control = str_arg(args, 0, "");
                let property = str_arg(args, 1, "");
                let val = gui.lock().unwrap().get_property(&control, &property);
                if val.is_empty() {
                    Value::Null
                } else {
                    Value::String(Arc::from(val.as_str()))
                }
            })
        });

        // Event registration
        vm.register_host_fn("vybe:gui", "onEvent", {
            let gui = gui.clone();
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let control = str_arg(args, 0, "");
                let event = str_arg(args, 1, "");
                let callback = args.get(2).cloned().unwrap_or(Value::Null);
                gui.lock()
                    .unwrap()
                    .register_event(&control, &event, callback);
                Value::Null
            })
        });

        vm.register_host_fn("vybe:gui", "addHandler", {
            let gui = gui.clone();
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let ctrl = str_arg(args, 0, "");
                let event = str_arg(args, 1, "");
                if let Some(callback) = args.get(2) {
                    gui.lock()
                        .unwrap()
                        .register_event(&ctrl, &event, callback.clone());
                }
                Value::Null
            })
        });

        vm.register_host_fn("vybe:gui", "removeHandler", Box::new(|_ctx, _| Value::Null));

        // Form lifecycle
        vm.register_host_fn("vybe:gui", "showForm", {
            let gui = gui.clone();
            let open_forms = open_forms.clone();
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let mut g = gui.lock().unwrap();
                g.should_run = true;
                if let Some(obj) = args.first().cloned() {
                    if let Value::Object(form_obj) = &obj {
                        refresh_form_component_collection(form_obj);
                    }
                    if let Value::Object(open_forms_obj) = &open_forms {
                        push_collection_value(open_forms_obj, obj.clone());
                    }
                    g.form_object = Some(obj);
                }
                Value::Null
            })
        });

        vm.register_host_fn("vybe:gui", "runApplication", {
            let gui = gui.clone();
            let open_forms = open_forms.clone();
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let mut g = gui.lock().unwrap();
                g.should_run = true;
                if let Some(obj) = args.first().cloned() {
                    if let Value::Object(o) = &obj {
                        let o = o.lock().unwrap();
                        if let Some(w) = o.properties.get("width") {
                            g.width = w.as_f64() as u32;
                        }
                        if let Some(h) = o.properties.get("height") {
                            g.height = h.as_f64() as u32;
                        }
                    }
                    if let Value::Object(form_obj) = &obj {
                        refresh_form_component_collection(form_obj);
                    }
                    if let Value::Object(open_forms_obj) = &open_forms {
                        push_collection_value(open_forms_obj, obj.clone());
                    }
                    g.form_object = Some(obj);
                }
                Value::Null
            })
        });

        vm.register_host_fn("vybe:gui", "closeForm", {
            let gui = gui.clone();
            let open_forms = open_forms.clone();
            Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
                let mut guard = gui.lock().unwrap();
                guard.close_requested = true;
                if let (Some(form_obj), Value::Object(open_forms_obj)) =
                    (guard.form_object.clone(), &open_forms)
                {
                    remove_collection_value(open_forms_obj, &form_obj);
                }
                Value::Null
            })
        });

        vm.register_host_fn("vybe:gui", "appExit", {
            let gui = gui.clone();
            let open_forms = open_forms.clone();
            Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
                let mut guard = gui.lock().unwrap();
                guard.close_requested = true;
                if let (Some(form_obj), Value::Object(open_forms_obj)) =
                    (guard.form_object.clone(), &open_forms)
                {
                    remove_collection_value(open_forms_obj, &form_obj);
                }
                Value::Null
            })
        });

        // MsgBox — show a native message dialog inline via vybe_widgets.
        //
        // Blocks the calling thread (and therefore the VM) until the user
        // dismisses the dialog. The previous implementation queued the
        // request on `GuiState::pending_dialogs` for the form runner to
        // drain — that queue is gone now: the host fn calls
        // `vybe_widgets::dialogs::MessageBox` directly, the OS handles
        // modality, and the VM resumes when the user clicks OK. Same
        // semantics as a native blocking call.
        vm.register_host_fn(
            "vybe:gui",
            "msgBox",
            Box::new(|_ctx: &mut HostContext, args: &[Value]| {
                let text = str_arg(args, 0, "");
                let title = str_arg(args, 1, "Message");
                vybe_widgets::dialogs::MessageBox::info(text, title);
                Value::Null
            }),
        );

        vm.register_host_fn("vybe:gui", "noop", Box::new(|_ctx, _| Value::Null));

        // Control constructors
        let control_types = [
            "Button",
            "Label",
            "TextBox",
            "CheckBox",
            "RadioButton",
            "ComboBox",
            "ListBox",
            "Panel",
            "GroupBox",
            "TabControl",
            "TabPage",
            "DataGridView",
            "ProgressBar",
            "TrackBar",
            "NumericUpDown",
            "DateTimePicker",
            "RichTextBox",
            "PictureBox",
            "MenuStrip",
            "ToolStrip",
            "StatusStrip",
            "SplitContainer",
            "FlowLayoutPanel",
            "HFlowLayoutPanel",
            "TableLayoutPanel",
            "LinkLabel",
            "MaskedTextBox",
            "HScrollBar",
            "VScrollBar",
            "MonthCalendar",
            "BindingNavigator",
            "BindingSource",
            "DataSet",
            "DataTable",
            "DataAdapter",
            "OpenFileDialog",
            "SaveFileDialog",
            "FontDialog",
            "ColorDialog",
            "FolderBrowserDialog",
            "PrintDialog",
            "PrintPreviewDialog",
            "ListView",
            "WebBrowser",
            "ContextMenuStrip",
            "Timer",
            "ImageList",
            "ToolTip",
            "NotifyIcon",
            "ErrorProvider",
            "HelpProvider",
            "BackgroundWorker",
            "Form",
            "TreeView",
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
        vm.register_host_fn(
            "vybe:gui",
            "__ctrl_show",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let (name, is_form) = read_this_identity(args);
                let mut g = gui_show.lock().unwrap();
                if is_form {
                    g.should_run = true;
                }
                if !name.is_empty() {
                    g.set_property(&name, "Visible", "true");
                }
                Value::Null
            }),
        );
        let gui_close = gui.clone();
        let open_forms_close = open_forms.clone();
        vm.register_host_fn(
            "vybe:gui",
            "__ctrl_close",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let (_name, is_form) = read_this_identity(args);
                let mut g = gui_close.lock().unwrap();
                if is_form {
                    g.close_requested = true;
                    if let (Some(form_obj), Value::Object(open_forms_obj)) =
                        (args.first().cloned(), &open_forms_close)
                    {
                        remove_collection_value(open_forms_obj, &form_obj);
                    }
                }
                // Non-form `Close` (rare — most controls don't expose Close in real
                // .NET; only forms do) is a no-op.
                Value::Null
            }),
        );
        let gui_hide = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            "__ctrl_hide",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let (name, _is_form) = read_this_identity(args);
                if !name.is_empty() {
                    gui_hide
                        .lock()
                        .unwrap()
                        .set_property(&name, "Visible", "false");
                }
                Value::Null
            }),
        );
        let gui_focus = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            "__ctrl_focus",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let (name, _) = read_this_identity(args);
                if !name.is_empty() {
                    let mut g = gui_focus.lock().unwrap();
                    g.form
                        .send_command(&name, &vybe_widgets::WidgetCommand::Focus);
                }
                Value::Null
            }),
        );
        // Refresh / Invalidate / Update all map to "request a repaint". Real
        // .NET distinguishes them (Refresh = Invalidate + Update), but our
        // renderer is fully repainting every frame, so the distinction
        // collapses to a single needs_repaint flag.
        let gui_refresh = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            "__ctrl_refresh",
            Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
                gui_refresh.lock().unwrap().needs_repaint = true;
                Value::Null
            }),
        );
        let gui_invalidate = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            "__ctrl_invalidate",
            Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
                gui_invalidate.lock().unwrap().needs_repaint = true;
                Value::Null
            }),
        );
        let gui_update = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            "__ctrl_update",
            Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
                gui_update.lock().unwrap().needs_repaint = true;
                Value::Null
            }),
        );
        // BringToFront / SendToBack — relayed to the widget via Custom command.
        // Form delegates to the OS window manager (out of scope for the
        // headless renderer); child controls reorder within the form. We
        // capture the intent on the property store too so tests can verify
        // the call landed.
        let gui_btf = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            "__ctrl_bring_to_front",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let (name, _) = read_this_identity(args);
                if !name.is_empty() {
                    let mut g = gui_btf.lock().unwrap();
                    g.form.send_command(
                        &name,
                        &vybe_widgets::WidgetCommand::Custom(
                            "BringToFront".into(),
                            vybe_widgets::CommandValue::None,
                        ),
                    );
                    g.properties
                        .insert((name, "__zorder".into()), "front".into());
                    g.needs_repaint = true;
                }
                Value::Null
            }),
        );
        let gui_stb = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            "__ctrl_send_to_back",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let (name, _) = read_this_identity(args);
                if !name.is_empty() {
                    let mut g = gui_stb.lock().unwrap();
                    g.form.send_command(
                        &name,
                        &vybe_widgets::WidgetCommand::Custom(
                            "SendToBack".into(),
                            vybe_widgets::CommandValue::None,
                        ),
                    );
                    g.properties
                        .insert((name, "__zorder".into()), "back".into());
                    g.needs_repaint = true;
                }
                Value::Null
            }),
        );
        // Dispose: hide the control + drop its event handlers + clear any
        // pending draw commands targeting it. Real .NET also frees native
        // GDI handles, which we don't have.
        let gui_dispose = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            "__ctrl_dispose",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
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
            }),
        );
        // Form.Activate — request the OS window be brought to the foreground.
        let gui_activate = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            "__form_activate",
            Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
                gui_activate.lock().unwrap().front_requested = true;
                Value::Null
            }),
        );
        // Form.CenterToScreen — compute the centered position from screen size
        // and set it on the form. Stored on the property store; the window
        // driver reads it on next show.
        let gui_center = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            "__form_center_to_screen",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
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
            }),
        );
        // ShowDialog — modal show. Sets should_run + close_requested
        // semantics is up to the runner; the return value is `DialogResult.OK = 1`
        // by convention, the same as the legacy stub. Real modal handling is a
        // separate workstream (it requires nested message loops).
        let gui_show_dlg = gui.clone();
        vm.register_host_fn(
            "vybe:gui",
            "__dlg_showdialog",
            Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
                gui_show_dlg.lock().unwrap().should_run = true;
                Value::I32(1) // DialogResult.OK
            }),
        );
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
            let add_ref = collection_add_ref.clone();
            let clear_ref = collection_clear_ref.clone();
            let contains_ref = collection_contains_ref.clone();
            vm.register_host_fn(
                "vybe:gui",
                &format!("new_{}", ct),
                Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
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
                    let obj = vybe_bytecode::heap::alloc(vybe_bytecode::value::Object::new());
                    {
                        let mut object = obj.lock().unwrap();
                        object.properties.insert(
                            "__control_type".into(),
                            Value::String(Arc::from(type_name.as_str())),
                        );
                        object.properties.insert(
                            "__control_name".into(),
                            Value::String(Arc::from(name.as_str())),
                        );
                        object.properties.insert(
                            "__type".into(),
                            Value::String(Arc::from(type_name.as_str())),
                        );
                        object
                            .properties
                            .insert("name".into(), Value::String(Arc::from(name.as_str())));
                        object.properties.insert("width".into(), Value::F64(100.0));
                        object.properties.insert("height".into(), Value::F64(30.0));
                        object.properties.insert("left".into(), Value::F64(0.0));
                        object.properties.insert("top".into(), Value::F64(0.0));
                        object.properties.insert("show".into(), show.clone());
                        object.properties.insert("close".into(), close.clone());
                        object.properties.insert("focus".into(), focus.clone());
                        object.properties.insert("hide".into(), hide.clone());
                        if matches!(
                            type_name.as_str(),
                            "OpenFileDialog"
                                | "SaveFileDialog"
                                | "FontDialog"
                                | "ColorDialog"
                                | "FolderBrowserDialog"
                                | "PrintDialog"
                                | "PrintPreviewDialog"
                        ) {
                            object.properties.insert("showdialog".into(), dlg.clone());
                        }
                    }
                    let owner = Value::Object(obj.clone());
                    if !is_non_visual_component_type(&type_name) {
                        let controls = create_collection_object(
                            "controls",
                            Some(owner.clone()),
                            &add_ref,
                            &clear_ref,
                            &contains_ref,
                        );
                        obj.lock()
                            .unwrap()
                            .properties
                            .insert("controls".into(), controls);
                    }
                    if type_name == "Form" {
                        let components = create_collection_object(
                            "components",
                            Some(owner.clone()),
                            &add_ref,
                            &clear_ref,
                            &contains_ref,
                        );
                        obj.lock()
                            .unwrap()
                            .properties
                            .insert("components".into(), components);
                    }
                    owner
                }),
            );
        }
    }

    fn str_arg(args: &[Value], idx: usize, default: &str) -> String {
        args.get(idx)
            .map(|v| format!("{}", v))
            .unwrap_or_else(|| default.into())
    }

    /// Extract `(__control_name, is_form)` from `args[0]`. Used by the
    /// control lifecycle host fns to find the target control. `is_form` is
    /// `true` when the object's `__type` is `"Form"` (or starts with
    /// `"Form"` for user subclasses) — used to gate form-specific side
    /// effects like `should_run = true` on `Show()`.
    fn read_this_identity(args: &[Value]) -> (String, bool) {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();
            let name = o
                .properties
                .get("__control_name")
                .map(|v| format!("{}", v).to_lowercase())
                .unwrap_or_default();
            let type_str = o
                .properties
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
        let idx = *vm
            .host_registry
            .get(&("vybe:gui".into(), name.into()))
            .unwrap();
        let mut o = vybe_bytecode::value::Object::new();
        o.kind = vybe_bytecode::value::ObjectKind::HostFunction(idx);
        Value::Object(vybe_bytecode::heap::alloc(o))
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

    fn apply_property(
        form: &mut vybe_widgets::Form,
        control_name: &str,
        property: &str,
        value: &str,
    ) {
        use vybe_widgets::{CommandValue, WidgetCommand};
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
                form.send_command(
                    control_name,
                    &WidgetCommand::Custom("SetReadOnly".into(), CommandValue::Bool(ro)),
                );
            }
            _ => {
                form.send_command(
                    control_name,
                    &WidgetCommand::Custom(
                        format!("Set{}", capitalize_first(property)),
                        CommandValue::Text(value.to_string()),
                    ),
                );
            }
        }
    }

    #[cfg(test)]
    mod tests {
        use super::*;

        #[test]
        fn live_empty_textbox_text_returns_empty_string_not_null() {
            let mut gui = GuiState::new();
            gui.add_widget("TextBox", "txtCalc", "", 0, 0, 100, 30);

            let value =
                control_property_value_from_live_or_fallback(&mut gui, "txtCalc", "Text", None);

            match value {
                Value::String(text) => assert!(text.is_empty()),
                other => panic!("expected empty text string, got {:?}", other),
            }
        }

        #[test]
        fn refresh_form_component_collection_only_collects_non_visual_components() {
            let add_ref = Value::Null;
            let clear_ref = Value::Null;
            let contains_ref = Value::Null;

            let form = vybe_bytecode::heap::alloc(Object::new());
            let owner = Value::Object(form.clone());
            let components = create_collection_object(
                "components",
                Some(owner),
                &add_ref,
                &clear_ref,
                &contains_ref,
            );
            form.lock()
                .unwrap()
                .properties
                .insert("components".into(), components.clone());

            let binding_source = vybe_bytecode::heap::alloc(Object::new());
            binding_source
                .lock()
                .unwrap()
                .properties
                .insert("__type".into(), Value::String(Arc::from("BindingSource")));
            form.lock()
                .unwrap()
                .properties
                .insert("bs1".into(), Value::Object(binding_source.clone()));

            let button = vybe_bytecode::heap::alloc(Object::new());
            button
                .lock()
                .unwrap()
                .properties
                .insert("__type".into(), Value::String(Arc::from("Button")));
            form.lock()
                .unwrap()
                .properties
                .insert("btn1".into(), Value::Object(button));

            refresh_form_component_collection(&form);

            let count = match components {
                Value::Object(collection) => {
                    let collection = collection.lock().unwrap();
                    match &collection.kind {
                        ObjectKind::Array(items) => items.len(),
                        _ => 0,
                    }
                }
                _ => 0,
            };

            assert_eq!(count, 1);
        }

        #[test]
        fn renamed_control_keeps_properties_and_events_reachable() {
            let mut gui = GuiState::new();
            gui.add_widget("Button", "Button_1", "", 0, 0, 100, 30);
            gui.set_property("Button_1", "Text", "Before");
            gui.register_event("Button_1", "Click", Value::Null);

            gui.rename_control("Button_1", "btnOk");

            assert_eq!(gui.get_property("btnOk", "Text"), "Before");
            assert!(gui.get_event_handler("btnOk", "Click").is_some());
            assert!(gui.get_event_handler("Button_1", "Click").is_some());

            gui.set_property("btnOk", "Text", "After");
            assert_eq!(gui.get_property("btnOk", "Text"), "After");
        }
    }
} // mod gui_impl

// Public re-export when gui feature is on
#[cfg(feature = "gui")]
pub use gui_impl::register;

// Non-GUI fallback: register stubs so compiled code does not crash.
#[cfg(not(feature = "gui"))]
pub fn register(vm: &mut vybe_bytecode::VM) {
    use vybe_bytecode::Value;
    let stubs = [
        "createForm",
        "addControl",
        "setProperty",
        "getProperty",
        "onEvent",
        "showForm",
        "runApplication",
        "msgBox",
        "closeForm",
        "appExit",
        "newControl",
        "controlSetProperty",
        "controlsAdd",
        "newForm",
        "newControlsCollection",
        "newComponentsCollection",
        "__collection_add",
        "__collection_clear",
        "__collection_contains",
        "noop",
        "addHandler",
        "removeHandler",
        "__ctrl_show",
        "__ctrl_close",
        "__ctrl_focus",
        "__ctrl_hide",
        "__dlg_showdialog",
        "__dlg_show",
    ];
    for name in stubs {
        vm.register_host_fn("vybe:gui", name, Box::new(|_ctx, _| Value::Null));
    }
}
