//! `vybe:gui` no-op stub host functions.
//!
//! In headless / non-GUI contexts (no `gui` feature, or `Gui` capability not
//! granted) compiled code may still emit `vybe:gui` control/form calls. These
//! stubs register no-op (or property-mirroring) versions so that code doesn't
//! fail with "Unresolved import". Moved out of `vybe_host` — this is `vybe:gui`
//! functionality and belongs in the vybe platform crate.

use std::sync::Arc;
use vybe_runtime::{VM, Value};

/// Register the `vybe:gui` no-op stubs on `vm`.
pub fn register_gui_stubs(vm: &mut VM) {
    // Non-GUI stub that still mirrors the property write onto the object's
    // properties dict — this is essential because the dotnet class wrappers
    // emit setter chunks that call this fn, and user code (and tests) read
    // back the values via `obj.field`. The real GUI version of this fn
    // (`vybe_host::modules::gui::register::controlSetProperty`) ALSO writes
    // to a separate gui_state property store, which we skip here because
    // we have no `GuiState` to write into.
    vm.register_host_fn(
        "vybe:gui",
        "controlSetProperty",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(obj)) = args.first() {
                let property = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
                let val = args.get(2).cloned().unwrap_or(Value::Null);
                let prop_lower = property.to_lowercase();
                let mut o = obj.lock().unwrap();
                o.properties.insert(prop_lower.clone(), val.clone());
                if prop_lower == "name" {
                    o.properties.insert("__control_name".into(), val);
                }
            }
            Value::Null
        }),
    );
    vm.register_host_fn(
        "vybe:gui",
        "controlGetProperty",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(obj)) = args.first() {
                let property = args
                    .get(1)
                    .map(|v| format!("{}", v).to_lowercase())
                    .unwrap_or_default();
                return obj
                    .lock()
                    .unwrap()
                    .properties
                    .get(&property)
                    .cloned()
                    .unwrap_or(Value::Null);
            }
            Value::Null
        }),
    );
    vm.register_host_fn("vybe:gui", "setProperty", Box::new(|_ctx, _| Value::Null));
    vm.register_host_fn("vybe:gui", "showForm", Box::new(|_ctx, _| Value::Null));
    vm.register_host_fn("vybe:gui", "closeForm", Box::new(|_ctx, _| Value::Null));
    vm.register_host_fn(
        "vybe:gui",
        "showFormDialog",
        Box::new(|_ctx, _| Value::Null),
    );
    vm.register_host_fn("vybe:gui", "noop", Box::new(|_ctx, _| Value::Null));
    vm.register_host_fn(
        "vybe:gui",
        "runApplication",
        Box::new(|_ctx, _| Value::Null),
    );
    vm.register_host_fn("vybe:gui", "onEvent", Box::new(|_ctx, _| Value::Null));
    vm.register_host_fn("vybe:gui", "controlsAdd", Box::new(|_ctx, _| Value::Null));
    vm.register_host_fn(
        "vybe:gui",
        "newControlsCollection",
        Box::new(|_ctx, args| {
            use vybe_runtime::value::Object;
            let owner = args.first().cloned();
            let mut collection = Object::new_array(vec![]);
            collection.properties.insert(
                "__type".into(),
                Value::String(Arc::from("ControlCollection")),
            );
            if let Some(owner) = owner {
                collection.properties.insert("__owner".into(), owner);
            }
            collection
                .properties
                .insert("count".into(), Value::F64(0.0));
            Value::Object(vybe_runtime::heap::alloc(collection))
        }),
    );
    vm.register_host_fn(
        "vybe:gui",
        "newComponentsCollection",
        Box::new(|_ctx, args| {
            use vybe_runtime::value::Object;
            let owner = args.first().cloned();
            let mut collection = Object::new_array(vec![]);
            collection.properties.insert(
                "__type".into(),
                Value::String(Arc::from("ComponentCollection")),
            );
            if let Some(owner) = owner {
                collection.properties.insert("__owner".into(), owner);
            }
            collection
                .properties
                .insert("count".into(), Value::F64(0.0));
            Value::Object(vybe_runtime::heap::alloc(collection))
        }),
    );
    vm.register_host_fn(
        "vybe:gui",
        "__collection_add",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(collection)) = args.first() {
                let value = args.get(1).cloned().unwrap_or(Value::Null);
                let mut collection = collection.lock().unwrap();
                let mut len = None;
                if let vybe_runtime::value::ObjectKind::Array(items) = &mut collection.kind {
                    if !items.iter().any(|existing| existing.eq(&value)) {
                        items.push(value);
                    }
                    len = Some(items.len());
                }
                if let Some(len) = len {
                    collection
                        .properties
                        .insert("count".into(), Value::F64(len as f64));
                    collection
                        .properties
                        .insert("length".into(), Value::F64(len as f64));
                }
            }
            Value::Null
        }),
    );
    vm.register_host_fn(
        "vybe:gui",
        "__collection_clear",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(collection)) = args.first() {
                let mut collection = collection.lock().unwrap();
                if let vybe_runtime::value::ObjectKind::Array(items) = &mut collection.kind {
                    items.clear();
                }
                collection
                    .properties
                    .insert("count".into(), Value::F64(0.0));
                collection
                    .properties
                    .insert("length".into(), Value::F64(0.0));
            }
            Value::Null
        }),
    );
    vm.register_host_fn(
        "vybe:gui",
        "__collection_contains",
        Box::new(|_ctx, args| {
            let Some(Value::Object(collection)) = args.first() else {
                return Value::Bool(false);
            };
            let needle = args.get(1).cloned().unwrap_or(Value::Null);
            let collection = collection.lock().unwrap();
            let contains = if let vybe_runtime::value::ObjectKind::Array(items) = &collection.kind {
                items.iter().any(|existing| existing.eq(&needle))
            } else {
                false
            };
            Value::Bool(contains)
        }),
    );
    vm.register_host_fn(
        "vybe:gui",
        "newForm",
        Box::new(|_ctx, args| {
            use vybe_runtime::value::Object;
            let title = args.first().map(|v| format!("{v}")).unwrap_or_default();
            let mut obj = Object::new();
            obj.properties
                .insert("__control_type".into(), Value::String(Arc::from("Form")));
            obj.properties
                .insert("text".into(), Value::String(Arc::from(title.as_str())));
            obj.properties
                .insert("name".into(), Value::String(Arc::from("form")));
            // Controls collection (no-op stub)
            let mut ctrls = Object::new_array(vec![]);
            ctrls.properties.insert(
                "__type".into(),
                Value::String(Arc::from("ControlCollection")),
            );
            ctrls.properties.insert("count".into(), Value::F64(0.0));
            obj.properties.insert(
                "controls".into(),
                Value::Object(vybe_runtime::heap::alloc(ctrls)),
            );
            let mut comps = Object::new_array(vec![]);
            comps.properties.insert(
                "__type".into(),
                Value::String(Arc::from("ComponentCollection")),
            );
            comps.properties.insert("count".into(), Value::F64(0.0));
            obj.properties.insert(
                "components".into(),
                Value::Object(vybe_runtime::heap::alloc(comps)),
            );
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );
    // ── Control / Form method stubs for the dotnet class wrappers ──
    // These are bound by `compiler_common::dotnet::classes::control::CONTROL_METHODS`
    // and `form::FORM_METHODS` as method thunks. Without the host
    // import target the VM would trap on unresolved import even
    // though no test actually exercises window lifecycle.
    for fn_name in &[
        "__ctrl_show",
        "__ctrl_hide",
        "__ctrl_focus",
        "__ctrl_close",
        "__ctrl_refresh",
        "__ctrl_invalidate",
        "__ctrl_update",
        "__ctrl_bring_to_front",
        "__ctrl_send_to_back",
        "__ctrl_dispose",
        "__form_activate",
        "__form_center_to_screen",
        "__dlg_showdialog",
        "__dlg_show",
    ] {
        vm.register_host_fn("vybe:gui", fn_name, Box::new(|_ctx, _| Value::Null));
    }

    // ── Per-control `new_<Type>` stubs for the dotnet class wrappers ──
    // The compiler_common::dotnet::classes layer emits ctor chunks that
    // call `vybe:gui::new_<ClassName>` for every concrete leaf. In test
    // / non-GUI contexts the real `gui::register` isn't called, so we
    // install no-op stubs that return a minimally-populated control
    // object — enough for the dotnet ctor's "transfer widget identity"
    // step to succeed without panicking.
    let dotnet_concrete_controls: &[&str] = &[
        "Form",
        // Buttons family
        "Button",
        "CheckBox",
        "RadioButton",
        // Text family
        "TextBox",
        "RichTextBox",
        "MaskedTextBox",
        // Labels
        "Label",
        "LinkLabel",
        // Lists
        "ComboBox",
        "ListBox",
        "ListView",
        "TreeView",
        // Containers
        "Panel",
        "GroupBox",
        "TabControl",
        "TabPage",
        "SplitContainer",
        "FlowLayoutPanel",
        "TableLayoutPanel",
        // Progress
        "ProgressBar",
        "TrackBar",
        "NumericUpDown",
        // Dates
        "DateTimePicker",
        "MonthCalendar",
        // Media
        "PictureBox",
        "WebBrowser",
        // Grids
        "DataGridView",
        // Strips
        "ToolStrip",
        "MenuStrip",
        "StatusStrip",
        "ContextMenuStrip",
        // Non-visual
        "Timer",
        "BindingSource",
        "ImageList",
        "ToolTip",
        "NotifyIcon",
        "ErrorProvider",
        "HelpProvider",
        "BackgroundWorker",
        // Dialogs
        "OpenFileDialog",
        "SaveFileDialog",
        "FontDialog",
        "ColorDialog",
        "FolderBrowserDialog",
        // Drawing
        "Canvas",
    ];
    for ct in dotnet_concrete_controls {
        let type_name = ct.to_string();
        vm.register_host_fn(
            "vybe:gui",
            &format!("new_{}", ct),
            Box::new(move |_ctx, _args| {
                use std::sync::atomic::{AtomicU32, Ordering};
                use vybe_runtime::value::Object;
                static COUNTER: AtomicU32 = AtomicU32::new(1);
                let id = COUNTER.fetch_add(1, Ordering::Relaxed);
                let name = format!("{}_{}", type_name.to_lowercase(), id);
                let mut obj = Object::new();
                obj.properties.insert(
                    "__control_type".into(),
                    Value::String(Arc::from(type_name.as_str())),
                );
                obj.properties.insert(
                    "__control_name".into(),
                    Value::String(Arc::from(name.as_str())),
                );
                obj.properties.insert(
                    "__type".into(),
                    Value::String(Arc::from(type_name.as_str())),
                );
                obj.properties
                    .insert("name".into(), Value::String(Arc::from(name.as_str())));
                Value::Object(vybe_runtime::heap::alloc(obj))
            }),
        );
    }
}
