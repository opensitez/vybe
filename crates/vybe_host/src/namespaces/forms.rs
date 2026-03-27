use super::*;

pub fn register(vm: &mut VM) {
    let forms = ensure_namespace(vm, &["Window", "Forms"]);

    let control_types = [
        "Button", "Label", "TextBox", "CheckBox", "RadioButton",
        "ComboBox", "ListBox", "Panel", "GroupBox", "TabControl",
        "TabPage", "DataGridView", "ProgressBar", "TrackBar",
        "NumericUpDown", "DateTimePicker", "RichTextBox", "PictureBox",
        "MenuStrip", "ToolStrip", "StatusStrip", "SplitContainer",
        "FlowLayoutPanel", "TableLayoutPanel", "LinkLabel", "MaskedTextBox",
        "ListView", "WebBrowser", "MonthCalendar", "ContextMenuStrip",
        "Timer", "BindingSource", "ToolTip", "ImageList",
        "OpenFileDialog", "SaveFileDialog", "FolderBrowserDialog",
        "ColorDialog", "FontDialog",
    ];

    for type_name in &control_types {
        let hn = format!("new_{}", type_name);
        let type_str = type_name.to_string();
        vm.register_host_fn("vybe:gui", &hn, {
            let type_str = type_str.clone();
            Box::new(move |_args: &[Value]| {
                use vybe_bytecode::value::Object;
                use std::sync::atomic::{AtomicU32, Ordering};
                static COUNTER: AtomicU32 = AtomicU32::new(1);
                let id = COUNTER.fetch_add(1, Ordering::Relaxed);
                let name = format!("{}_{}", type_str, id);
                let mut obj = Object::new();
                obj.properties.insert("__control_type".into(), Value::String(Rc::from(type_str.as_str())));
                obj.properties.insert("__control_name".into(), Value::String(Rc::from(name.as_str())));
                obj.properties.insert("name".into(), Value::String(Rc::from(name.as_str())));
                obj.properties.insert("width".into(), Value::F64(100.0));
                obj.properties.insert("height".into(), Value::F64(30.0));
                obj.properties.insert("left".into(), Value::F64(0.0));
                obj.properties.insert("top".into(), Value::F64(0.0));
                Value::Object(Rc::new(RefCell::new(obj)))
            })
        });
        set_prop(&forms, &type_name.to_lowercase(), host_fn_ref(vm, "vybe:gui", &hn));
    }

    set_prop(&forms, "form", host_fn_ref(vm, "vybe:gui", "newForm"));
}
