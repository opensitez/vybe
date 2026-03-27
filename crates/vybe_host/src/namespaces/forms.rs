use super::*;

pub fn register(vm: &mut VM) {
    let forms = ensure_namespace(vm, &["Window", "Forms"]);

    // Control type constructors
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

    // --- WinForms Constants & Enums ---

    // DialogResult
    let dr = ensure_namespace(vm, &["Window", "Forms", "DialogResult"]);
    set_prop(&dr, "none", Value::F64(0.0));
    set_prop(&dr, "ok", Value::F64(1.0));
    set_prop(&dr, "cancel", Value::F64(2.0));
    set_prop(&dr, "abort", Value::F64(3.0));
    set_prop(&dr, "retry", Value::F64(4.0));
    set_prop(&dr, "ignore", Value::F64(5.0));
    set_prop(&dr, "yes", Value::F64(6.0));
    set_prop(&dr, "no", Value::F64(7.0));

    // MessageBoxButtons
    let mbb = ensure_namespace(vm, &["Window", "Forms", "MessageBoxButtons"]);
    set_prop(&mbb, "ok", Value::F64(0.0));
    set_prop(&mbb, "okcancel", Value::F64(1.0));
    set_prop(&mbb, "abortretryignore", Value::F64(2.0));
    set_prop(&mbb, "yesnocancel", Value::F64(3.0));
    set_prop(&mbb, "yesno", Value::F64(4.0));
    set_prop(&mbb, "retrycancel", Value::F64(5.0));

    // MessageBoxIcon
    let mbi = ensure_namespace(vm, &["Window", "Forms", "MessageBoxIcon"]);
    set_prop(&mbi, "none", Value::F64(0.0));
    set_prop(&mbi, "error", Value::F64(16.0));
    set_prop(&mbi, "question", Value::F64(32.0));
    set_prop(&mbi, "warning", Value::F64(48.0));
    set_prop(&mbi, "information", Value::F64(64.0));

    // Keys
    let keys = ensure_namespace(vm, &["Window", "Forms", "Keys"]);
    set_prop(&keys, "none", Value::F64(0.0));
    set_prop(&keys, "back", Value::F64(8.0));
    set_prop(&keys, "tab", Value::F64(9.0));
    set_prop(&keys, "return", Value::F64(13.0));
    set_prop(&keys, "enter", Value::F64(13.0));
    set_prop(&keys, "escape", Value::F64(27.0));
    set_prop(&keys, "space", Value::F64(32.0));
    set_prop(&keys, "left", Value::F64(37.0));
    set_prop(&keys, "up", Value::F64(38.0));
    set_prop(&keys, "right", Value::F64(39.0));
    set_prop(&keys, "down", Value::F64(40.0));
    set_prop(&keys, "delete", Value::F64(46.0));
    set_prop(&keys, "insert", Value::F64(45.0));
    set_prop(&keys, "shiftkey", Value::F64(16.0));
    set_prop(&keys, "controlkey", Value::F64(17.0));
    set_prop(&keys, "menu", Value::F64(18.0));

    // FormBorderStyle
    let fbs = ensure_namespace(vm, &["Window", "Forms", "FormBorderStyle"]);
    set_prop(&fbs, "none", Value::F64(0.0));
    set_prop(&fbs, "fixedsingle", Value::F64(1.0));
    set_prop(&fbs, "sizable", Value::F64(4.0));
    set_prop(&fbs, "fixeddialog", Value::F64(3.0));
    set_prop(&fbs, "fixedtoolwindow", Value::F64(5.0));
    set_prop(&fbs, "sizabletoolwindow", Value::F64(6.0));

    // FormStartPosition
    let fsp = ensure_namespace(vm, &["Window", "Forms", "FormStartPosition"]);
    set_prop(&fsp, "manual", Value::F64(0.0));
    set_prop(&fsp, "centerscreen", Value::F64(1.0));
    set_prop(&fsp, "windowsdefaultlocation", Value::F64(2.0));
    set_prop(&fsp, "windowsdefaultbounds", Value::F64(3.0));
    set_prop(&fsp, "centerparent", Value::F64(4.0));

    // FormWindowState
    let fws = ensure_namespace(vm, &["Window", "Forms", "FormWindowState"]);
    set_prop(&fws, "normal", Value::F64(0.0));
    set_prop(&fws, "minimized", Value::F64(1.0));
    set_prop(&fws, "maximized", Value::F64(2.0));

    // DockStyle
    let ds = ensure_namespace(vm, &["Window", "Forms", "DockStyle"]);
    set_prop(&ds, "none", Value::F64(0.0));
    set_prop(&ds, "top", Value::F64(1.0));
    set_prop(&ds, "bottom", Value::F64(2.0));
    set_prop(&ds, "left", Value::F64(3.0));
    set_prop(&ds, "right", Value::F64(4.0));
    set_prop(&ds, "fill", Value::F64(5.0));

    // AnchorStyles
    let anch = ensure_namespace(vm, &["Window", "Forms", "AnchorStyles"]);
    set_prop(&anch, "none", Value::F64(0.0));
    set_prop(&anch, "top", Value::F64(1.0));
    set_prop(&anch, "bottom", Value::F64(2.0));
    set_prop(&anch, "left", Value::F64(4.0));
    set_prop(&anch, "right", Value::F64(8.0));

    // --- System.Drawing ---

    // Color
    let color = ensure_namespace(vm, &["System", "Drawing", "Color"]);
    set_prop(&color, "black", Value::F64(0x000000 as f64));
    set_prop(&color, "white", Value::F64(0xFFFFFF as f64));
    set_prop(&color, "red", Value::F64(0xFF0000 as f64));
    set_prop(&color, "green", Value::F64(0x00FF00 as f64));
    set_prop(&color, "blue", Value::F64(0x0000FF as f64));
    set_prop(&color, "yellow", Value::F64(0xFFFF00 as f64));
    set_prop(&color, "gray", Value::F64(0x808080 as f64));
    set_prop(&color, "darkgray", Value::F64(0xA9A9A9 as f64));
    set_prop(&color, "lightgray", Value::F64(0xD3D3D3 as f64));
    set_prop(&color, "transparent", Value::F64(-1.0));

    // Color shortcut
    let color_short = ensure_namespace(vm, &["Color"]);
    set_prop(&color_short, "black", Value::F64(0x000000 as f64));
    set_prop(&color_short, "white", Value::F64(0xFFFFFF as f64));
    set_prop(&color_short, "red", Value::F64(0xFF0000 as f64));
    set_prop(&color_short, "green", Value::F64(0x00FF00 as f64));
    set_prop(&color_short, "blue", Value::F64(0x0000FF as f64));
    set_prop(&color_short, "yellow", Value::F64(0xFFFF00 as f64));
    set_prop(&color_short, "gray", Value::F64(0x808080 as f64));

    // ContentAlignment
    let ca = ensure_namespace(vm, &["System", "Drawing", "ContentAlignment"]);
    set_prop(&ca, "topleft", Value::F64(1.0));
    set_prop(&ca, "topcenter", Value::F64(2.0));
    set_prop(&ca, "topright", Value::F64(4.0));
    set_prop(&ca, "middleleft", Value::F64(16.0));
    set_prop(&ca, "middlecenter", Value::F64(32.0));
    set_prop(&ca, "middleright", Value::F64(64.0));
    set_prop(&ca, "bottomleft", Value::F64(256.0));
    set_prop(&ca, "bottomcenter", Value::F64(512.0));
    set_prop(&ca, "bottomright", Value::F64(1024.0));
}
