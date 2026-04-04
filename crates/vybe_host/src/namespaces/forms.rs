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
        // Only create namespace references — host functions are registered by gui::register
        // If not registered yet (non-GUI context), register a stub
        if vm.host_registry.get(&("vybe:gui".to_string(), hn.clone())).is_none() {
            let type_str = type_name.to_string();
            vm.register_host_fn("vybe:gui", &hn, {
                let type_str = type_str.clone();
                Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
                    use vybe_bytecode::value::Object;
                    let mut obj = Object::new();
                    obj.properties.insert("__control_type".into(), Value::String(Rc::from(type_str.as_str())));
                    obj.properties.insert("name".into(), Value::String(Rc::from(type_str.to_lowercase().as_str())));
                    Value::Object(Rc::new(RefCell::new(obj)))
                })
            });
        }
        set_prop(&forms, &type_name.to_lowercase(), host_fn_ref(vm, "vybe:gui", &hn));
        // Also register as bare global: Button, TextBox, etc.
        let bare_fn = host_fn_ref(vm, "vybe:gui", &hn);
        vm.globals.insert(type_name.to_lowercase(), bare_fn);
    }

    set_prop(&forms, "form", host_fn_ref(vm, "vybe:gui", "newForm"));
    vm.globals.insert("form".into(), host_fn_ref(vm, "vybe:gui", "newForm"));

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

    // BorderStyle (control-level: None=0, FixedSingle=1, Fixed3D=2)
    let bs = ensure_namespace(vm, &["Window", "Forms", "BorderStyle"]);
    set_prop(&bs, "none", Value::F64(0.0));
    set_prop(&bs, "fixedsingle", Value::F64(1.0));
    set_prop(&bs, "fixed3d", Value::F64(2.0));

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

    // Color.FromArgb — callable from namespace
    set_prop(&color, "fromargb", host_fn_ref(vm, "vybe:drawing", "color.fromargb"));

    // ColorTranslator
    let ct = ensure_namespace(vm, &["System", "Drawing", "ColorTranslator"]);
    set_prop(&ct, "fromhtml", host_fn_ref(vm, "vybe:drawing", "colortranslator.fromhtml"));

    // Color shortcut
    let color_short = ensure_namespace(vm, &["Color"]);
    set_prop(&color_short, "black", Value::F64(0x000000 as f64));
    set_prop(&color_short, "white", Value::F64(0xFFFFFF as f64));
    set_prop(&color_short, "red", Value::F64(0xFF0000 as f64));
    set_prop(&color_short, "green", Value::F64(0x00FF00 as f64));
    set_prop(&color_short, "blue", Value::F64(0x0000FF as f64));
    set_prop(&color_short, "yellow", Value::F64(0xFFFF00 as f64));
    set_prop(&color_short, "gray", Value::F64(0x808080 as f64));
    set_prop(&color_short, "fromargb", host_fn_ref(vm, "vybe:drawing", "color.fromargb"));

    // BorderStyle shortcut (bare name)
    let bs_short = ensure_namespace(vm, &["BorderStyle"]);
    set_prop(&bs_short, "none", Value::F64(0.0));
    set_prop(&bs_short, "fixedsingle", Value::F64(1.0));
    set_prop(&bs_short, "fixed3d", Value::F64(2.0));

    // FormBorderStyle shortcut (bare name)
    let fbs_short = ensure_namespace(vm, &["FormBorderStyle"]);
    set_prop(&fbs_short, "none", Value::F64(0.0));
    set_prop(&fbs_short, "fixedsingle", Value::F64(1.0));
    set_prop(&fbs_short, "fixed3d", Value::F64(2.0));
    set_prop(&fbs_short, "sizable", Value::F64(4.0));
    set_prop(&fbs_short, "fixedtoolwindow", Value::F64(5.0));
    set_prop(&fbs_short, "sizabletoolwindow", Value::F64(6.0));

    // ContentAlignment shortcut (bare name)
    let ca_short = ensure_namespace(vm, &["ContentAlignment"]);
    set_prop(&ca_short, "topleft", Value::F64(1.0));
    set_prop(&ca_short, "topcenter", Value::F64(2.0));
    set_prop(&ca_short, "topright", Value::F64(4.0));
    set_prop(&ca_short, "middleleft", Value::F64(16.0));
    set_prop(&ca_short, "middlecenter", Value::F64(32.0));
    set_prop(&ca_short, "middleright", Value::F64(64.0));
    set_prop(&ca_short, "bottomleft", Value::F64(256.0));
    set_prop(&ca_short, "bottomcenter", Value::F64(512.0));
    set_prop(&ca_short, "bottomright", Value::F64(1024.0));

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

    // --- Additional missing enums/classes ---

    // CloseReason
    let cr = ensure_namespace(vm, &["Window", "Forms", "CloseReason"]);
    set_prop(&cr, "none", Value::F64(0.0));
    set_prop(&cr, "windowsshutdown", Value::F64(1.0));
    set_prop(&cr, "userclosing", Value::F64(3.0));
    set_prop(&cr, "applicationexitcall", Value::F64(5.0));

    // MouseButtons
    let mb = ensure_namespace(vm, &["Window", "Forms", "MouseButtons"]);
    set_prop(&mb, "none", Value::F64(0.0));
    set_prop(&mb, "left", Value::F64(1.0));
    set_prop(&mb, "right", Value::F64(2.0));
    set_prop(&mb, "middle", Value::F64(4.0));

    // More Keys
    let keys = ensure_namespace(vm, &["Window", "Forms", "Keys"]);
    set_prop(&keys, "shift", Value::F64(16.0));
    set_prop(&keys, "control", Value::F64(17.0));
    set_prop(&keys, "alt", Value::F64(18.0));
    set_prop(&keys, "f1", Value::F64(112.0));
    set_prop(&keys, "f2", Value::F64(113.0));
    set_prop(&keys, "f3", Value::F64(114.0));
    set_prop(&keys, "f4", Value::F64(115.0));
    set_prop(&keys, "f5", Value::F64(116.0));
    set_prop(&keys, "f6", Value::F64(117.0));
    set_prop(&keys, "f7", Value::F64(118.0));
    set_prop(&keys, "f8", Value::F64(119.0));
    set_prop(&keys, "f9", Value::F64(120.0));
    set_prop(&keys, "f10", Value::F64(121.0));
    set_prop(&keys, "f11", Value::F64(122.0));
    set_prop(&keys, "f12", Value::F64(123.0));

    // MessageBox.Show
    let msgbox = ensure_namespace(vm, &["Window", "Forms", "MessageBox"]);
    set_prop(&msgbox, "show", host_fn_ref(vm, "vybe:gui", "msgBox"));

    // Application.Exit
    let app = ensure_namespace(vm, &["Window", "Forms", "Application"]);
    set_prop(&app, "run", host_fn_ref(vm, "vybe:gui", "runApplication"));
    set_prop(&app, "exit", host_fn_ref(vm, "wasi:cli", "exit"));

    // Event args — empty objects
    let null = Value::Null;
    let swf = ensure_namespace(vm, &["Window", "Forms"]);
    set_prop(&swf, "keyeventargs", null.clone());
    set_prop(&swf, "keypresseventargs", null.clone());
    set_prop(&swf, "mouseeventargs", null.clone());
    set_prop(&swf, "painteventargs", null.clone());
    set_prop(&swf, "formclosedeventargs", null.clone());
    set_prop(&swf, "formclosingeventargs", null);

    // --- Mirror under System.Windows.Forms.* ---
    // All the above are under Window.Forms, also make them available under System.Windows.Forms
    let _sys_wf = ensure_namespace(vm, &["System", "Windows", "Forms"]);

    // DialogResult
    let sys_dr = ensure_namespace(vm, &["System", "Windows", "Forms", "DialogResult"]);
    for (k, v) in &[("none",0),("ok",1),("cancel",2),("abort",3),("retry",4),("ignore",5),("yes",6),("no",7)] {
        set_prop(&sys_dr, k, Value::F64(*v as f64));
    }

    // MessageBoxButtons
    let sys_mbb = ensure_namespace(vm, &["System", "Windows", "Forms", "MessageBoxButtons"]);
    for (k, v) in &[("ok",0),("okcancel",1),("abortretryignore",2),("yesnocancel",3),("yesno",4),("retrycancel",5)] {
        set_prop(&sys_mbb, k, Value::F64(*v as f64));
    }

    // MessageBoxIcon
    let sys_mbi = ensure_namespace(vm, &["System", "Windows", "Forms", "MessageBoxIcon"]);
    for (k, v) in &[("none",0),("error",16),("question",32),("warning",48),("information",64)] {
        set_prop(&sys_mbi, k, Value::F64(*v as f64));
    }

    // Keys
    let sys_keys = ensure_namespace(vm, &["System", "Windows", "Forms", "Keys"]);
    for (k, v) in &[
        ("none",0),("back",8),("tab",9),("return",13),("enter",13),("escape",27),
        ("space",32),("left",37),("up",38),("right",39),("down",40),
        ("delete",46),("insert",45),("shift",16),("shiftkey",16),
        ("control",17),("controlkey",17),("alt",18),("menu",18),
        ("f1",112),("f2",113),("f3",114),("f4",115),("f5",116),("f6",117),
        ("f7",118),("f8",119),("f9",120),("f10",121),("f11",122),("f12",123),
    ] {
        set_prop(&sys_keys, k, Value::F64(*v as f64));
    }

    // DockStyle
    let sys_ds = ensure_namespace(vm, &["System", "Windows", "Forms", "DockStyle"]);
    for (k, v) in &[("none",0),("top",1),("bottom",2),("left",3),("right",4),("fill",5)] {
        set_prop(&sys_ds, k, Value::F64(*v as f64));
    }

    // AnchorStyles
    let sys_as = ensure_namespace(vm, &["System", "Windows", "Forms", "AnchorStyles"]);
    for (k, v) in &[("none",0),("top",1),("bottom",2),("left",4),("right",8)] {
        set_prop(&sys_as, k, Value::F64(*v as f64));
    }

    // FormBorderStyle
    let sys_fbs = ensure_namespace(vm, &["System", "Windows", "Forms", "FormBorderStyle"]);
    for (k, v) in &[("none",0),("fixedsingle",1),("fixeddialog",3),("sizable",4),("fixedtoolwindow",5),("sizabletoolwindow",6)] {
        set_prop(&sys_fbs, k, Value::F64(*v as f64));
    }

    // FormStartPosition
    let sys_fsp = ensure_namespace(vm, &["System", "Windows", "Forms", "FormStartPosition"]);
    for (k, v) in &[("manual",0),("centerscreen",1),("windowsdefaultlocation",2),("windowsdefaultbounds",3),("centerparent",4)] {
        set_prop(&sys_fsp, k, Value::F64(*v as f64));
    }

    // FormWindowState
    let sys_fws = ensure_namespace(vm, &["System", "Windows", "Forms", "FormWindowState"]);
    for (k, v) in &[("normal",0),("minimized",1),("maximized",2)] {
        set_prop(&sys_fws, k, Value::F64(*v as f64));
    }

    // CloseReason
    let sys_cr = ensure_namespace(vm, &["System", "Windows", "Forms", "CloseReason"]);
    for (k, v) in &[("none",0),("windowsshutdown",1),("userclosing",3),("applicationexitcall",5)] {
        set_prop(&sys_cr, k, Value::F64(*v as f64));
    }

    // MouseButtons
    let sys_mb = ensure_namespace(vm, &["System", "Windows", "Forms", "MouseButtons"]);
    for (k, v) in &[("none",0),("left",1),("right",2),("middle",4)] {
        set_prop(&sys_mb, k, Value::F64(*v as f64));
    }

    // BorderStyle (control-level)
    let sys_bs = ensure_namespace(vm, &["System", "Windows", "Forms", "BorderStyle"]);
    for (k, v) in &[("none",0),("fixedsingle",1),("fixed3d",2)] {
        set_prop(&sys_bs, k, Value::F64(*v as f64));
    }

    // MessageBox
    let sys_msgbox = ensure_namespace(vm, &["System", "Windows", "Forms", "MessageBox"]);
    set_prop(&sys_msgbox, "show", host_fn_ref(vm, "vybe:gui", "msgBox"));

    // Application
    let sys_app = ensure_namespace(vm, &["System", "Windows", "Forms", "Application"]);
    set_prop(&sys_app, "run", host_fn_ref(vm, "vybe:gui", "runApplication"));
    set_prop(&sys_app, "exit", host_fn_ref(vm, "wasi:cli", "exit"));
}
