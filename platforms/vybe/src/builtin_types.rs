//! `vybe:gui` built-in **types** — the runtime TypeRegistry vtables for the
//! WinForms-shaped control hierarchy (`Control` + concrete controls + `Form`),
//! the associated WinForms enums (`DialogResult`, `Keys`, …), and the control
//! constructors (`new_Button`, …).
//!
//! The `register_type` counterpart to the vybe plugin's `vybe:gui` host-fn
//! `init`: the plugin declares its own types here, in its `finalize`. Control
//! methods/ctors resolve `vybe:gui` host fns by registry index when the `Gui`
//! capability is granted; without it those fns are absent and the entries are
//! simply skipped, but the TypeDefs still register (as before).

use vybe_runtime::Framework;
use vybe_runtime::{Method, TypeDef};

/// The concrete control types — each a subtype of `Control`, each with a
/// `new_<Name>` constructor. (Several `.NET` non-visual components — `Timer`,
/// `BindingSource`, `DataSet`, `ImageList`, `ToolTip`, … — are historically
/// modelled as controls here; kept as-is.)
const CONTROL_TYPE_NAMES: &[&str] = &[
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
    "TableLayoutPanel",
    "LinkLabel",
    "MaskedTextBox",
    "ListView",
    "WebBrowser",
    "MonthCalendar",
    "ContextMenuStrip",
    "Timer",
    "BindingSource",
    "DataSet",
    "ImageList",
    "ToolTip",
    "NotifyIcon",
    "ErrorProvider",
    "HelpProvider",
    "BackgroundWorker",
    "TreeView",
];

/// Register the `vybe:gui` control hierarchy, enums, and constructors into the
/// VM's TypeRegistry. Called from the vybe plugin's `finalize`.
pub fn register_types(fw: &mut Framework<'_>) {
    // Idempotent: the gui flow runs the vybe plugin's `finalize` twice — once
    // for the base (drawing-only) plugin, then again for the gui-variant
    // plugin that owns the `GuiState`. Register the control/enum surface once.
    if fw.type_id("Control").is_some() {
        return;
    }

    // --- Control (abstract base for all UI controls) ---
    let control_id = {
        let mut t = TypeDef::new("Control");
        for (method, fname) in &[
            ("show", "__ctrl_show"),
            ("close", "__ctrl_close"),
            ("focus", "__ctrl_focus"),
            ("hide", "__ctrl_hide"),
        ] {
            if let Some(idx) = fw.host_fn_index("vybe:gui", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0); // inherits from Object
        fw.register_type(t)
    };

    // Concrete control types — subtypes of Control, each with its `new_<Name>`
    // constructor bound at build time.
    for ct in CONTROL_TYPE_NAMES {
        let mut t = TypeDef::new(ct);
        t.parent = Some(control_id);
        if let Some(idx) = fw.host_fn_index("vybe:gui", &format!("new_{ct}")) {
            t.constructor = Some(Method::HostFn(idx));
        }
        fw.register_type(t);
    }

    // Form — inherits from Control, adds its own methods + `new_Form` ctor.
    {
        let mut t = TypeDef::new("Form");
        for (method, fname) in &[("show", "__ctrl_show"), ("close", "__ctrl_close")] {
            if let Some(idx) = fw.host_fn_index("vybe:gui", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        if let Some(idx) = fw.host_fn_index("vybe:gui", "new_Form") {
            t.constructor = Some(Method::HostFn(idx));
        }
        t.parent = Some(control_id);
        fw.register_type(t);
    }

    // --- WinForms enums (compile-time constants) ---
    register_enum(
        fw,
        "DialogResult",
        &[
            ("none", 0),
            ("ok", 1),
            ("cancel", 2),
            ("abort", 3),
            ("retry", 4),
            ("ignore", 5),
            ("yes", 6),
            ("no", 7),
        ],
    );
    register_enum(
        fw,
        "MessageBoxButtons",
        &[
            ("ok", 0),
            ("okcancel", 1),
            ("abortretryignore", 2),
            ("yesnocancel", 3),
            ("yesno", 4),
            ("retrycancel", 5),
        ],
    );
    register_enum(
        fw,
        "MessageBoxIcon",
        &[
            ("none", 0),
            ("error", 16),
            ("question", 32),
            ("warning", 48),
            ("information", 64),
        ],
    );
    register_enum(
        fw,
        "Keys",
        &[
            ("none", 0),
            ("back", 8),
            ("tab", 9),
            ("return", 13),
            ("enter", 13),
            ("escape", 27),
            ("space", 32),
            ("left", 37),
            ("up", 38),
            ("right", 39),
            ("down", 40),
            ("delete", 46),
            ("insert", 45),
            ("shift", 16),
            ("control", 17),
            ("alt", 18),
            ("f1", 112),
            ("f2", 113),
            ("f3", 114),
            ("f4", 115),
            ("f5", 116),
            ("f6", 117),
            ("f7", 118),
            ("f8", 119),
            ("f9", 120),
            ("f10", 121),
            ("f11", 122),
            ("f12", 123),
        ],
    );
    register_enum(
        fw,
        "FormBorderStyle",
        &[
            ("none", 0),
            ("fixedsingle", 1),
            ("fixeddialog", 3),
            ("sizable", 4),
            ("fixedtoolwindow", 5),
            ("sizabletoolwindow", 6),
        ],
    );
    register_enum(
        fw,
        "FormStartPosition",
        &[
            ("manual", 0),
            ("centerscreen", 1),
            ("windowsdefaultlocation", 2),
            ("windowsdefaultbounds", 3),
            ("centerparent", 4),
        ],
    );
    register_enum(
        fw,
        "FormWindowState",
        &[("normal", 0), ("minimized", 1), ("maximized", 2)],
    );
    register_enum(
        fw,
        "DockStyle",
        &[
            ("none", 0),
            ("top", 1),
            ("bottom", 2),
            ("left", 3),
            ("right", 4),
            ("fill", 5),
        ],
    );
    register_enum(
        fw,
        "AnchorStyles",
        &[
            ("none", 0),
            ("top", 1),
            ("bottom", 2),
            ("left", 4),
            ("right", 8),
        ],
    );
    register_enum(
        fw,
        "CloseReason",
        &[
            ("none", 0),
            ("windowsshutdown", 1),
            ("userclosing", 3),
            ("applicationexitcall", 5),
        ],
    );
    register_enum(
        fw,
        "MouseButtons",
        &[("none", 0), ("left", 1), ("right", 2), ("middle", 4)],
    );

    // No value-type constructors here. `Point`/`Size`/`Font` were bound to
    // `vybe:gui` host fns that no longer exist, and the binding was already a
    // no-op — their TypeDefs are not registered, as the comment that stood here
    // admitted. A value type is composed in BYTECODE (`common_ctor_for` →
    // `dotnet.point_new`), because it has no element and nothing to insert.
}

/// Register an enum TypeDef with its compile-time integer constants.
fn register_enum(fw: &mut Framework<'_>, name: &str, entries: &[(&str, i64)]) {
    let mut t = TypeDef::new(name);
    for (k, v) in entries {
        t.constants.insert(k.to_string(), *v);
    }
    fw.register_type(t);
}
