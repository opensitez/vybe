//! GCL: Pascal GUI component library surface.
//!
//! This is a Pascal-shaped adapter for VCL/LCL-style source. It lowers
//! Pascal classes such as `TForm` and `TButton` to the existing generic
//! `vybe:gui` host imports without adding Pascal-specific host functions.

pub mod builder;

use vybe_compiler::compiler::gui;

#[derive(Debug, Clone, Copy)]
pub struct GclClass {
    pub name: &'static str,
    pub parent: Option<&'static str>,
    pub properties: &'static [&'static str],
    pub methods: &'static [GclMethod],
    pub ctor_arity: u8,
    pub widget_host_fn: Option<&'static str>,
}

#[derive(Debug, Clone, Copy)]
pub struct GclMethod {
    pub name: &'static str,
    pub arity: u8,
    pub target: GclMethodTarget,
}

#[derive(Debug, Clone, Copy)]
pub enum GclMethodTarget {
    Host {
        module: &'static str,
        fn_name: &'static str,
    },
}

impl GclMethodTarget {
    pub const fn host(fn_name: &'static str) -> Self {
        GclMethodTarget::Host {
            module: gui::GUI_MODULE,
            fn_name,
        }
    }
}

pub fn is_gcl_unit(path: &str) -> bool {
    matches!(
        path.to_ascii_lowercase().as_str(),
        "forms"
            | "controls"
            | "stdctrls"
            | "extctrls"
            | "comctrls"
            | "dialogs"
            | "grids"
            | "menus"
            | "graphics"
            | "buttons"
    )
}

const CONTROL_PROPERTIES: &[&str] = &[
    "Name",
    "Text",
    "Caption",
    "Left",
    "Top",
    "Width",
    "Height",
    "Enabled",
    "Visible",
    "Align",
    "Anchors",
    "Color",
    "Parent",
    "Hint",
    "ShowHint",
    "Tag",
    "OnClick",
    "OnChange",
    "OnCreate",
    "OnClose",
    "OnTimer",
    "OnKeyPress",
    "OnKeyDown",
    "OnKeyUp",
];

const FORM_PROPERTIES: &[&str] = &[
    "Name",
    "Text",
    "Caption",
    "BorderStyle",
    "Position",
    "MainMenu",
    "PopupMenu",
    "ClientWidth",
    "ClientHeight",
    "OnCreate",
    "OnClose",
];
const TEXT_PROPERTIES: &[&str] = &["PasswordChar", "ReadOnly", "MaxLength"];
const CHECK_PROPERTIES: &[&str] = &["Checked", "State"];
const LIST_PROPERTIES: &[&str] = &["Items", "ItemIndex", "Sorted"];
const GRID_PROPERTIES: &[&str] = &["ColCount", "RowCount", "FixedCols", "FixedRows", "Cells"];
const RANGE_PROPERTIES: &[&str] = &["Min", "Max", "Position", "Step"];
const SPIN_PROPERTIES: &[&str] = &["MinValue", "MaxValue", "Value", "Increment"];
const PAGE_PROPERTIES: &[&str] = &["ActivePage", "PageIndex", "PageControl"];
const MENU_PROPERTIES: &[&str] = &[
    "Caption", "Text", "Name", "ShortCut", "Checked", "Enabled", "Visible", "OnClick",
];

const SHOW_METHODS: &[GclMethod] = &[
    // Same targets as the dotnet WinForms adapters: `Show` marks the form
    // visible + should_run, `ShowModal` returns a DialogResult. The event
    // loop itself belongs to `Application.Run` (HOST_FN_RUN_APPLICATION) —
    // routing Show/ShowModal there blocked the VM in the native loop.
    GclMethod {
        name: "Show",
        arity: 1,
        target: GclMethodTarget::host("__ctrl_show"),
    },
    GclMethod {
        name: "ShowModal",
        arity: 1,
        target: GclMethodTarget::host("__dlg_showdialog"),
    },
    GclMethod {
        name: "Close",
        arity: 1,
        target: GclMethodTarget::host("closeForm"),
    },
];

const ADD_METHODS: &[GclMethod] = &[GclMethod {
    name: "Add",
    arity: 2,
    target: GclMethodTarget::host(gui::HOST_FN_ADD_CHILD),
}];

const EMPTY_METHODS: &[GclMethod] = &[];

pub fn gcl_classes() -> &'static [GclClass] {
    &CLASSES
}

macro_rules! widget_class {
    ($name:literal, $parent:literal, $host:literal, $props:expr, $methods:expr) => {
        GclClass {
            name: $name,
            parent: Some($parent),
            properties: $props,
            methods: $methods,
            ctor_arity: 1,
            widget_host_fn: Some($host),
        }
    };
}

static CLASSES: &[GclClass] = &[
    GclClass {
        name: "TObject",
        parent: None,
        properties: &[],
        methods: EMPTY_METHODS,
        ctor_arity: 0,
        widget_host_fn: None,
    },
    GclClass {
        name: "TComponent",
        parent: Some("TObject"),
        properties: &["Name", "Tag", "Owner"],
        methods: ADD_METHODS,
        ctor_arity: 1,
        widget_host_fn: None,
    },
    GclClass {
        name: "TControl",
        parent: Some("TComponent"),
        properties: CONTROL_PROPERTIES,
        methods: EMPTY_METHODS,
        ctor_arity: 1,
        widget_host_fn: None,
    },
    GclClass {
        name: "TWinControl",
        parent: Some("TControl"),
        properties: &["Controls"],
        methods: ADD_METHODS,
        ctor_arity: 1,
        widget_host_fn: None,
    },
    widget_class!(
        "TForm",
        "TWinControl",
        "new_Form",
        FORM_PROPERTIES,
        SHOW_METHODS
    ),
    widget_class!("TButton", "TWinControl", "new_Button", &[], EMPTY_METHODS),
    widget_class!("TLabel", "TControl", "new_Label", &[], EMPTY_METHODS),
    widget_class!(
        "TEdit",
        "TWinControl",
        "new_TextBox",
        TEXT_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TMemo",
        "TWinControl",
        "new_RichTextBox",
        TEXT_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TCheckBox",
        "TWinControl",
        "new_CheckBox",
        CHECK_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TRadioButton",
        "TWinControl",
        "new_RadioButton",
        CHECK_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TRadioGroup",
        "TWinControl",
        "new_GroupBox",
        LIST_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TComboBox",
        "TWinControl",
        "new_ComboBox",
        LIST_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TListBox",
        "TWinControl",
        "new_ListBox",
        LIST_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TGroupBox",
        "TWinControl",
        "new_GroupBox",
        &[],
        EMPTY_METHODS
    ),
    widget_class!("TPanel", "TWinControl", "new_Panel", &[], EMPTY_METHODS),
    widget_class!("TImage", "TControl", "new_PictureBox", &[], EMPTY_METHODS),
    widget_class!(
        "TShape",
        "TControl",
        "new_Panel",
        &["Shape", "Brush", "Pen"],
        EMPTY_METHODS
    ),
    widget_class!(
        "TBevel",
        "TControl",
        "new_Panel",
        &["Shape", "Style"],
        EMPTY_METHODS
    ),
    widget_class!("TSplitter", "TControl", "new_Panel", &[], EMPTY_METHODS),
    widget_class!(
        "TPageControl",
        "TWinControl",
        "new_TabControl",
        PAGE_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TTabSheet",
        "TWinControl",
        "new_TabPage",
        PAGE_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TTimer",
        "TComponent",
        "new_Timer",
        &["Interval", "Enabled", "OnTimer"],
        EMPTY_METHODS
    ),
    widget_class!(
        "TStringGrid",
        "TWinControl",
        "new_DataGridView",
        GRID_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TTrackBar",
        "TWinControl",
        "new_TrackBar",
        RANGE_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TProgressBar",
        "TWinControl",
        "new_ProgressBar",
        RANGE_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TSpinEdit",
        "TWinControl",
        "new_NumericUpDown",
        SPIN_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TListView",
        "TWinControl",
        "new_ListView",
        &["ViewStyle", "Items"],
        EMPTY_METHODS
    ),
    widget_class!(
        "TTreeView",
        "TWinControl",
        "new_TreeView",
        &["Items"],
        EMPTY_METHODS
    ),
    widget_class!(
        "TColorDialog",
        "TComponent",
        "new_ColorDialog",
        &["Color"],
        EMPTY_METHODS
    ),
    widget_class!(
        "TMainMenu",
        "TComponent",
        "new_MenuStrip",
        MENU_PROPERTIES,
        ADD_METHODS
    ),
    widget_class!(
        "TPopupMenu",
        "TComponent",
        "new_ContextMenuStrip",
        MENU_PROPERTIES,
        ADD_METHODS
    ),
    widget_class!(
        "TMenuItem",
        "TComponent",
        "new_MenuStrip",
        MENU_PROPERTIES,
        ADD_METHODS
    ),
];
