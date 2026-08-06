//! GCL: Pascal GUI component library surface.
//!
//! This is a Pascal-shaped adapter for VCL/LCL-style source. It lowers
//! Pascal classes such as `TForm` and `TButton` to the existing generic
//! `vybe:gui` host imports without adding Pascal-specific host functions.

pub mod builder;

use vybe_compiler::primitives::gui;

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
    // Delphi unit scope names are a namespace prefix — `Vcl.Forms` IS `Forms`.
    // Match on the last segment so both spellings answer the same.
    let path = path.rsplit('.').next().unwrap_or(path);
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
        "body",
        FORM_PROPERTIES,
        SHOW_METHODS
    ),
    widget_class!("TButton", "TWinControl", "button", &[], EMPTY_METHODS),
    widget_class!("TLabel", "TControl", "label", &[], EMPTY_METHODS),
    widget_class!(
        "TEdit",
        "TWinControl",
        "input:text",
        TEXT_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TMemo",
        "TWinControl",
        "textarea",
        TEXT_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TCheckBox",
        "TWinControl",
        "input:checkbox",
        CHECK_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TRadioButton",
        "TWinControl",
        "input:radio",
        CHECK_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TRadioGroup",
        "TWinControl",
        "fieldset",
        LIST_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TComboBox",
        "TWinControl",
        "select",
        LIST_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TListBox",
        "TWinControl",
        "ul",
        LIST_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TGroupBox",
        "TWinControl",
        "fieldset",
        &[],
        EMPTY_METHODS
    ),
    widget_class!("TPanel", "TWinControl", "div", &[], EMPTY_METHODS),
    widget_class!("TImage", "TControl", "vybe-picturebox", &[], EMPTY_METHODS),
    widget_class!(
        "TShape",
        "TControl",
        "vybe-shape",
        &["Shape", "Brush", "Pen"],
        EMPTY_METHODS
    ),
    widget_class!(
        "TBevel",
        "TControl",
        "vybe-bevel",
        &["Shape", "Style"],
        EMPTY_METHODS
    ),
    widget_class!("TSplitter", "TControl", "vybe-splitter", &[], EMPTY_METHODS),
    widget_class!(
        "TPageControl",
        "TWinControl",
        "vybe-tabcontrol",
        PAGE_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TTabSheet",
        "TWinControl",
        "vybe-tabpage",
        PAGE_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TTimer",
        "TComponent",
        "vybe-timer",
        &["Interval", "Enabled", "OnTimer"],
        EMPTY_METHODS
    ),
    widget_class!(
        "TStringGrid",
        "TWinControl",
        "table",
        GRID_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TTrackBar",
        "TWinControl",
        "input:range",
        RANGE_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TProgressBar",
        "TWinControl",
        "progress",
        RANGE_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TSpinEdit",
        "TWinControl",
        "input:number",
        SPIN_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TListView",
        "TWinControl",
        "ul",
        &["ViewStyle", "Items"],
        EMPTY_METHODS
    ),
    widget_class!(
        "TTreeView",
        "TWinControl",
        "ul",
        &["Items"],
        EMPTY_METHODS
    ),
    widget_class!(
        "TColorDialog",
        "TComponent",
        "vybe-colordialog",
        &["Color"],
        EMPTY_METHODS
    ),
    widget_class!(
        "TMainMenu",
        "TComponent",
        "vybe-menustrip",
        MENU_PROPERTIES,
        ADD_METHODS
    ),
    widget_class!(
        "TPopupMenu",
        "TComponent",
        "vybe-menustrip",
        MENU_PROPERTIES,
        ADD_METHODS
    ),
    widget_class!(
        "TMenuItem",
        "TComponent",
        "vybe-menustrip",
        MENU_PROPERTIES,
        ADD_METHODS
    ),
];
