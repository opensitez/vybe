//! GCL: Pascal GUI component library surface.
//!
//! This is a Pascal-shaped adapter for VCL/LCL-style source. It lowers
//! Pascal classes such as `TForm` and `TButton` onto the DOM through `web:*`,
//! without adding Pascal-specific host functions.

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
    /// A shared emit in `primitives/gui.rs`, named rather than called through a
    /// host function — the same way a property role binds.
    Common { emit: &'static str },
}

impl GclMethodTarget {
    /// A shared GUI emit, named by ROLE. The only kind of target a GCL method
    /// has: a frontend contributes DECLARED DATA, and every GUI emit lives in
    /// `primitives/gui.rs` (guiplan.md, "The rules").
    pub const fn common(emit: &'static str) -> Self {
        GclMethodTarget::Common { emit }
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
    "Menu",
    "MainMenu",
    "PopupMenu",
    "ClientWidth",
    "ClientHeight",
    "OnCreate",
    "OnClose",
];
const TEXT_PROPERTIES: &[&str] = &["PasswordChar", "ReadOnly", "MaxLength", "Alignment"];

/// The classes VCL gives an `Alignment` to **and whose text we draw**.
///
/// It is NOT a `TControl` property — a `TButton` has no `Alignment` — so
/// declaring it on the common list would answer a question Delphi itself
/// rejects. `TLabel` has no property list of its own and gets this one;
/// `TEdit` and `TMemo` fold it into theirs.
///
/// `TPanel` is deliberately absent even though VCL declares `Alignment` on it.
/// A panel's alignment positions its CAPTION, and a `div` draws no text at all
/// — so declaring it would put a value in the store that answers through
/// `style_properties` and paints nothing. A declared property with no effect is
/// worse than an undeclared one: it reads back correctly and lies.
const ALIGNED_PROPERTIES: &[&str] = &["Alignment"];
/// A memo is an edit that also answers questions about its LINES.
///
/// `Count` is written `Lines.Count`, and `Lines` yields the receiver (see
/// `SELF_MEMBERS`), so the count is asked of the memo itself — which is why it
/// is declared here rather than on some separate strings object.
const MEMO_PROPERTIES: &[&str] = &[
    "PasswordChar",
    "ReadOnly",
    "MaxLength",
    "Count",
    // `ScrollBars` IS CSS `overflow` — the constants it takes are declared in
    // the Pascal profile as the CSS keywords, so the value arrives ready to
    // use. Undeclared, the whole assignment was invisible.
    "ScrollBars",
    "Alignment",
];
const CHECK_PROPERTIES: &[&str] = &["Checked", "State"];
const LIST_PROPERTIES: &[&str] = &["Items", "ItemIndex", "Sorted"];
const GRID_PROPERTIES: &[&str] = &["ColCount", "RowCount", "FixedCols", "FixedRows", "Cells"];
const RANGE_PROPERTIES: &[&str] = &["Min", "Max", "Position", "Step"];
const SPIN_PROPERTIES: &[&str] = &["MinValue", "MaxValue", "Value", "Increment"];
const PAGE_PROPERTIES: &[&str] = &["ActivePage", "PageIndex", "PageControl"];
/// Common-dialog properties, shared by the file dialogs the same way WinForms
/// hangs them off `FileDialog`. Declaring them is what makes
/// `D.Filter := '...'` a property write instead of a call into `undefined`.
const FILE_DIALOG_PROPERTIES: &[&str] = &[
    "FileName",
    "Filter",
    "FilterIndex",
    "DefaultExt",
    "InitialDir",
    "Title",
    "Options",
];
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
        target: GclMethodTarget::common("gui.ctrl.show_form"),
    },
    GclMethod {
        name: "ShowModal",
        arity: 1,
        target: GclMethodTarget::common("gui.ctrl.show_modal"),
    },
    GclMethod {
        name: "Close",
        arity: 1,
        target: GclMethodTarget::common("gui.ctrl.close"),
    },
];

const ADD_METHODS: &[GclMethod] = &[GclMethod {
    name: "Add",
    arity: 2,
    target: GclMethodTarget::Common {
        emit: gui::APPEND_CHILD_EMIT,
    },
}];

/// `SetFocus`, on every control that can take focus.
///
/// The verb IS `HTMLElement.focus()` — one of the few toolkit methods the web
/// platform names outright — so it needs no emit of its own: `gui.ctrl.<verb>`
/// imports `web:html`'s function of that name, and `focus(doc, node)` is
/// already the two-argument shape that lowering emits.
///
/// Declared on `TWinControl` rather than `TControl` because that IS the VCL's
/// own line: a `TControl` has no window handle and cannot be focused. `TLabel`
/// descends from `TControl` and correctly does not inherit this.
const WIN_CONTROL_METHODS: &[GclMethod] = &[
    GclMethod {
        name: "Add",
        arity: 2,
        target: GclMethodTarget::Common {
            emit: gui::APPEND_CHILD_EMIT,
        },
    },
    GclMethod {
        name: "SetFocus",
        arity: 1,
        target: GclMethodTarget::Common {
            emit: "gui.ctrl.focus",
        },
    },
];

/// `Add` for a list whose entries are STRINGS — `FList.Items.Add('alpha')`.
///
/// The same spelling as `ADD_METHODS` and a different operation, which is why
/// each class declares which one it means: a container is handed a control it
/// did not make, a list is handed text and makes the `<option>` itself. The
/// call site cannot tell them apart, and guessing from the argument's type
/// would be a guess.
const ITEM_METHODS: &[GclMethod] = &[
    GclMethod {
        name: "Add",
        arity: 2,
        target: GclMethodTarget::Common {
            emit: gui::APPEND_ITEM_EMIT,
        },
    },
    // `Items.Clear` — `select.length = 0`, which `web:html` already exposes as
    // `clearItems(doc, node)`. The same spelling on a TEdit or TMemo means
    // something else entirely (empty the TEXT, not the option list), which is
    // why it is declared here on the list classes and not on `TWinControl`.
    GclMethod {
        name: "Clear",
        arity: 1,
        target: GclMethodTarget::Common {
            emit: "gui.ctrl.clearItems",
        },
    },
    // `Items.Delete(i)` — `select.remove(i)`. Delphi's `TStrings` spells
    // removal `Delete`; the DOM spells it `remove`. One operation.
    GclMethod {
        name: "Delete",
        arity: 2,
        target: GclMethodTarget::Common {
            emit: gui::REMOVE_ITEM_EMIT,
        },
    },
];

/// `Clear` for a control whose contents are TEXT — `FInput.Clear`.
///
/// The list classes declare the same spelling against `clearItems` (see
/// `ITEM_METHODS`); a text field has no option list to empty, its `Clear` is
/// `value = ""`. Nothing at the call site distinguishes them, so each class
/// says which it means.
const TEXT_METHODS: &[GclMethod] = &[GclMethod {
    name: "Clear",
    arity: 1,
    target: GclMethodTarget::Common {
        emit: "gui.ctrl.clear",
    },
}];

/// Members that ARE the control, as `(class, member, reads back as)`.
///
/// VCL wraps a control's contents in a helper object — `TMainMenu.Items` is a
/// `TMenuItem`, `TMemo.Lines` is a `TStrings` — but in the document the element
/// already IS that container. So the member yields the receiver and nothing is
/// allocated, and the declared return type is what the NEXT hop resolves
/// against: `M.Items.Add(x)` looks `Add` up on `TMenuItem`, `FMemo.Lines.Text`
/// looks `Text` up on `TMemo`. Without the return type the chain resolved
/// against nothing and called `undefined` (menus), or — worse — silently wrote
/// to a property no element had (`Lines.Text`, measured: the textarea stayed
/// empty and nothing errored).
///
/// The menu entries declare `TMenuItem` while the getter hands back the
/// receiver, so for `TMainMenu` the compiler believes `TMenuItem` and the
/// runtime holds the `TMainMenu` element. That is sound only because a menu and
/// an item are the SAME element here — both are `menu`, with the same member
/// set, and `MenuStrip` is both the bar and the submenu a bar item opens. It is
/// a deliberate alias, not a coincidence to preserve blindly: give a menu item
/// its own tag and this line stops being true. (HTML would wrap each item in an
/// `<li>`; nothing renders the wrapper, so nothing here allocates one.)
///
/// A list's `Items` is the same shape for the same reason — `<select>` and
/// `<ul>` ARE their option list, there is no second object — but it reads back
/// as the list's OWN class, because what `Add` means differs: a menu is handed
/// a control, a list is handed text (`ADD_METHODS` vs `ITEM_METHODS`).
pub const SELF_MEMBERS: &[(&str, &str, &str)] = &[
    ("TMainMenu", "Items", "TMenuItem"),
    ("TPopupMenu", "Items", "TMenuItem"),
    ("TMenuItem", "Items", "TMenuItem"),
    ("TMemo", "Lines", "TMemo"),
    ("TListBox", "Items", "TListBox"),
    ("TComboBox", "Items", "TComboBox"),
    ("TRadioGroup", "Items", "TRadioGroup"),
];

/// Classes whose option list is readable and writable BY INDEX — `Items[i]`.
///
/// Registered as the instance property `Item`, which is the name the index site
/// already looks for: a type declaring `Item` with a common emit in each
/// direction makes `x[i]` lower to that emit. .NET spells the same thing
/// `this[int]` and Delphi spells it `TStrings.Strings[i]`, its default indexed
/// property — one concept, and no new compiler mechanism to reach it.
///
/// Both directions are required. The index site asks for the pair and declines
/// the branch unless it gets both, which is what stops a type offering a
/// readable index and a write that goes nowhere.
///
/// `TRadioGroup` is deliberately NOT here even though it declares `Items`. It
/// is a `<fieldset>`, not a `<select>`, and its options are child
/// `<input type=radio>` elements rather than an option list — so the widget has
/// no item to answer with and `Items[0]` reads `""`. Measured, not assumed.
/// Listing it would have bought an empty string in place of a caption, which is
/// the same silent-wrong-answer this whole change set exists to remove.
pub const INDEXED_ITEM_CLASSES: &[&str] = &["TListBox", "TComboBox"];

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
        methods: WIN_CONTROL_METHODS,
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
    widget_class!("TLabel", "TControl", "label", ALIGNED_PROPERTIES, EMPTY_METHODS),
    widget_class!(
        "TEdit",
        "TWinControl",
        "input:text",
        TEXT_PROPERTIES,
        TEXT_METHODS
    ),
    widget_class!(
        "TMemo",
        "TWinControl",
        "textarea",
        MEMO_PROPERTIES,
        TEXT_METHODS
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
        ITEM_METHODS
    ),
    widget_class!(
        "TComboBox",
        "TWinControl",
        "select",
        LIST_PROPERTIES,
        ITEM_METHODS
    ),
    widget_class!(
        "TListBox",
        "TWinControl",
        // HTML's list box IS `<select size="N">` — the spec's own rule is that
        // a size above 1 renders as a list and 1 (or absent) as a dropdown, so
        // this and `TComboBox` are the same element differing by one attribute,
        // exactly as they are in the VCL.
        //
        // It was `<ul>`, which RENDERS as a list box and cannot answer a single
        // question about one: no `selectedIndex`, no indexed option text, no
        // `remove(i)`. `ItemIndex`, `Items[i]` and `Items.Delete` all read null
        // in silence. Adding those to `<ul>` in `web:html` would have been
        // inventing surface the spec does not have — the element was wrong, not
        // the platform.
        //
        // ⛔ `size` is a content ATTRIBUTE (`@size=6`), not the input type. The
        // `tag:type` slot becomes `type="…"`, which `<select>` does not have —
        // so spelling it there produced `<select type="6">`, failed the
        // engine's `size > 1` test, and rendered the dropdown this exists to
        // stop being.
        "select;@size=6",
        LIST_PROPERTIES,
        ITEM_METHODS
    ),
    widget_class!("TGroupBox", "TWinControl", "fieldset", &[], EMPTY_METHODS),
    widget_class!("TPanel", "TWinControl", "div", &[], EMPTY_METHODS),
    // A `TImage` IS a drawing surface, and HTML spells that `<canvas>` — the
    // same element WinForms' `PictureBox` resolves to, and the only node that
    // owns a recording for `Canvas`/`Graphics` calls to land on.
    widget_class!("TImage", "TControl", "canvas", &[], EMPTY_METHODS),
    // `TShape` and `TBevel` draw one primitive — a rectangle, ellipse or a
    // bevelled line. That is a drawing surface too, not a control with
    // behaviour, so both are the element a drawing goes on.
    widget_class!(
        "TShape",
        "TControl",
        "canvas",
        &["Shape", "Brush", "Pen"],
        EMPTY_METHODS
    ),
    widget_class!(
        "TBevel",
        "TControl",
        "canvas",
        &["Shape", "Style"],
        EMPTY_METHODS
    ),
    // The bare drag bar between two panes — a box with a resize cursor.
    widget_class!(
        "TSplitter",
        "TControl",
        "div;background-color:#c8c8c8;cursor:col-resize",
        &[],
        EMPTY_METHODS
    ),
    // A tab strip over its pages, and a page. Ordinary containers; which page
    // shows is a `display` question, not a different element.
    widget_class!(
        "TPageControl",
        "TWinControl",
        "div;display:flex;flex-direction:column",
        PAGE_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TTabSheet",
        "TWinControl",
        "section",
        PAGE_PROPERTIES,
        EMPTY_METHODS
    ),
    // A `components` member: present and scriptable, never painted, which is
    // what `display: none` says in HTML's own vocabulary.
    //
    // ⚠ If a timer still does not tick after this, that is the known dead
    // `GuiState` defect and not the element — Cairn flagged it directly.
    widget_class!(
        "TTimer",
        "TComponent",
        "div;display:none",
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
    widget_class!("TTreeView", "TWinControl", "ul", &["Items"], EMPTY_METHODS),
    // `<input type=color>` IS a colour chooser, hidden until `Execute` opens
    // it — the same element WinForms' `ColorDialog` resolves to.
    widget_class!(
        "TColorDialog",
        "TComponent",
        "input:color;display:none",
        &["Color"],
        EMPTY_METHODS
    ),
    // The file dialogs. DECLARED so construction and every property work; the
    // picker itself is `Execute`, which VCL defines as a BLOCKING modal
    // returning Boolean and the web has no synchronous equivalent of. That is
    // a scope decision, not something to fake — a stub returning True would
    // send the program down the "user chose a file" branch with no file.
    widget_class!(
        "TOpenDialog",
        "TComponent",
        "input:file;display:none",
        FILE_DIALOG_PROPERTIES,
        EMPTY_METHODS
    ),
    widget_class!(
        "TSaveDialog",
        "TComponent",
        "input:file;display:none",
        FILE_DIALOG_PROPERTIES,
        EMPTY_METHODS
    ),
    // `menu` is a real HTML element and the document already knows it. A
    // pseudo-tag no `control_kind` arm matches renders as a 120x20 LABEL at the
    // origin. A menu ITEM is the same tag on purpose: see `SELF_MEMBERS`.
    widget_class!(
        "TMainMenu",
        "TComponent",
        "menu",
        MENU_PROPERTIES,
        ADD_METHODS
    ),
    widget_class!(
        "TPopupMenu",
        "TComponent",
        "menu",
        MENU_PROPERTIES,
        ADD_METHODS
    ),
    widget_class!(
        "TMenuItem",
        "TComponent",
        "menu",
        MENU_PROPERTIES,
        ADD_METHODS
    ),
];
