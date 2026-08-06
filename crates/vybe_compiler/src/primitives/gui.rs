//! Canonical GUI emit layer — shared by every language frontend.
//!
//! This module is the **single source of truth** for what a GUI control IS
//! and how to emit bytecode that creates, configures, and wires it. Every
//! framework frontend (`dotnet.rs` for WinForms, future `maui.rs`, `flutter.rs`,
//! `tkinter.rs`, etc.) delegates here for the actual emit. The frontends only
//! deal with surface naming and convention; the canonical button/textbox/etc.
//! and the host call vocabulary live here.
//!
//! This mirrors how `compiler_common::loops` powers every language's `for` /
//! `foreach` / `while`, and how `compiler_common::collections` powers
//! `arr.length` / `len(arr)` / `Length(arr)` from a single emit path.
//!
//! ## Architecture
//!
//! ```text
//! VB walker  ───┐
//! C# walker  ───┤   .NET surface         vybe:gui::*
//! F# walker  ───┴──> dotnet.rs ──┐
//!                                │
//! Dart walker ─────> flutter.rs ─┼──> compiler_common::gui ──> host fn
//!                                │
//! Python walker ───> tkinter.rs ─┘
//! ```
//!
//! All frontends produce the SAME bytecode for the same canonical operation.
//! Switching the host's GUI backend (or running on a non-Vybe VM with a
//! different GUI binding) requires no compiler changes.

use std::sync::Arc;
use super::Compiler;
use vybe_ast::Expression;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

// ─── Canonical control type registry ─────────────────────────────────────────
//
// These are the GUI primitives that exist in every modern UI framework
// (WinForms, MAUI, WPF, Tkinter, Flutter, GTK, Qt, web DOM, etc.). The names
// are the canonical PascalCase form. Frontends map their framework-specific
// surface names to these.

/// Returns the canonical PascalCase control name if `name` (case-insensitive)
/// matches a known GUI control type. Returns empty string otherwise.
///
/// Frontends use this to test whether an identifier in source code refers to
/// a canonical GUI control. For example, .NET's `Button`, MAUI's `Button`,
/// and Flutter's `ElevatedButton` would all map to the canonical "Button".
pub fn canonical_control_name(name: &str) -> String {
    match name.to_ascii_lowercase().as_str() {
        // ── Buttons & inputs ──
        "button" => "Button",
        "checkbox" => "CheckBox",
        "radiobutton" => "RadioButton",
        "togglebutton" => "ToggleButton",
        "linkbutton" | "linklabel" => "LinkLabel",

        // ── Text ──
        "label" => "Label",
        "textbox" | "entry" | "textfield" => "TextBox",
        "richtextbox" => "RichTextBox",
        "maskedtextbox" => "MaskedTextBox",

        // ── Selection ──
        "combobox" | "dropdown" => "ComboBox",
        "listbox" => "ListBox",
        "listview" => "ListView",
        "treeview" => "TreeView",

        // ── Containers ──
        "panel" | "container" => "Panel",
        "groupbox" => "GroupBox",
        "tabcontrol" | "tabbedpane" => "TabControl",
        "tabpage" => "TabPage",
        "splitcontainer" => "SplitContainer",
        "flowlayoutpanel" => "FlowLayoutPanel",
        "tablelayoutpanel" => "TableLayoutPanel",

        // ── Date / time / numeric ──
        "datetimepicker" => "DateTimePicker",
        "monthcalendar" => "MonthCalendar",
        "numericupdown" => "NumericUpDown",

        // ── Progress / indicators ──
        "progressbar" => "ProgressBar",
        "trackbar" | "slider" => "TrackBar",

        // ── Images / media ──
        "picturebox" | "image" => "PictureBox",
        "webbrowser" => "WebBrowser",

        // ── Data / grids ──
        "datagridview" | "datagrid" => "DataGridView",

        // ── Menus / strips ──
        "menustrip" | "menubar" => "MenuStrip",
        "toolstrip" | "toolbar" => "ToolStrip",
        "statusstrip" | "statusbar" => "StatusStrip",
        "contextmenustrip" | "contextmenu" => "ContextMenuStrip",

        // ── Scrollbars ──
        "hscrollbar" => "HScrollBar",
        "vscrollbar" => "VScrollBar",

        // ── Dialogs ──
        "openfiledialog" => "OpenFileDialog",
        "savefiledialog" => "SaveFileDialog",
        "folderbrowserdialog" => "FolderBrowserDialog",
        "colordialog" => "ColorDialog",
        "fontdialog" => "FontDialog",

        // ── Non-visual / lifecycle ──
        "timer" => "Timer",
        "tooltip" => "ToolTip",
        "imagelist" => "ImageList",
        "notifyicon" => "NotifyIcon",
        "errorprovider" => "ErrorProvider",
        "helpprovider" => "HelpProvider",
        "backgroundworker" => "BackgroundWorker",
        "bindingsource" => "BindingSource",
        "bindingnavigator" => "BindingNavigator",

        // ── Forms ──
        "form" | "window" => "Form",

        _ => return String::new() }
    .to_string()
}

/// Returns true if `name` is a recognized canonical GUI control type
/// (case-insensitive).
pub fn is_control_type(name: &str) -> bool {
    !canonical_control_name(name).is_empty()
}

// ─── Host function naming ────────────────────────────────────────────────────
//
// These are the canonical host functions every GUI backend (Vybe's native
// renderer, future Tauri/Electron/web backends) must implement. Frontends emit
// bytecode that calls these names; the host registry resolves them.

/// Build the host fn name for "create a new control of this type".
/// e.g. canonical "Button" → "new_Button".
///
/// A top-level `Form` uses `newForm`, which materialises a real form window
/// (proper default size + `Controls`/`components` collections), rather than the
/// generic `new_<Type>` control factory (which defaults to a tiny child-control
/// size).
pub fn host_fn_new_control(canonical: &str) -> String {
    if canonical == "Form" {
        return "newForm".to_string();
    }
    format!("new_{}", canonical)
}

/// Host fn name for "set a property on a control object".
/// Stack at call site: [obj, prop_name, value]
pub const HOST_FN_SET_PROPERTY: &str = "controlSetProperty";

/// Host fn name for "get a property from a control object".
/// Stack at call site: [obj, prop_name]
pub const HOST_FN_GET_PROPERTY: &str = "controlGetProperty";

/// Host fn name for "create a controls collection bound to this owner".
/// Stack at call site: [owner]
pub const HOST_FN_NEW_CONTROLS_COLLECTION: &str = "newControlsCollection";

/// Host fn name for "create a components collection bound to this owner".
/// Stack at call site: [owner]
pub const HOST_FN_NEW_COMPONENTS_COLLECTION: &str = "newComponentsCollection";

/// Host fn name for "register an event handler on a control".
/// Stack at call site: [control_name_string, event_name, handler_fn_ref]
pub const HOST_FN_BIND_EVENT: &str = "onEvent";

/// Host fn name for "remove an event handler from a control".
pub const HOST_FN_UNBIND_EVENT: &str = "removeEvent";

/// Host fn name for "add a control as a child of another control's
/// .Controls collection".
/// Stack at call site: [parent, child]
pub const HOST_FN_ADD_CHILD: &str = "controlsAdd";

/// Host fn name for "run the application event loop with this form".
pub const HOST_FN_RUN_APPLICATION: &str = "runApplication";

/// Host fn name for "exit the application".
pub const HOST_FN_APP_EXIT: &str = "appExit";

/// Host fn name for "fire a custom event on the current control/form".
/// Stack at call site: [arg0, arg1, ..., event_name_string]
pub const HOST_FN_RAISE_EVENT: &str = "raiseEvent";

pub const GUI_MODULE: &str = "vybe:gui";

/// WHATWG DOM — where a control is actually created.
pub const DOM_MODULE: &str = "web:dom";
/// WHATWG HTML — `document`, and the element IDL properties.
pub const DOCUMENT_MODULE: &str = "web:html";
pub const HOST_FN_CREATE_ELEMENT: &str = "createElement";
/// `window.document` of the current browsing context.
pub const HOST_FN_ACTIVE_DOCUMENT: &str = "activeDocument";
/// CSSOM — `element.style`.
pub const CSSOM_MODULE: &str = "web:cssom";

// ─── Component Model Registration ────────────────────────────────────────────

/// Register all `vybe:gui` host functions as component module exports.
/// This is called by the primitives/Linker to populate the component descriptor
/// so all languages automatically get GUI functions without per-language
/// profile duplication (similar to WASI registration).
pub fn gui_component_exports() -> Vec<vybe_runtime::component_model::ComponentExport> {
    use vybe_runtime::component::{FuncSig, ValType};
    use vybe_runtime::component_model::{ComponentExport, ComponentItemKind};

    vec![
        // Control creation
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "createForm".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "createForm".to_string(),
                params: vec![ValType::String],
                results: vec![ValType::Any] }) },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "addControl".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "addControl".to_string(),
                params: vec![],
                results: vec![] }) },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "setProperty".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "setProperty".to_string(),
                params: vec![],
                results: vec![] }) },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "getProperty".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "getProperty".to_string(),
                params: vec![],
                results: vec![ValType::Any] }) },
        // Event handling
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "onEvent".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "onEvent".to_string(),
                params: vec![],
                results: vec![] }) },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "removeEvent".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "removeEvent".to_string(),
                params: vec![],
                results: vec![] }) },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "raiseEvent".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "raiseEvent".to_string(),
                params: vec![],
                results: vec![] }) },
        // Collections
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "newControlsCollection".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "newControlsCollection".to_string(),
                params: vec![],
                results: vec![ValType::Any] }) },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "newComponentsCollection".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "newComponentsCollection".to_string(),
                params: vec![],
                results: vec![ValType::Any] }) },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "controlsAdd".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "controlsAdd".to_string(),
                params: vec![],
                results: vec![] }) },
        // Application
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "runApplication".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "runApplication".to_string(),
                params: vec![],
                results: vec![] }) },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "appExit".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "appExit".to_string(),
                params: vec![],
                results: vec![] }) },
        // Form lifecycle
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "showForm".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "showForm".to_string(),
                params: vec![],
                results: vec![] }) },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "closeForm".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "closeForm".to_string(),
                params: vec![],
                results: vec![] }) },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "showFormDialog".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "showFormDialog".to_string(),
                params: vec![],
                results: vec![] }) },
        // Dialog
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "msgBox".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "msgBox".to_string(),
                params: vec![],
                results: vec![ValType::I32] }) },
    ]
}

// ─── Emit helpers ────────────────────────────────────────────────────────────
//
// Canonical patterns. Every language frontend uses these directly or via a
// framework-specific resolver (`dotnet.rs`, etc.) that calls these.
//
// All emit functions are pure WASM bytecode + standard host imports — no
// custom opcodes, no language-specific knowledge. They define the calling
// convention so call sites are uniform across compilers.
//
// IMPORTANT: imports must be registered against a single chunk (typically the
// script chunk, chunk[0]) so the VM's import_table resolution works. Callers
// pass a pre-resolved `import_idx` obtained via their compiler's `import()`
// helper (which delegates to `chunks[0].add_import`). This keeps gui.rs
// chunk-agnostic — it doesn't need to know which chunk it's emitting into.

// ─── Lowering a control to the web platform ──────────────────────────────

// ─── The one GUI voice ───────────────────────────────────────────────────
//
// A control property is a ROLE, not a spelling. Pascal's `Caption`, .NET's
// `Text` and Flutter's `child` are the same role, and each frontend lowers
// its own word to it — the spelling stops at the frontend, exactly as a VB
// constructor is recognised as a constructor because it FILLS a role.
//
// So nothing here speaks any language. The role IS the WHATWG IDL property
// name, because the DOM is what we talk to and its vocabulary is already
// standard; this maps role → the DOM operation that performs it, once.

/// `gui.prop_set.<role>` / `gui.prop_get.<role>` — the emit a platform
/// declares instead of naming a host function.
pub const PROP_SET_EMIT: &str = "gui.prop_set.";
pub const PROP_GET_EMIT: &str = "gui.prop_get.";

/// The DOM operation a property role IS. `(module, func, attribute-key)`.
///
/// The roles ARE `vybe:gui`'s canonical property names — the vocabulary every
/// language was already lowering to. Nothing was invented here; only the
/// TARGET changed, from a custom host function to a compliant DOM operation.
/// That is also why dotnet needs no mapping: it already emits these names.
///
/// Pascal never learns any of this. It calls with the same intent it always
/// had; `vybe_widgets` is HTML underneath, which is not its business.
///
/// A role with no IDL counterpart becomes an attribute — where unknown
/// properties belong on the web — so this stays the handful HTML treats
/// specially rather than a table that grows per control.
fn property_op(role: &str, setting: bool) -> (&'static str, &'static str, Option<&'static str>) {
    match role {
        // The widget resolves what "text" means for the control it is: a
        // `SetText` on a text field sets its value, on a label its caption.
        // So this needs no element test — the engine already knows.
        "text" | "caption" => (
            DOM_MODULE,
            if setting { "setTextContent" } else { "textContent" },
            None,
        ),
        "value" => (
            DOCUMENT_MODULE,
            if setting { "setValue" } else { "value" },
            None,
        ),
        "checked" | "ischecked" => (
            DOCUMENT_MODULE,
            if setting { "setChecked" } else { "checked" },
            None,
        ),
        // A control's `Name` IS the element id — what `getElementById` and
        // `<label for>` resolve. Not HTML's `name`, the submission key.
        "name" => (
            DOM_MODULE,
            if setting { "setAttribute" } else { "getAttribute" },
            Some("id"),
        ),
        // Boolean content attributes, INVERTED: true by PRESENCE, so
        // `Enabled := False` ADDS `disabled`. `toggleAttribute` is the DOM's
        // own add-or-remove.
        "enabled" => (
            DOM_MODULE,
            if setting { "toggleAttribute" } else { "getAttribute" },
            Some("disabled"),
        ),
        "visible" => (
            DOM_MODULE,
            if setting { "toggleAttribute" } else { "getAttribute" },
            Some("hidden"),
        ),
        "left" | "top" | "width" | "height" => (
            CSSOM_MODULE,
            if setting { "setStyleProperty" } else { "getStyleProperty" },
            Some(""),
        ),
        _ => (
            DOM_MODULE,
            if setting { "setAttribute" } else { "getAttribute" },
            Some(""),
        ) }
}

/// The event type an `on<type>` role registers, if it is one.
///
/// A ROLE, exactly like the property roles: every language spells its handler
/// slot `OnClick` / `onClick` / `Click`, and each lowers to the one DOM event
/// type. Nothing here is per-language and nothing is enumerated — HTML's own
/// convention is that an event handler IDL attribute is `on` + the type.
fn event_role_type(role: &str) -> Option<&str> {
    role.strip_prefix("on").filter(|type_name| !type_name.is_empty())
}

impl Compiler {
    /// Lower `gui.prop_get.<role>` — stack in `[control]`, out `[value]`.
    pub fn emit_gui_property_get(&mut self, role: &str, line: u32) {
        let (module, func, key) = property_op(role, false);
        let ctrl = self.define_local("__gui_prop_ctrl");
        self.emit_u16(Op::LOCAL_SET, ctrl);
        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
        self.chunk().emit_call(doc_idx, 0, line);
        self.emit_u16(Op::LOCAL_GET, ctrl);
        let argc = match key {
            Some(k) => {
                emit_string_const(self.chunk(), if k.is_empty() { role } else { k }, line);
                3
            }
            None => 2 };
        let idx = self.import(module, func);
        self.emit_host_call(idx, argc);
        // The DOM answers in the DOM's own terms; the role's value type is
        // what the caller asked for. Both conversions below are the exact
        // inverse of what `emit_gui_property_set` writes, so a round trip is
        // an identity — which is what `readback=` in the repros checks.
        match role {
            // `getAttribute` is null when ABSENT and "" when present, and
            // `disabled`/`hidden` are the INVERSE of the role. Absent means
            // enabled/visible, so the answer is "was it absent", as a real
            // boolean rather than the attribute's own text.
            "enabled" | "visible" => {
                self.emit(Op::REF_IS_NULL);
                self.chunk().emit_if_value(line);
                self.emit_const(Value::Bool(true));
                self.chunk().emit_else(line);
                self.emit_const(Value::Bool(false));
                self.chunk().emit_end(line);
            }
            // CSS geometry is TEXT with units (`"10px"`) — that is the spec,
            // and `vybe_widgets` is right to store it that way. A control's
            // `Left` is a number, so parse the unit back off here.
            "left" | "top" | "width" | "height" => {
                let parse_float = self.import("ecma:number", "parseFloat");
                self.emit_host_call(parse_float, 1);
            }
            _ => {}
        }
    }

    /// Lower `gui.prop_set.<role>` — stack in `[control, value]`, out `[_]`.
    pub fn emit_gui_property_set(&mut self, role: &str, line: u32) {
        // `OnClick := handler` IS `addEventListener("click", handler)`. HTML
        // spells the same thing as an `on<type>` IDL attribute, so the role
        // needs no translation table — the event type is whatever follows
        // `on`, and `addEventListener` takes any type string, which is what
        // lets `OnTimer` and `OnCreate` register alongside `click` without
        // inventing anything.
        if let Some(event_type) = event_role_type(role) {
            let value = self.define_local("__gui_event_handler");
            let ctrl = self.define_local("__gui_event_target");
            self.emit_u16(Op::LOCAL_SET, value);
            self.emit_u16(Op::LOCAL_SET, ctrl);
            let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
            self.chunk().emit_call(doc_idx, 0, line);
            self.emit_u16(Op::LOCAL_GET, ctrl);
            emit_string_const(self.chunk(), event_type, line);
            self.emit_u16(Op::LOCAL_GET, value);
            let idx = self.import(DOM_MODULE, "addEventListener");
            self.emit_host_call(idx, 4);
            return;
        }
        let (module, func, key) = property_op(role, true);
        let value = self.define_local("__gui_prop_value");
        let ctrl = self.define_local("__gui_prop_ctrl");
        self.emit_u16(Op::LOCAL_SET, value);
        self.emit_u16(Op::LOCAL_SET, ctrl);
        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
        self.chunk().emit_call(doc_idx, 0, line);
        self.emit_u16(Op::LOCAL_GET, ctrl);
        let argc = match key {
            Some(k) => {
                emit_string_const(self.chunk(), if k.is_empty() { role } else { k }, line);
                self.emit_u16(Op::LOCAL_GET, value);
                // `disabled`/`hidden` are the INVERSE of enabled/visible; the
                // frontend lowered to the ATTRIBUTE role, so negate here once.
                if matches!(role, "enabled" | "visible") {
                    self.chunk().emit_op(Op::I32_EQZ, line);
                }
                4
            }
            None => {
                self.emit_u16(Op::LOCAL_GET, value);
                3
            }
        };
        let idx = self.import(module, func);
        self.emit_host_call(idx, argc);
    }
}

/// The HTML element a control IS — tag, plus `type` for `<input>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ControlElement {
    pub tag: String,
    pub input_type: String,
}

impl ControlElement {
    /// Parse a platform's declaration of what its control is.
    ///
    /// A platform declares the ELEMENT (`"button"`, `"input:checkbox"`,
    /// `"body"`) because it owns the vocabulary — plib knows `TEdit` is a text
    /// input, and nothing in a shared crate should have to. That is the whole
    /// point: no per-language table lives here.
    fn parse(decl: &str) -> ControlElement {
        let (tag, input_type) = decl.split_once(':').unwrap_or((decl, ""));
        ControlElement {
            tag: tag.trim().to_ascii_lowercase(),
            input_type: input_type.trim().to_ascii_lowercase() }
    }

    /// Is this element FORM-ASSOCIATED — i.e. does it belong to
    /// `form.elements` and get submitted?
    ///
    /// HTML's own list, not a judgement call: button, fieldset, input,
    /// object, output, select, textarea. A `<img>`, `<div>`, `<ul>`,
    /// `<progress>` or `<table>` is NOT one — a `PictureBox` or `Panel` is a
    /// control in the toolkit sense but carries no submission identity, and
    /// `name` on it is non-conforming markup that submits nothing.
    pub fn is_form_associated(&self) -> bool {
        matches!(
            self.tag.as_str(),
            "button" | "fieldset" | "input" | "object" | "output" | "select" | "textarea"
        )
    }

    /// A control with no conforming HTML counterpart becomes a CUSTOM
    /// ELEMENT — `<vybe-picturebox>`, `<vybe-timer>`.
    ///
    /// This is valid HTML, not a fudge: a custom element name only has to
    /// contain a hyphen, and a real browser gives it behaviour through
    /// `customElements.define`. It beats both alternatives — a `<div>` says
    /// nothing about what the control is, and forcing a `PictureBox` into
    /// `<img>` claims image semantics it does not have (no `src`, no
    /// decoding, and it is not form-associated).
    ///
    /// Also where platforms that still name a `new_<Type>` factory land
    /// (flutter, next to migrate), so their widgets are at least named
    /// rather than anonymous boxes.
    fn custom(type_name: &str) -> ControlElement {
        let bare = type_name.rsplit(['.', ':']).next().unwrap_or(type_name);
        let bare = bare.trim_start_matches(['T', 't']).to_ascii_lowercase();
        ControlElement {
            tag: format!("vybe-{}", if bare.is_empty() { "control" } else { &bare }),
            input_type: String::new() }
    }
}

/// What the REGISTRY says this type's control is — the same authority
/// `is_framework_control_parent` consults, so the two can never disagree.
pub fn registered_control_element(
    type_scopes: &[String],
    type_name: &str,
) -> Option<ControlElement> {
    let spec = vybe_runtime::namespaces::lookup_type_ctor_spec(type_scopes, type_name)?;
    let decl = spec.control_fn?;
    Some(if decl.starts_with("new_") {
        ControlElement::custom(type_name)
    } else {
        ControlElement::parse(&decl)
    })
}

impl Compiler {
    /// The element a type IS, following a user class up to the control it
    /// derives from.
    ///
    /// `class TForm1 = class(TForm)` IS a form: a subclass of a control is a
    /// control, the same way `is_framework_control_parent` already treats the
    /// parent. Asking the registry alone answered None for every user-declared
    /// form, so `Self.OnCreate := handler` — a property write on the form
    /// itself — missed the DOM path that its own buttons took.
    pub fn control_element_for_type(&self, type_name: &str) -> Option<ControlElement> {
        if let Some(element) =
            registered_control_element(&self.profile.namespaces.type_scopes, type_name)
        {
            return Some(element);
        }
        // Walk the declared parents. Bounded by the chain itself, and a cycle
        // in it would already have broken construction long before here.
        let mut current = self.pending_class_parent(type_name);
        while let Some(parent) = current {
            if let Some(element) =
                registered_control_element(&self.profile.namespaces.type_scopes, &parent)
            {
                return Some(element);
            }
            current = self.pending_class_parent(&parent);
        }
        None
    }

    /// `control.<prop> = value` → the DOM.
    ///
    /// Property NAMES arrive already normalised by the frontend (Pascal's
    /// `Caption` and .NET's `Text` both reach here as `text`), so there is no
    /// per-language vocabulary in this file — only the web one.
    ///
    /// Anything without an IDL counterpart becomes `setAttribute`, which is
    /// where unknown properties belong on the web anyway; that keeps the match
    /// to the handful of properties HTML actually treats specially instead of
    /// a table that has to grow per control.
    ///
    /// Stack on entry: [value]. Stack on exit: empty.
    pub fn emit_control_property_set(
        &mut self,
        object: &Expression,
        type_name: &str,
        prop: &str,
        line: u32,
    ) -> Result<(), String> {
        let value_tmp = self.define_local("__ctrl_prop_value");
        self.emit_u16(Op::LOCAL_SET, value_tmp);

        // ONE role→DOM mapping, not two. This used to carry its own copy of
        // the `property_op` match, and the copies drifted the moment a role
        // was added to one of them: `OnClick` reached the DOM as an
        // `setAttribute("onclick", <closure>)` — a stringified function on an
        // attribute — because only the other path had learned that an
        // `on<type>` role registers a listener.
        let prop = prop.to_ascii_lowercase();
        self.compile_expr(object)?;
        self.emit_u16(Op::LOCAL_GET, value_tmp);
        self.emit_gui_property_set(&prop, line);
        self.emit(Op::DROP);

        // A designer `Name` is BOTH of HTML's two identifiers, which are not
        // the same thing: `id` is unique per document and is what
        // `getElementById` and `<label for>` resolve, while `name` is the
        // form-control submission key that `form.elements[…]` and
        // serialization read. A control that set only `id` would look right
        // and submit nothing, so set both.
        let form_associated = registered_control_element(
            &self.profile.namespaces.type_scopes,
            type_name,
        )
        .map(|e| e.is_form_associated())
        .unwrap_or(false);
        if prop == "name" && form_associated {
            let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
            self.chunk().emit_call(doc_idx, 0, line);
            self.compile_expr(object)?;
            emit_string_const(self.chunk(), "name", line);
            self.emit_u16(Op::LOCAL_GET, value_tmp);
            let idx = self.import(DOM_MODULE, "setAttribute");
            self.emit_host_call(idx, 4);
            self.emit(Op::DROP);
        }
        Ok(())
    }

    /// `control.<prop>` → the DOM. Stack on exit: [value].
    pub fn emit_control_property_get(
        &mut self,
        object: &Expression,
        prop: &str,
        line: u32,
    ) -> Result<(), String> {
        let prop = prop.to_ascii_lowercase();
        let (module, func, key): (&str, &str, Option<&str>) = match prop.as_str() {
            "text" | "caption" => (DOM_MODULE, "textContent", None),
            "value" => (DOCUMENT_MODULE, "value", None),
            "checked" | "ischecked" => (DOCUMENT_MODULE, "checked", None),
            "name" => (DOM_MODULE, "getAttribute", Some("id")),
            "left" | "top" | "width" | "height" => (CSSOM_MODULE, "getStyleProperty", Some("")),
            _ => (DOM_MODULE, "getAttribute", Some("")) };

        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
        self.chunk().emit_call(doc_idx, 0, line);
        self.compile_expr(object)?;
        let argc = match key {
            Some(k) => {
                let name = if k.is_empty() { prop.as_str() } else { k };
                emit_string_const(self.chunk(), name, line);
                3
            }
            None => 2 };
        let idx = self.import(module, func);
        self.emit_host_call(idx, argc);
        Ok(())
    }

    /// Create a control — the ONE place a frontend turns a canonical control
    /// name into bytecode.
    ///
    /// A control is not a bespoke host function, it is
    /// `document.createElement(tag)`. Lowering here rather than naming a
    /// `new_<Type>` factory means `web:*` stays WHATWG (there is no
    /// `new_Button` in any spec), and every language on this path — Pascal
    /// today, Flutter next — gets the same element.
    ///
    /// Constructor arguments are evaluated for their side effects and
    /// dropped: an owner/parent argument is a toolkit convention, and
    /// parenting happens through `appendChild` when the control is added.
    ///
    /// Stack on exit: [element]
    pub fn emit_control_element(&mut self, type_name: &str, argc: u8, line: u32) {
        let element = registered_control_element(&self.profile.namespaces.type_scopes, type_name)
            .unwrap_or_else(|| ControlElement::custom(type_name));
        for _ in 0..argc {
            self.chunk().emit_op(Op::DROP, line);
        }
        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
        self.chunk().emit_call(doc_idx, 0, line);
        // A document has exactly ONE body and you cannot create another —
        // `createElement("body")` is legal but yields a detached second body
        // that renders nothing. A form IS the document's body, so take it.
        // (It also showed up as an extra entry everywhere controls are
        // enumerated: a two-control form reported three.)
        if element.tag == "body" {
            let body_idx = self.import(DOCUMENT_MODULE, "body");
            self.chunk().emit_call(body_idx, 1, line);
            return;
        }
        emit_string_const(self.chunk(), &element.tag, line);
        emit_string_const(self.chunk(), &element.input_type, line);
        let create_idx = self.import(DOM_MODULE, HOST_FN_CREATE_ELEMENT);
        self.chunk().emit_call(create_idx, 3, line);
    }
}

/// Emit `vybe:gui::new_<Type>(args)` to create a new control.
/// `import_idx` must be the result of `compiler.import("vybe:gui", host_fn_new_control(canonical_type).as_str())`.
/// Stack on entry: [arg0, arg1, ...] (constructor args, if any)
/// Stack on exit: [control]
pub fn emit_new_control(chunk: &mut Chunk, import_idx: u16, argc: u8, line: u32) {
    chunk.emit_call(import_idx, argc, line);
}

/// Emit `vybe:gui::onEvent(control_name, event_name, handler_fn)`.
/// `import_idx` must be the result of `compiler.import("vybe:gui", HOST_FN_BIND_EVENT)`.
///
/// Stack on entry: [control_name_string, event_name_string, handler_fn]
/// Stack on exit:  [host_call_result]   (caller is responsible for dropping)
///
/// All `gui::emit_*` helpers leave the host call's return value on the stack
/// to keep the same convention as `compile_expr` and `compile_call`. Statement-
/// level callers (AddHandler / Controls.Add as a stmt / etc.) emit a `drop`
/// after calling these helpers. Expression-level callers leave it.
///
/// Frontends produce the three operands in their language-specific way:
/// - VB walker emits this for `Handles ctrl.Event` clauses
/// - C# walker emits this for `ctrl.Event += handler` statements
/// - JS walker emits this for `ctrl.addEventListener(event, handler)`
/// - etc.
pub fn emit_bind_event(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 3, line);
}

/// Emit `vybe:gui::removeEvent(control_name, event_name, handler_fn)`.
/// `import_idx` must be the result of `compiler.import("vybe:gui", HOST_FN_UNBIND_EVENT)`.
/// Caller drops the result if used as a statement.
pub fn emit_unbind_event(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 3, line);
}

/// Emit `vybe:gui::controlsAdd(parent, child)`.
/// `import_idx` must be `compiler.import("vybe:gui", HOST_FN_ADD_CHILD)`.
/// Stack on entry: [parent, child]
/// Stack on exit:  [host_call_result]   (caller drops if statement)
pub fn emit_add_child(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 2, line);
}

/// Emit `vybe:gui::runApplication(form)`.
/// `import_idx` must be `compiler.import("vybe:gui", HOST_FN_RUN_APPLICATION)`.
pub fn emit_run_application(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 1, line);
}

/// Emit `vybe:gui::appExit()`.
/// `import_idx` must be `compiler.import("vybe:gui", HOST_FN_APP_EXIT)`.
pub fn emit_app_exit(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 0, line);
}

/// Emit `vybe:gui::raiseEvent(arg0, ..., argN, event_name)`.
/// `import_idx` must be `compiler.import("vybe:gui", HOST_FN_RAISE_EVENT)`.
/// Stack on entry: [arg0, arg1, ..., event_name_string]
/// Stack on exit:  [host_call_result]   (caller drops if statement)
pub fn emit_raise_event(chunk: &mut Chunk, import_idx: u16, total_args: u8, line: u32) {
    chunk.emit_call(import_idx, total_args, line);
}

/// Push a string constant onto the stack (helper used when assembling
/// arguments for the GUI host calls above).
pub fn emit_string_const(chunk: &mut Chunk, s: &str, line: u32) {
    chunk.emit_string_const(s, line);
}

// ─── Read a control name from an object ──────────────────────────────────────

/// Field key under which a control instance stores its name.
/// Frontends and host fns both look this up.
pub const CONTROL_NAME_FIELD: &str = "__control_name";

/// Field key under which a control instance stores its type tag (e.g. "Button").
pub const CONTROL_TYPE_FIELD: &str = "__control_type";

/// Emit a struct_get to read the control's name field.
/// Stack on entry: [control_obj]
/// Stack on exit: [name_string]
pub fn emit_get_control_name(chunk: &mut Chunk, line: u32) {
    let key = chunk.add_constant(Value::String(Arc::from(CONTROL_NAME_FIELD)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
}

/// Emit a struct_get to read the control's type tag field.
/// Stack on entry: [control_obj]
/// Stack on exit: [type_string]
pub fn emit_get_control_type(chunk: &mut Chunk, line: u32) {
    let key = chunk.add_constant(Value::String(Arc::from(CONTROL_TYPE_FIELD)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
}
