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

use super::Compiler;
use super::{collections, ops, strings};
use std::sync::Arc;
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
        // `BindingSource` is NOT here. Naming a type in this table is what makes
        // `New <Type>()` an ELEMENT construction: `constructed_control_type_name`
        // gates on exactly this answer, and a name with no HTML counterpart
        // falls through to `ControlElement::custom` → `<vybe-bindingsource>`.
        // A binding source is a cursor over data — it has no element, nothing
        // paints it, and every member is a position or a list. Listed here it
        // got a document node, its `Position` became a CSS write, and the
        // platform's own constructor never ran. `BindingNavigator` stays: that
        // one IS a toolbar the user sees.
        "bindingnavigator" => "BindingNavigator",

        // ── Forms ──
        "form" | "window" => "Form",

        _ => return String::new(),
    }
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

/// The legacy GUI module, addressed BY CONTROL NAME rather than by object.
///
/// `vybe.gui.setProperty("display", "Text", v)` is a program-visible surface —
/// programs call it directly — so it survives the conversion as a spelling and
/// is lowered onto the document here, rather than reaching a host registry the
/// renderer no longer paints from.
pub const GUI_MODULE: &str = "vybe:gui";

/// Host fn name for "set a property on the control with this NAME".
/// Stack at call site: [name, prop_name, value]
pub const HOST_FN_SET_PROPERTY_BY_NAME: &str = "setProperty";

/// Host fn name for "get a property from the control with this NAME".
/// Stack at call site: [name, prop_name]
pub const HOST_FN_GET_PROPERTY_BY_NAME: &str = "getProperty";

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
                results: vec![ValType::Any],
            }),
        },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "addControl".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "addControl".to_string(),
                params: vec![],
                results: vec![],
            }),
        },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "setProperty".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "setProperty".to_string(),
                params: vec![],
                results: vec![],
            }),
        },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "getProperty".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "getProperty".to_string(),
                params: vec![],
                results: vec![ValType::Any],
            }),
        },
        // Event handling
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "onEvent".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "onEvent".to_string(),
                params: vec![],
                results: vec![],
            }),
        },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "removeEvent".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "removeEvent".to_string(),
                params: vec![],
                results: vec![],
            }),
        },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "raiseEvent".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "raiseEvent".to_string(),
                params: vec![],
                results: vec![],
            }),
        },
        // Collections
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "newControlsCollection".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "newControlsCollection".to_string(),
                params: vec![],
                results: vec![ValType::Any],
            }),
        },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "newComponentsCollection".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "newComponentsCollection".to_string(),
                params: vec![],
                results: vec![ValType::Any],
            }),
        },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "controlsAdd".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "controlsAdd".to_string(),
                params: vec![],
                results: vec![],
            }),
        },
        // Application
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "runApplication".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "runApplication".to_string(),
                params: vec![],
                results: vec![],
            }),
        },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "appExit".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "appExit".to_string(),
                params: vec![],
                results: vec![],
            }),
        },
        // Form lifecycle
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "showForm".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "showForm".to_string(),
                params: vec![],
                results: vec![],
            }),
        },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "closeForm".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "closeForm".to_string(),
                params: vec![],
                results: vec![],
            }),
        },
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "showFormDialog".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "showFormDialog".to_string(),
                params: vec![],
                results: vec![],
            }),
        },
        // Dialog
        ComponentExport {
            interface: GUI_MODULE.to_string(),
            name: "msgBox".to_string(),
            kind: ComponentItemKind::Function(FuncSig {
                name: "msgBox".to_string(),
                params: vec![],
                results: vec![ValType::I32],
            }),
        },
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

/// A control METHOD, by role — `gui.ctrl.<verb>`.
///
/// The same contract as the property roles one line up: a platform declares the
/// VERB its framework spells (`Show`, `Hide`, `BringToFront`) and this is the
/// only place a verb becomes a DOM operation. `TButton.Show` and
/// `Button.Show()` are one verb, so VCL and WinForms reach one implementation.
pub const CTRL_METHOD_EMIT: &str = "gui.ctrl.";

/// `container.Add(child)` — insertion spelled as a METHOD rather than as the
/// child's `Parent` property. Same DOM operation either way.
///
/// VCL menus are the case that needs it: `FMainMenu.Items.Add(MenuFile)`. The
/// old binding was `vybe:gui.controlsAdd`, which nests inside `GuiState` — so
/// a menu built that way never entered the document and could not render,
/// be hit-tested, or be listed by `widgets`.
pub const APPEND_CHILD_EMIT: &str = "gui.append_child";

/// `List.Items.Add(text)` — the sibling of [`APPEND_CHILD_EMIT`] for a list
/// whose entries are STRINGS rather than controls the caller built. A platform
/// declares this one on its list classes and the other on its containers; the
/// difference is who creates the element, and it is not inferable from the
/// call, which is why it is two declarations rather than one emit with a test.
pub const APPEND_ITEM_EMIT: &str = "gui.append_item";

/// `List.Items.Delete(index)` — `select.remove(index)`.
pub const REMOVE_ITEM_EMIT: &str = "gui.remove_item";

/// `List.Items[i]` — `select.options[i].text`, the DECLARED INDEXER pair. A
/// platform declares these as the `Item` property's two directions; both must
/// be present or the index site does not take the branch.
pub const ITEM_TEXT_EMIT: &str = "gui.item_text";
pub const SET_ITEM_TEXT_EMIT: &str = "gui.set_item_text";

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
            if setting {
                "setTextContent"
            } else {
                "textContent"
            },
            None,
        ),
        // The line count is DERIVED from the text — the DOM has no property
        // for it, and inventing a host function to answer it would put a
        // toolkit's question into `web:*`. So it reads the text like any other
        // text role and `emit_gui_property_get` counts, exactly as the geometry
        // roles read a CSS string and parse the unit off. Nothing writes it.
        "linecount" => (
            DOM_MODULE,
            if setting {
                "setTextContent"
            } else {
                "textContent"
            },
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
        // `ItemIndex` / `SelectedIndex` — `HTMLSelectElement.selectedIndex`,
        // its own IDL member and NOT `value`, which is the selected option's
        // value STRING. Without this arm the role fell to the attribute
        // fallback and read `getAttribute("selectedindex")`, which nothing ever
        // writes: every list answered null, so `ItemIndex >= 0` was false
        // everywhere and the kanban's Move Right exited before doing anything.
        // Silent, like every unmapped role.
        "selectedindex" => (
            DOCUMENT_MODULE,
            if setting {
                "setSelectedIndex"
            } else {
                "selectedIndex"
            },
            None,
        ),
        // A control's `Hint` IS the `title` attribute — HTML's own tooltip,
        // shown by every browser on hover with no script. The role was already
        // named `tooltip` and had no arm, so it wrote `tooltip="…"`, an
        // attribute nothing has ever read.
        "tooltip" => (
            DOM_MODULE,
            if setting {
                "setAttribute"
            } else {
                "getAttribute"
            },
            Some("title"),
        ),
        // VCL's `Tag` is a scratch integer the application gives meaning to.
        // HTML has exactly that: a `data-*` attribute is author data the UA
        // ignores. It round-tripped before as a bare `tag="…"`, which works and
        // is non-conforming markup — `data-tag` is the same behaviour spelled
        // the way the spec allows.
        "tag" => (
            DOM_MODULE,
            if setting {
                "setAttribute"
            } else {
                "getAttribute"
            },
            Some("data-tag"),
        ),
        // A control's `Name` IS the element id — what `getElementById` and
        // `<label for>` resolve. Not HTML's `name`, the submission key.
        "name" => (
            DOM_MODULE,
            if setting {
                "setAttribute"
            } else {
                "getAttribute"
            },
            Some("id"),
        ),
        // Boolean content attributes, INVERTED: true by PRESENCE, so
        // `Enabled := False` ADDS `disabled`. `toggleAttribute` is the DOM's
        // own add-or-remove.
        "enabled" => (
            DOM_MODULE,
            if setting {
                "toggleAttribute"
            } else {
                "getAttribute"
            },
            Some("disabled"),
        ),
        "visible" => (
            DOM_MODULE,
            if setting {
                "toggleAttribute"
            } else {
                "getAttribute"
            },
            Some("hidden"),
        ),
        // A control's own colour is its BACKGROUND — VCL's `Color`, WinForms'
        // `BackColor`. Text colour is a separate role because it is a separate
        // CSS property (`Font.Color` / `ForeColor` → `color`), and conflating
        // them is how a panel ends up painting its caption instead of itself.
        "backcolor" => (
            CSSOM_MODULE,
            if setting {
                "setStyleProperty"
            } else {
                "getStyleProperty"
            },
            Some("background-color"),
        ),
        "forecolor" => (
            CSSOM_MODULE,
            if setting {
                "setStyleProperty"
            } else {
                "getStyleProperty"
            },
            Some("color"),
        ),
        // Where the caption sits inside the control's own box. VCL spells it
        // `Alignment`, WinForms `TextAlign`, and both mean CSS `text-align` —
        // an INHERITED property, so a form declaring it once is how a whole
        // panel of labels lines up.
        //
        // Left unmapped it wrote `alignment="right"`, an attribute no element
        // reads: the calculator's display declared `taRightJustify` and its
        // text sat on the left.
        "textalign" => (
            CSSOM_MODULE,
            if setting {
                "setStyleProperty"
            } else {
                "getStyleProperty"
            },
            Some("text-align"),
        ),
        // Scrolling is a CSS property, not a control mode — `overflow` is the
        // whole of what a toolkit's `ScrollBars` means, and the frontend's
        // constants are declared as the CSS keywords so no enum is translated
        // on the way.
        "overflow" => (
            CSSOM_MODULE,
            if setting {
                "setStyleProperty"
            } else {
                "getStyleProperty"
            },
            Some("overflow"),
        ),
        // `dock` joins these because it is geometry too, just expressed as a
        // rule instead of a number: the container computes the rect from it.
        // A frontend that spells it `Align` (VCL) or `Dock` (WinForms) reaches
        // the same style property, and `vybe_widgets` owns the result.
        "left" | "top" | "width" | "height" | "dock" => (
            CSSOM_MODULE,
            if setting {
                "setStyleProperty"
            } else {
                "getStyleProperty"
            },
            Some(""),
        ),
        _ => (
            DOM_MODULE,
            if setting {
                "setAttribute"
            } else {
                "getAttribute"
            },
            Some(""),
        ),
    }
}

/// The two roles a PAIR role decomposes into, with the field each component
/// reads off the value.
///
/// `Location` and `Size` are one framework value carrying two CSS
/// declarations. Answering here rather than in a `property_op` row is what
/// keeps the decomposition a re-entry into the ordinary write path: the units,
/// the `px` suffix and the CSS operation are `left`/`top`/`width`/`height`'s
/// own, stated once.
///
/// The field names are the ones the value type actually stores (`vybe:gui`'s
/// `pointNew`/`sizeNew`), not the framework's property spelling — `Point`
/// declares `X`/`Y` and stores `x`/`y`.
fn pair_role_components(role: &str) -> Option<[(&'static str, &'static str); 2]> {
    match role {
        "location" => Some([("x", "left"), ("y", "top")]),
        "size" => Some([("width", "width"), ("height", "height")]),
        _ => None,
    }
}

/// The TYPE a property role's value has, for the roles whose type the DOM
/// operation above already fixes.
///
/// A property that reads through `textContent` yields text, one that reads a
/// boolean content attribute yields a boolean, and CSS geometry comes back a
/// number because `emit_gui_property_get` parses the unit off. Declaring it is
/// what lets the ordinary expression machinery work on the result — Delphi's
/// `(Sender as TButton).Caption[1]` is a string subscript, and with no declared
/// type it read `null` instead of the character.
///
/// Only the roles whose type the operation settles are answered. An unmapped
/// role becomes an attribute of arbitrary meaning (`Anchors`, `Font`, `Items`),
/// and guessing there would be worse than the `None` that leaves inference
/// exactly as it was.
pub fn property_value_type(role: &str) -> Option<&'static str> {
    match role {
        "text" | "caption" | "value" | "name" => Some("string"),
        "checked" | "ischecked" | "enabled" | "visible" => Some("Boolean"),
        "left" | "top" | "width" | "height" | "linecount" | "selectedindex" => Some("Integer"),
        _ => None,
    }
}

/// The event type an `on<type>` role registers, if it is one.
///
/// A ROLE, exactly like the property roles: every language spells its handler
/// slot `OnClick` / `onClick` / `Click`, and each lowers to the one DOM event
/// type. Nothing here is per-language and nothing is enumerated — HTML's own
/// convention is that an event handler IDL attribute is `on` + the type.
fn event_role_type(role: &str) -> Option<&str> {
    role.strip_prefix("on")
        .filter(|type_name| !type_name.is_empty())
}

impl Compiler {
    /// The slot holding the receiver, when this code is inside a method or a
    /// constructor. Resolved the same way `ExprKind::This` resolves it, and
    /// through the profile's own keyword so no language is named.
    fn receiver_slot(&mut self) -> Option<u16> {
        let self_keyword = self.profile.self_keyword.clone();
        self.scope()
            .resolve(&self_keyword)
            .or_else(|| self.scope().resolve("Self"))
            .or_else(|| self.scope().resolve("self"))
            .or_else(|| self.scope().resolve("this"))
    }

    /// Lower `gui.append_child` — stack in `[parent, child]`, out `[child]`.
    ///
    /// The operands already arrive container-first, which is `appendChild`'s
    /// own order — unlike the `parent` ROLE, where the assignment names the
    /// child first and the two have to be swapped.
    /// `List.Items.Add(text)` — stack in `[list, text]`, out `[]`.
    ///
    /// An item is an OPTION, not an attribute — and the DOM already has the
    /// operation: `web:html.addItem` IS `select.add(option)`, which is why this
    /// builds no element of its own. Creating an `<option>` here and appending
    /// it would have been a second way to say the same thing, and it renders
    /// nothing: a list widget takes items, not child controls, so the element
    /// would have been silently detached.
    ///
    /// Without it `Items` had no role at all and fell to the `_` arm as
    /// `getAttribute("items")`, which answers null — so `Items.Add` called
    /// `undefined` and took two shipped examples down with it. The lesson is
    /// the attribute FALLBACK: an unmapped role is silent, and silently wrong
    /// for anything whose value is not text.
    ///
    /// Distinct from `TMainMenu.Items`, whose entries are controls the caller
    /// already built — that stays `appendChild` of the element it was handed.
    /// Here the caller has a STRING and the element is ours to make.
    pub fn emit_gui_append_item(&mut self, line: u32) {
        let text = self.define_local("__gui_item_text");
        let list = self.define_local("__gui_item_list");
        self.emit_u16(Op::LOCAL_SET, text);
        self.emit_u16(Op::LOCAL_SET, list);
        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
        self.chunk().emit_call(doc_idx, 0, line);
        self.emit_u16(Op::LOCAL_GET, list);
        self.emit_u16(Op::LOCAL_GET, text);
        let idx = self.import(DOCUMENT_MODULE, "addItem");
        self.emit_host_call(idx, 3);
    }

    /// `List.Items.Delete(index)` — `select.remove(index)`.
    ///
    /// A control VERB cannot express this: `gui.ctrl.<verb>` is `[control]` in,
    /// one value out, with nowhere to put an argument. So the index-taking
    /// members of the option list get their own emit, exactly as `Add` did.
    pub fn emit_gui_remove_item(&mut self, line: u32) {
        let index = self.define_local("__gui_item_index");
        let list = self.define_local("__gui_item_list");
        self.emit_u16(Op::LOCAL_SET, index);
        self.emit_u16(Op::LOCAL_SET, list);
        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
        self.chunk().emit_call(doc_idx, 0, line);
        self.emit_u16(Op::LOCAL_GET, list);
        self.emit_u16(Op::LOCAL_GET, index);
        let idx = self.import(DOCUMENT_MODULE, "removeItem");
        self.emit_host_call(idx, 3);
    }

    /// `List.Items[i]` — `select.options[i].text`.
    ///
    /// Reached through the DECLARED INDEXER route, not a new mechanism: a
    /// registered type declaring an instance property named `Item` with a
    /// common emit in each direction makes `x[i]` lower to a two-argument emit
    /// here (`declared_indexer_emits`). That is .NET's `this[int]`, and
    /// Delphi's `TStrings.Strings[i]` is the same default indexed property.
    ///
    /// Stack in `[control, index]`, out `[text]`.
    pub fn emit_gui_item_text(&mut self, line: u32) {
        let index = self.define_local("__gui_item_index");
        let list = self.define_local("__gui_item_list");
        self.emit_u16(Op::LOCAL_SET, index);
        self.emit_u16(Op::LOCAL_SET, list);
        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
        self.chunk().emit_call(doc_idx, 0, line);
        self.emit_u16(Op::LOCAL_GET, list);
        self.emit_u16(Op::LOCAL_GET, index);
        let idx = self.import(DOCUMENT_MODULE, "itemText");
        self.emit_host_call(idx, 3);
    }

    /// The write half — `List.Items[i] := 'text'`. Declared alongside the read
    /// because `declared_indexer_emits` requires BOTH directions before it will
    /// take the branch at all, which is what stops a type offering a readable
    /// index and a silently ignored write.
    ///
    /// Stack in `[control, index, text]`, out `[value]` — the caller drops it.
    pub fn emit_gui_set_item_text(&mut self, line: u32) {
        let text = self.define_local("__gui_item_text");
        let index = self.define_local("__gui_item_index");
        let list = self.define_local("__gui_item_list");
        self.emit_u16(Op::LOCAL_SET, text);
        self.emit_u16(Op::LOCAL_SET, index);
        self.emit_u16(Op::LOCAL_SET, list);
        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
        self.chunk().emit_call(doc_idx, 0, line);
        self.emit_u16(Op::LOCAL_GET, list);
        self.emit_u16(Op::LOCAL_GET, index);
        self.emit_u16(Op::LOCAL_GET, text);
        let idx = self.import(DOCUMENT_MODULE, "setItemText");
        self.emit_host_call(idx, 4);
    }

    pub fn emit_gui_append_child(&mut self, line: u32) {
        let child = self.define_local("__gui_add_child");
        let parent = self.define_local("__gui_add_parent");
        self.emit_u16(Op::LOCAL_SET, child);
        self.emit_u16(Op::LOCAL_SET, parent);
        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
        self.chunk().emit_call(doc_idx, 0, line);
        self.emit_u16(Op::LOCAL_GET, parent);
        self.emit_u16(Op::LOCAL_GET, child);
        let idx = self.import(DOM_MODULE, "appendChild");
        self.emit_host_call(idx, 3);
    }

    /// Apply one declared constructor argument to the control being built.
    ///
    /// A DECLARATIVE frontend puts configuration in the constructor —
    /// `ElevatedButton(onPressed: f, child: Text('7'))` — so the same two
    /// operations an imperative frontend spells as later statements arrive at
    /// the construction site instead. `FieldGui` is the adapter's declaration of
    /// which one each argument is, and this is where it becomes bytecode: the
    /// emit stays in `gui.rs`, and the construction site only calls it.
    ///
    /// Stack in `[control, value]`, out `[]` — the value is consumed.
    ///
    /// This is what replaces stamping an `__ops` array for a realizer written in
    /// the target language to walk. A role only exists at compile time (it is
    /// carried in the emit NAME, `gui.prop_set.<role>`), so target-language
    /// source can only ever re-derive it — which is the duplicated role→DOM
    /// match the plan forbids. Applying it here is the one place the role is
    /// still known.
    pub fn emit_gui_field(
        &mut self,
        field: &vybe_runtime::namespaces::FieldGui,
        line: u32,
    ) {
        use vybe_runtime::namespaces::FieldGui;
        match field {
            // A child widget nests; a scalar sets the property. Both spellings
            // reach an operation that already exists, so neither needs a new
            // one.
            FieldGui::NestOrProp(key) => {
                let role = key.to_ascii_lowercase();
                self.emit_gui_property_set(&role, line);
            }
            // A LIST of children — `Column(children: [...])`. Each element
            // appends, in order.
            //
            // The value is an array, not one widget, so this is a loop and not
            // a single `appendChild`. Passing the array itself would append one
            // node that is not an element and silently produce an empty
            // container.
            FieldGui::Children => {
                let list = self.define_local("__gui_children");
                let parent = self.define_local("__gui_children_parent");
                let index = self.define_local("__gui_children_i");
                let count = self.define_local("__gui_children_n");
                self.emit_u16(Op::LOCAL_SET, list);
                self.emit_u16(Op::LOCAL_SET, parent);

                self.emit_u16(Op::LOCAL_GET, list);
                collections::emit_array_length(self.chunk(), line);
                self.emit_u16(Op::LOCAL_SET, count);
                self.emit_const(Value::I32(0));
                self.emit_u16(Op::LOCAL_SET, index);

                let current = self.current;
                let state =
                    crate::primitives::loops::emit_loop_start(&mut self.chunks, current, line);
                self.emit_u16(Op::LOCAL_GET, index);
                self.emit_u16(Op::LOCAL_GET, count);
                ops::emit_dyn_lt(&mut self.chunks[current], line);
                crate::primitives::loops::emit_loop_cond(&mut self.chunks, current, line);

                self.emit_u16(Op::LOCAL_GET, parent);
                self.emit_u16(Op::LOCAL_GET, list);
                self.emit_u16(Op::LOCAL_GET, index);
                collections::emit_get(&mut self.chunks, current, line);
                self.emit_gui_append_child(line);

                self.emit_u16(Op::LOCAL_GET, index);
                self.emit_const(Value::I32(1));
                self.emit(Op::I32_ADD);
                self.emit_u16(Op::LOCAL_SET, index);
                crate::primitives::loops::emit_loop_end(&mut self.chunks, current, state, line);
            }
            // A handler wires to the control's event. `on` + the type is the
            // whole rule — the same one a property-spelled handler goes through.
            FieldGui::Event(name) => {
                let role = format!("on{}", name.to_ascii_lowercase());
                self.emit_gui_property_set(&role, line);
            }
            // The child's text IS this control's caption, not a nested control.
            FieldGui::Caption => {
                self.emit_gui_property_set("text", line);
            }
        }
    }

    /// Lower `gui.ctrl.<verb>` — stack in `[control]`, out `[value]`.
    ///
    /// A toolkit's control verbs ARE DOM operations once you stop treating the
    /// control as an object beside the document:
    ///
    /// | verb | what it IS |
    /// |---|---|
    /// | `show` / `hide` | the `visible` ROLE — the `hidden` attribute |
    /// | `focus` | `HTMLElement.focus()`, which the DOM has outright |
    /// | `bring_to_front` | re-`appendChild` — last child paints on top |
    /// | `send_to_back` | `insertBefore(parent.firstChild)` |
    /// | `refresh` / `invalidate` / `update` | **nothing** |
    ///
    /// The repaint trio really is nothing, and that is a statement about the
    /// model rather than a shortcut. WinForms and the VCL are IMMEDIATE-mode at
    /// the paint layer: the app owns the pixels and must ask for them back. A
    /// document is RETAINED — mutate the tree and the engine repaints what
    /// changed. There is no DOM call these can lower to because there is
    /// nothing left for the author to ask for. Emitting a no-op is the honest
    /// answer; inventing a host function to receive it would be a shim.
    ///
    /// Every arm leaves exactly one value so a call in expression position
    /// behaves like any other, and statement position drops it as usual.
    pub fn emit_gui_control_method(&mut self, verb: &str, line: u32) {
        // `Show`/`Hide` are the `visible` role with the answer already known,
        // so they route through the SAME lowering a `Visible := True` write
        // takes. Not a parallel path — one role, two spellings, which is what
        // stops the two drifting on what `hidden` means.
        if let Some(visible) = match verb {
            "show" => Some(true),
            "hide" => Some(false),
            _ => None,
        } {
            self.emit_const(Value::Bool(visible));
            self.emit_gui_property_set("visible", line);
            return;
        }

        // `Clear` on a TEXT control — `FInput.Clear`, `FMemo.Clear`. The DOM
        // has no `clear()`: emptying a field IS `value = ""`, so this takes the
        // same route a `Text := ''` write takes, exactly as `Show`/`Hide` take
        // the `visible` route above. One role, two spellings.
        //
        // A LIST's `Clear` is a different operation on a different element
        // (`select.length = 0`) and is declared separately as
        // `gui.ctrl.clearItems`, which the web platform names outright. The two
        // share a spelling and nothing else, which is why each class says which
        // one it means rather than this guessing from the receiver.
        if verb == "clear" {
            emit_string_const(self.chunk(), "", line);
            self.emit_gui_property_set("text", line);
            return;
        }

        // `Refresh`/`Invalidate`/`Update` — see the table above. Drop the
        // receiver and answer null, so the stack contract still holds.
        if matches!(verb, "refresh" | "invalidate" | "update") {
            self.emit(Op::DROP);
            self.emit_null();
            return;
        }

        let ctrl = self.define_local("__gui_ctrl_recv");
        self.emit_u16(Op::LOCAL_SET, ctrl);
        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);

        // Z-ORDER IS DOCUMENT ORDER. Both verbs re-parent the control among its
        // existing siblings rather than assigning a `z-index`: painting order
        // is what the toolkit means, `appendChild` on an already-parented node
        // MOVES it (the DOM's own rule, not a trick), and a stacking context
        // created by `z-index` would change how descendants composite.
        if matches!(verb, "bring_to_front" | "send_to_back") {
            let parent = self.define_local("__gui_ctrl_parent");
            self.chunk().emit_call(doc_idx, 0, line);
            self.emit_u16(Op::LOCAL_GET, ctrl);
            let parent_idx = self.import(DOM_MODULE, "parentNode");
            self.emit_host_call(parent_idx, 2);
            self.emit_u16(Op::LOCAL_SET, parent);

            self.chunk().emit_call(doc_idx, 0, line);
            self.emit_u16(Op::LOCAL_GET, parent);
            self.emit_u16(Op::LOCAL_GET, ctrl);
            if verb == "bring_to_front" {
                let idx = self.import(DOM_MODULE, "appendChild");
                self.emit_host_call(idx, 3);
            } else {
                // The reference node is the CURRENT first child. Reading it
                // before the move is why this is `insertBefore` and not a
                // second `appendChild`.
                self.chunk().emit_call(doc_idx, 0, line);
                self.emit_u16(Op::LOCAL_GET, parent);
                let first_idx = self.import(DOM_MODULE, "firstChild");
                self.emit_host_call(first_idx, 2);
                let idx = self.import(DOM_MODULE, "insertBefore");
                self.emit_host_call(idx, 4);
            }
            return;
        }

        // `focus` — and anything a platform declares later that the web
        // platform names identically. The verb IS the method, which is the
        // point of spelling these as roles.
        //
        // It lives on `web:html`, not `web:dom`: focusing is an HTMLElement
        // behaviour, not a Node one, and the split is the spec's own. Naming
        // the wrong module is not a silent miss — it is an `Unresolved import`
        // at run time, which is how this was caught.
        self.chunk().emit_call(doc_idx, 0, line);
        self.emit_u16(Op::LOCAL_GET, ctrl);
        let idx = self.import(DOCUMENT_MODULE, verb);
        self.emit_host_call(idx, 2);
    }

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
            None => 2,
        };
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
            // Lines, from the text just read. Splitting on `\n` gives one more
            // element than there are lines whenever the text is empty or ends
            // in a break — and in exactly those cases the LAST element is the
            // empty string. So the correction is `1` precisely when that
            // element has length zero, which `i32.eqz` already answers as 1 or
            // 0; subtracting it needs no branch.
            //
            //   ""       → [""]            → 1 - 1 = 0
            //   "a"      → ["a"]           → 1 - 0 = 1
            //   "a\n"    → ["a", ""]       → 2 - 1 = 1
            //   "a\nb"   → ["a", "b"]      → 2 - 0 = 2
            //
            // A `\r\n` break leaves its `\r` on the end of the previous line,
            // which changes no length that matters here.
            "linecount" => {
                emit_string_const(self.chunk(), "\n", line);
                strings::emit_split(self.chunk(), line);
                let lines = self.define_local("__gui_lines");
                self.emit_u16(Op::LOCAL_SET, lines);

                self.emit_u16(Op::LOCAL_GET, lines);
                collections::emit_array_length(self.chunk(), line);

                self.emit_u16(Op::LOCAL_GET, lines);
                self.emit_u16(Op::LOCAL_GET, lines);
                collections::emit_array_length(self.chunk(), line);
                self.emit_const(Value::I32(1));
                self.emit(Op::I32_SUB);
                self.emit(Op::ARRAY_GET);
                strings::emit_length(self.chunk(), line);
                self.emit(Op::I32_EQZ);

                self.emit(Op::I32_SUB);
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
            // A DOM listener is called by the document with an Event and
            // NOTHING else — there is no second channel for a receiver, and a
            // host that keeps one (`GuiState.form_object`, `__f`) is holding
            // form state the document already owns.
            //
            // Every language that spells this assignment means a BOUND method:
            // Delphi's `OnClick := bClick` inside `TForm.Create` is a method
            // pointer, which is a `(Self, code)` pair, and a C#/VB
            // `Click += OnClick` is a delegate over `this`. The receiver is
            // part of the VALUE in all of them, so it is bound here, once,
            // with the spec's own operation rather than carried out of band.
            // Outside a method there is no receiver and none is bound — a free
            // function is already a complete listener.
            if let Some(receiver) = self.receiver_slot() {
                self.emit_u16(Op::LOCAL_GET, receiver);
                // HOW a receiver reaches a body is the profile's answer, not a
                // language's name. Where `this` is ambient (ECMA §10.2.1.1 —
                // JS, Dart) `bind`'s `thisArg` IS the binding and the handler's
                // own parameters are untouched. Where the receiver is an
                // explicit first parameter it must also arrive as a bound
                // ARGUMENT, or slot 0 stays empty and the handler reads the
                // Event as its own `Self` — which is precisely the nil-receiver
                // failure this replaces.
                //
                // The SAME frameworks bind the control in as well. A DOM
                // listener is handed an Event; a method-pointer/delegate
                // framework declares its handler `(Sender: TObject)` /
                // `(object sender, EventArgs e)` and means the control the
                // listener is attached to — `currentTarget`, which is exactly
                // the `ctrl` this registration is for. Bound here, it arrives
                // ahead of the Event, so VCL reads the control as `Sender` and
                // WinForms reads `(control, event)`; without it `Sender` IS the
                // Event and `(Sender as TButton).Caption` reads nothing.
                let argc = if self.ambient_this() {
                    2
                } else {
                    self.emit_u16(Op::LOCAL_GET, receiver);
                    self.emit_u16(Op::LOCAL_GET, ctrl);
                    4
                };
                let bind_idx = self.import("ecma:function", "bind");
                self.emit_host_call(bind_idx, argc);
            }
            let idx = self.import(DOM_MODULE, "addEventListener");
            self.emit_host_call(idx, 4);
            return;
        }
        // `child.Parent := container` IS `container.appendChild(child)`. VCL,
        // WinForms and MAUI all spell insertion as a property on the CHILD,
        // which is why it needs its own branch rather than a `property_op` row:
        // the operands are the other way round from every other setter, and the
        // DOM's insertion op takes the container first.
        //
        // Without this the assignment fell through to `setAttribute("parent",
        // <element>)` — the control stored as markup on itself, never inserted,
        // so a form built the VCL way came up empty while every control existed.
        // The same insertion the other way round: `form.Menu := m` names the
        // CONTAINER on the left, which is already the operand order the DOM
        // wants, so this is `Add` spelled as a property. Same emit, so a
        // frontend may map whichever spelling it has and get one behavior.
        if role == "child" {
            self.emit_gui_append_child(line);
            return;
        }
        if role == "parent" {
            let parent = self.define_local("__gui_parent");
            let child = self.define_local("__gui_child");
            self.emit_u16(Op::LOCAL_SET, parent);
            self.emit_u16(Op::LOCAL_SET, child);
            let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
            self.chunk().emit_call(doc_idx, 0, line);
            self.emit_u16(Op::LOCAL_GET, parent);
            self.emit_u16(Op::LOCAL_GET, child);
            let idx = self.import(DOM_MODULE, "appendChild");
            self.emit_host_call(idx, 3);
            return;
        }
        // A role whose VALUE is a pair — `Location` is `(Left, Top)` and `Size`
        // is `(Width, Height)`. Unlike `Anchor`/`Font`/`BackColor`, these two DO
        // have IDL counterparts; they are just spelled as one value type in the
        // framework and as two declarations in CSS. So this decomposes and
        // re-enters, which is why the units, the `px` suffix and the write path
        // are shared with `Left`/`Top` rather than restated.
        //
        // Without it, `Me.btn0.Location = New Point(20, 50)` fell to the
        // `setAttribute` catch-all and the document serialised
        // `location="[object]"` — a designer-generated form is written almost
        // ENTIRELY in these two properties, so every control in it sat at the
        // origin while the form otherwise looked correct.
        if let Some([(x_field, x_role), (y_field, y_role)]) = pair_role_components(role) {
            let value = self.define_local("__gui_pair_value");
            let ctrl = self.define_local("__gui_pair_ctrl");
            self.emit_u16(Op::LOCAL_SET, value);
            self.emit_u16(Op::LOCAL_SET, ctrl);
            for (index, (field, component_role)) in [(x_field, x_role), (y_field, y_role)]
                .into_iter()
                .enumerate()
            {
                self.emit_u16(Op::LOCAL_GET, ctrl);
                self.emit_u16(Op::LOCAL_GET, value);
                let key = self.str_const(field);
                self.emit_struct_field_op(Op::STRUCT_GET, 0, key);
                self.emit_gui_property_set(component_role, line);
                // Each write leaves the host call's result, and this function
                // promises exactly one — so all but the last are dropped.
                if index == 0 {
                    self.emit(Op::DROP);
                }
            }
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
                // A CSS length is TEXT WITH A UNIT. `Left := 8` is
                // `style.left = "8px"`, and a browser DROPS `"8"` outright —
                // `<length>` has no unitless form outside `0`. The read path
                // has always said so, parsing the unit back off with
                // `parseFloat`; the write side did not, and only
                // `vybe_widgets`' own lenient `parse_px` hid it. That is a
                // defect no capture can show, because both ends of OUR
                // pipeline agreed on the wrong thing.
                //
                // `dock` shares this arm and must NOT get a unit: it carries
                // an edge keyword, not a length.
                if matches!(role, "left" | "top" | "width" | "height") {
                    strings::emit_to_string(self.chunk(), line);
                    emit_string_const(self.chunk(), "px", line);
                    ops::emit_dyn_add(self.chunk(), line);
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
            input_type: input_type.trim().to_ascii_lowercase(),
        }
    }

    /// Is this element FORM-ASSOCIATED — i.e. does it belong to
    /// `form.elements` and get submitted?
    ///
    /// HTML's own list, not a judgement call: button, fieldset, input,
    /// object, output, select, textarea. A `<img>`, `<div>`, `<ul>`,
    /// `<progress>` or `<table>` is NOT one — a `PictureBox` or `Panel` is a
    /// control in the toolkit sense but carries no submission identity, and
    /// `name` on it is non-conforming markup that submits nothing.
    /// Does this element need to be a containing block for its children?
    ///
    /// The elements that hold other controls. `body` is excluded: the viewport
    /// is already the initial containing block, so a child of the form resolves
    /// correctly with nothing declared — and a browser would agree.
    ///
    /// Anything that is not a container has no positioned descendants to
    /// anchor, so declaring it would be noise in the markup.
    pub fn establishes_containing_block(&self) -> bool {
        matches!(
            self.tag.as_str(),
            "div"
                | "section"
                | "article"
                | "main"
                | "aside"
                | "header"
                | "footer"
                | "nav"
                | "form"
                | "fieldset"
                | "dialog"
                | "li"
                | "td"
                | "th"
        ) || self.tag.starts_with("vybe-")
    }

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
        // ⚠ Strip Delphi's `T` prefix — and ONLY that.
        //
        // `trim_start_matches(['T', 't'])` removed EVERY leading T, and after
        // lowercasing there was no case left to tell a prefix from a word:
        // `TableLayoutPanel` became `<vybe-ablelayoutpanel>` and `ToolStrip`
        // would have become `<vybe-oolstrip>`. A tag naming no known control
        // degrades to a 120x20 label, so the whole control silently vanished
        // — see [[project_pseudo_tag_renders_as_a_label]].
        //
        // The convention is `T` followed by an UPPERCASE letter (`TButton`,
        // `TForm`), which is why the test runs before lowercasing. A WinForms
        // name has no prefix at all and now keeps every letter.
        let stripped = bare
            .strip_prefix('T')
            .filter(|rest| rest.chars().next().is_some_and(|c| c.is_ascii_uppercase()))
            .unwrap_or(bare);
        let bare = stripped.to_ascii_lowercase();
        ControlElement {
            tag: format!("vybe-{}", if bare.is_empty() { "control" } else { &bare }),
            input_type: String::new(),
        }
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

    /// The control type a `Name(...)` / `A.B.Name(...)` call CONSTRUCTS, if it
    /// constructs one.
    ///
    /// One fact asked at both ends. `compile_call` uses this to decide it
    /// should emit `emit_control_element` instead of a function call, and
    /// `infer_expr_type_hint` uses it to say what such a call is WORTH. Asked
    /// with two different predicates the two answers drift apart, and a call
    /// typed as a control that was never lowered as one sends every later
    /// property write down the DOM path for something that is not an element —
    /// silently, which is the failure this exists to prevent.
    ///
    /// The callee resolves by its LAST segment; namespace qualifiers are
    /// ignored, matching how the class name used to resolve as a global.
    /// `canonical_control_name` alone cannot decide it — that shared table
    /// holds generic words (`image`, `panel`, `label`) — so REGISTRATION in
    /// this profile's own type scopes is the signal, and a user
    /// function/class/local of the same name shadows.
    pub(super) fn constructed_control_type_name(&self, callee: &Expression) -> Option<String> {
        let parts = self.flatten_member_chain(callee);
        let last = parts.last()?;
        if canonical_control_name(last).is_empty() {
            return None;
        }
        if parts
            .first()
            .is_some_and(|first| self.scope().resolve(first).is_some())
        {
            return None;
        }
        let canon_last = self.canon(last);
        if self.defined_functions.contains(&canon_last)
            || self.defined_classes.contains(&canon_last)
        {
            return None;
        }
        vybe_runtime::namespaces::is_registered_type(&self.profile.namespaces.type_scopes, last)
            .then_some(canon_last)
    }

    /// The ROLE a class declares for one of its properties.
    ///
    /// The class is the authority on what its property means: plib declares
    /// that `ClientWidth` fills the `width` role, `Hint` the `tooltip` role,
    /// `Menu` the `child` role. Asking it here is what keeps the per-language
    /// vocabulary in the platform that owns it — this file only ever sees a
    /// role.
    ///
    /// Answers `None` when nothing is declared, and the caller keeps the source
    /// spelling. That is why the bypass went unnoticed for so long: a property
    /// whose Pascal word happens to EQUAL its role (`Caption`, `Width`,
    /// `Enabled`) lands correctly either way, so only the renames were broken —
    /// silently, as an attribute no widget reads.
    ///
    /// Walks the user chain for the same reason `control_element_for_type`
    /// does: `TForm1 = class(TForm)` inherits `TForm`'s declarations, and the
    /// receiver's static type is the subclass.
    fn declared_property_role(&self, type_name: &str, prop: &str, setting: bool) -> Option<String> {
        let scopes = &self.profile.namespaces.type_scopes;
        let declared = |name: &str| {
            let target = if setting {
                vybe_runtime::namespaces::lookup_type_property_setter_target(scopes, name, prop)
            } else {
                vybe_runtime::namespaces::lookup_type_property_target(scopes, name, prop)
            }?;
            match target {
                vybe_runtime::component_model::InstancePropertyTarget::Common { emit } => emit
                    .strip_prefix(if setting {
                        PROP_SET_EMIT
                    } else {
                        PROP_GET_EMIT
                    })
                    .map(str::to_string),
                // A host-backed accessor is already a complete target; it is
                // not a role and must not be rewritten into one.
                _ => None,
            }
        };
        if let Some(role) = declared(type_name) {
            return Some(role);
        }
        let mut current = self.pending_class_parent(type_name);
        while let Some(parent) = current {
            if let Some(role) = declared(&parent) {
                return Some(role);
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
    /// Is a LATE-BOUND member access worth a runtime element test?
    ///
    /// Only where the compiler has no static answer at all. A receiver whose
    /// type is known takes the direct emit above; this is for the case that was
    /// already heading for a bare `struct_get`/`struct_set`, so the branch adds
    /// an answer where there was none and changes nothing that had one.
    ///
    /// Gated on the profile declaring a type tree, which is the same structural
    /// test every other tree consult here uses — a language that registers none
    /// cannot own a control and pays nothing.
    pub(super) fn member_access_is_late_bound(&self, object: &Expression, field: &str) -> bool {
        if self.profile.namespaces.type_scopes.is_empty() {
            return false;
        }
        // The receiver must be a VALUE held in a slot. A namespace root, a
        // class name or a chain has no type hint either, and none of them can
        // be an element — testing them at runtime would put a branch in front
        // of every static member read for nothing.
        if !matches!(&object.kind, vybe_ast::ExprKind::Ident(name)
            if self.scope().resolve(name).is_some())
        {
            return false;
        }
        // A field whose STORAGE name differs from its spelling is resolved
        // already — that resolution is what the ordinary read/write path
        // applies, and this branch would drop it.
        if self
            .field_storage_name_for_receiver(object, field)
            .is_some()
        {
            return false;
        }
        match self.infer_expr_type_hint(object) {
            None => true,
            Some(hint) => {
                let class_name = Self::normalize_type_hint(&hint);
                let scopes = &self.profile.namespaces.type_scopes;
                // The test is "does the receiver's type DECLARE this member",
                // not "is the type registered". `Object` IS registered — it is
                // the root of the .NET hierarchy — so an `is_registered_type`
                // test answered yes for the one spelling this exists to catch
                // and the branch never fired.
                self.control_element_for_type(&class_name).is_none()
                    && self
                        .resolve_pending_class_name_for_type_hint(&class_name)
                        .is_none()
                    && vybe_runtime::namespaces::lookup_type_property_target(
                        scopes,
                        &class_name,
                        field,
                    )
                    .is_none()
                    && vybe_runtime::namespaces::lookup_type_property_setter_target(
                        scopes,
                        &class_name,
                        field,
                    )
                    .is_none()
                    && vybe_runtime::namespaces::lookup_type_instance_member(
                        scopes,
                        &class_name,
                        field,
                    )
                    .is_none()
                    && !self.is_declared_instance_field(&class_name, field)
            }
        }
    }

    /// `x.<prop> = value` where the compiler does not know what `x` IS.
    ///
    /// **This restores a decision the conversion moved.** `vybe:gui`'s
    /// `controlSetProperty(obj, prop, value)` took the OBJECT at runtime, so a
    /// late-bound receiver reached the widget no matter what the compiler knew.
    /// The DOM path is chosen from the receiver's STATIC type, which is
    /// strictly better where a type exists — and where none does, left no path
    /// at all. `Sub HandleClick(btn As Object)` writing `btn.Text` emitted a
    /// bare `struct_set` onto the element and the document never heard about
    /// it; Pascal was unaffected only because `(Sender as TButton).Caption` and
    /// `FDisplay: TEdit` name a type at every access.
    ///
    /// So the test moves to runtime for exactly the accesses that have no
    /// static answer: `__control_type` is stamped on every control at
    /// construction, so "is this a node in the document" is one field read.
    ///
    /// Stack on entry: [value]. Stack on exit: empty.
    pub fn emit_late_bound_property_set(
        &mut self,
        object: &Expression,
        prop: &str,
        line: u32,
    ) -> Result<(), String> {
        let value_tmp = self.define_local("__late_prop_value");
        self.emit_u16(Op::LOCAL_SET, value_tmp);
        let obj_tmp = self.define_local("__late_prop_obj");
        self.compile_expr(object)?;
        self.emit_u16(Op::LOCAL_SET, obj_tmp);

        self.emit_u16(Op::LOCAL_GET, obj_tmp);
        let type_key = self.str_const(CONTROL_TYPE_FIELD);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, type_key);
        let undef_idx = self.import("wasm:js-undefined", "test");
        self.emit_host_call(undef_idx, 1);
        self.chunk().emit_if(line);
        // Not a control — the ordinary object property, exactly as before.
        self.emit_u16(Op::LOCAL_GET, obj_tmp);
        self.emit_u16(Op::LOCAL_GET, value_tmp);
        let prop_key = self.str_const(&self.canon(prop));
        self.emit_struct_field_op(Op::STRUCT_SET, 0, prop_key);
        self.chunk().emit_else(line);
        // A control. There is no class to ask for a ROLE here — that is what
        // "late bound" means — so the property's own spelling IS the role, the
        // same fallback `emit_control_property_set` uses when a class declares
        // nothing for the name.
        self.emit_u16(Op::LOCAL_GET, obj_tmp);
        self.emit_u16(Op::LOCAL_GET, value_tmp);
        self.emit_gui_property_set(&prop.to_ascii_lowercase(), line);
        self.emit(Op::DROP);
        self.chunk().emit_end(line);
        Ok(())
    }

    /// `x.<prop>` where the compiler does not know what `x` IS — the mirror of
    /// [`Self::emit_late_bound_property_set`], and needed for the same reason:
    /// `If btn.Text <> ""` read `undefined` off the element and took the wrong
    /// branch before any write was even attempted.
    ///
    /// Stack on entry: empty. Stack on exit: [value].
    pub fn emit_late_bound_property_get(
        &mut self,
        object: &Expression,
        prop: &str,
        line: u32,
    ) -> Result<(), String> {
        let obj_tmp = self.define_local("__late_get_obj");
        self.compile_expr(object)?;
        self.emit_u16(Op::LOCAL_SET, obj_tmp);

        self.emit_u16(Op::LOCAL_GET, obj_tmp);
        let type_key = self.str_const(CONTROL_TYPE_FIELD);
        self.emit_struct_field_op(Op::STRUCT_GET, 0, type_key);
        let undef_idx = self.import("wasm:js-undefined", "test");
        self.emit_host_call(undef_idx, 1);
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::LOCAL_GET, obj_tmp);
        let prop_key = self.str_const(&self.canon(prop));
        self.emit_struct_field_op(Op::STRUCT_GET, 0, prop_key);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, obj_tmp);
        self.emit_gui_property_get(&prop.to_ascii_lowercase(), line);
        self.chunk().emit_end(line);
        Ok(())
    }

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
        //
        // The ROLE comes from the CLASS, not from the source word. The frontend
        // declares `ClientWidth`→`width`, `Hint`→`tooltip`, `Menu`→`child` in
        // its own platform tree, and this is the write path that has to honour
        // it — lowercasing the Pascal spelling and calling it a role sent every
        // RENAMED property to `setAttribute("clientwidth", …)`, which no widget
        // reads, with no error to show for it.
        let prop = prop.to_ascii_lowercase();
        let role = self
            .declared_property_role(type_name, &prop, true)
            .unwrap_or_else(|| prop.clone());
        self.compile_expr(object)?;
        self.emit_u16(Op::LOCAL_GET, value_tmp);
        self.emit_gui_property_set(&role, line);
        self.emit(Op::DROP);

        // A designer `Name` is BOTH of HTML's two identifiers, which are not
        // the same thing: `id` is unique per document and is what
        // `getElementById` and `<label for>` resolve, while `name` is the
        // form-control submission key that `form.elements[…]` and
        // serialization read. A control that set only `id` would look right
        // and submit nothing, so set both.
        let form_associated =
            registered_control_element(&self.profile.namespaces.type_scopes, type_name)
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

    // There is deliberately no `emit_control_property_get` here. A property
    // READ already arrives as the `gui.prop_get.<role>` the platform tree
    // declared (`builtins.rs` strips the prefix and calls
    // `emit_gui_property_get`), so the role is canonical by construction and
    // needs no second resolution. The function that used to sit here had no
    // caller and carried a THIRD copy of `property_op` — one that never
    // learned the conversions `emit_gui_property_get` does, so it would have
    // read `Enabled` back as the raw `disabled` attribute instead of a boolean
    // and `Width` as `"780px"` instead of a number.

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
        self.emit_control_type_stamp(type_name, line);
        if element.establishes_containing_block() {
            self.emit_container_is_positioned(line);
        }
    }

    /// Stamp `__control_type` — "this object is a node in the document, and
    /// this is the control it is".
    ///
    /// Here, in the ONE place an element is created, rather than in a
    /// construction path: `New Button()` reaches
    /// `emit_tree_ctor_construction`, `Window.Forms.Button()` reaches
    /// `compile_call`, and a frontend that grows a third way would have needed
    /// a third copy. Stamping at creation means the answer exists for every
    /// element however it was spelled.
    ///
    /// It is what makes a LATE-BOUND member access decidable at all — see
    /// `emit_late_bound_property_set`. It is also read by the host
    /// (`platforms/vybe/src/gui.rs`, `gui_launch.rs`, `simd.rs`), which is why
    /// its value is the canonical control NAME (`Button`), not the tag.
    ///
    /// The element is on the stack and stays there; this consumes a copy.
    fn emit_control_type_stamp(&mut self, type_name: &str, line: u32) {
        let canonical = canonical_control_name(type_name);
        let stamped = if canonical.is_empty() {
            type_name.to_string()
        } else {
            canonical
        };
        let element = self.define_local("__gui_stamped_element");
        self.emit_u16(Op::LOCAL_TEE, element);
        self.emit_const(Value::String(Arc::from(stamped.as_str())));
        let key = self.str_const(CONTROL_TYPE_FIELD);
        self.emit_struct_field_op(Op::STRUCT_SET, 0, key);

        // A control has a NAME whether the program gives it one or not. Every
        // framework here guarantees it — WinForms' designer assigns `Button1`,
        // the VCL assigns `Button1` — and programs read it: the calculator does
        // `displayName := display.Name` and addresses the control by that
        // string afterwards. The `vybe:gui` factory generated one; a bare
        // `createElement` does not, so `.Name` answered `""` and every
        // by-name write went to a control called nothing.
        //
        // It is the element's `id`, which is what `getElementById` resolves —
        // the same role a program-assigned `Name` fills, so a later `Name = x`
        // simply replaces it.
        self.gui_auto_name_counter += 1;
        let auto_name = format!(
            "{}{}",
            stamped.to_ascii_lowercase(),
            self.gui_auto_name_counter
        );
        self.emit_u16(Op::LOCAL_GET, element);
        self.emit_const(Value::String(Arc::from(auto_name.as_str())));
        self.emit_gui_property_set("name", line);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, element);
    }

    /// `vybe.gui.setProperty(name, prop, value)` / `getProperty(name, prop)` —
    /// the BY-NAME surface, lowered onto the document.
    ///
    /// These are not internal plumbing: a program writes them itself, so they
    /// are part of the language surface and had to survive the move to the DOM.
    /// They used to reach `GuiState`, which the renderer no longer paints from
    /// once a document has content — so the calculator's every display update
    /// wrote to a registry nothing reads, while its clicks fired correctly.
    ///
    /// A name IS `getElementById`; that is what the `name` role has always
    /// meant. The property spelling must be a literal, because the ROLE decides
    /// the DOM operation — a computed one falls through to the old call rather
    /// than guess.
    pub(super) fn try_emit_gui_property_by_name(
        &mut self,
        module: &str,
        func: &str,
        args: &[&Expression],
    ) -> Result<bool, String> {
        if !module.eq_ignore_ascii_case(GUI_MODULE) {
            return Ok(false);
        }
        let setting = func.eq_ignore_ascii_case(HOST_FN_SET_PROPERTY_BY_NAME);
        let getting = func.eq_ignore_ascii_case(HOST_FN_GET_PROPERTY_BY_NAME);
        if !(setting || getting) {
            return Ok(false);
        }
        let expected = if setting { 3 } else { 2 };
        if args.len() != expected {
            return Ok(false);
        }
        let Some(role) = args.get(1).and_then(|arg| match &arg.kind {
            vybe_ast::ExprKind::Lit(vybe_ast::Literal::Str(prop)) => {
                Some(prop.to_ascii_lowercase())
            }
            _ => None,
        }) else {
            return Ok(false);
        };

        let line = self.line;
        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
        self.chunk().emit_call(doc_idx, 0, line);
        self.compile_expr(args[0])?;
        let by_id = self.import(DOM_MODULE, "getElementById");
        self.emit_host_call(by_id, 2);
        if setting {
            self.compile_expr(args[2])?;
            self.emit_gui_property_set(&role, line);
        } else {
            self.emit_gui_property_get(&role, line);
        }
        Ok(true)
    }

    /// Declare a container `position: absolute`.
    ///
    /// A child's `Left`/`Top` is relative to its parent in every framework here
    /// — VCL, WinForms and Flutter's `Positioned` all mean it. CSS spells that
    /// as an absolutely positioned box resolving against its nearest
    /// **positioned** ancestor, and a `position: static` box is not one. So a
    /// container has to say it is positioned, or a browser handed this document
    /// would place every nested control against the body.
    ///
    /// `absolute` rather than `relative`, and the difference is not cosmetic.
    /// Both are positioned, so both establish a containing block — but a
    /// container in these frameworks carries its OWN `Left`/`Top` too, and
    /// those two values mean different things per CSS: an absolute box is
    /// placed AT its coordinates, a relative box keeps its flow slot and is
    /// merely offset from it. Declaring `relative` said the second and meant
    /// the first, which only worked because the widget layer treated any
    /// positioned box with coordinates as out of flow. That conflation is what
    /// left `position: relative` with no way to mean what CSS says it means.
    ///
    /// It is DECLARED, into the document, rather than assumed by the widget
    /// layer. A behavioural assumption would render correctly here and wrongly
    /// in a real engine — which defeats the point of being HTML underneath. The
    /// emitted markup is now what you would write by hand:
    ///
    /// ```html
    /// <div style="position: absolute; left: 10px; top: 10px">
    ///   <button style="position: absolute; left: 20px; top: 50px">…</button>
    /// </div>
    /// ```
    ///
    /// The element is on the stack and stays there; this consumes a copy.
    fn emit_container_is_positioned(&mut self, line: u32) {
        // `LOCAL_TEE` stores and leaves the value on the stack, so the element
        // stays put for the caller while this reads a copy of it.
        let element = self.define_local("__gui_container");
        self.emit_u16(Op::LOCAL_TEE, element);
        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
        self.chunk().emit_call(doc_idx, 0, line);
        self.emit_u16(Op::LOCAL_GET, element);
        emit_string_const(self.chunk(), "position", line);
        emit_string_const(self.chunk(), "absolute", line);
        let set_idx = self.import(CSSOM_MODULE, "setStyleProperty");
        self.emit_host_call(set_idx, 4);
        // A host call leaves its result on the stack — the convention every
        // `gui::emit_*` helper follows. `emit_control_element` promises to
        // leave exactly the element, so the result is dropped here. Without
        // this the extra value corrupts the rest of the constructor: the
        // calculator lost all twenty of its form-parented buttons and the form
        // collapsed to the height of its one panel.
        self.chunk().emit_op(Op::DROP, line);
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
