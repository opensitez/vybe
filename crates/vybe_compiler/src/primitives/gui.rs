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
//! C# walker  ───┤   .NET surface
//! F# walker  ───┴──> dotnet.rs ──┐
//!                                │
//! Dart walker ─────> flutter.rs ─┼──> compiler_common::gui ──> web:dom / web:cssom
//!                                │
//! Python walker ───> tkinter.rs ─┘
//! ```
//!
//! All frontends produce the SAME bytecode for the same canonical operation.
//! Switching the host's GUI backend (or running on a non-Vybe VM with a
//! different GUI binding) requires no compiler changes.

use crate::primitives::class_slots;
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

/// WHATWG DOM — where a control is actually created.
pub const DOM_MODULE: &str = "web:dom";
/// WHATWG HTML — `document`, and the element IDL properties.
pub const DOCUMENT_MODULE: &str = "web:html";
pub const HOST_FN_CREATE_ELEMENT: &str = "createElement";
/// `window.document` of the current browsing context.
pub const HOST_FN_ACTIVE_DOCUMENT: &str = "activeDocument";
/// CSSOM — `element.style`.
pub const CSSOM_MODULE: &str = "web:cssom";

/// The live element type a control's ancestry must end in to carry a real rtt.
///
/// Declared in the tree vocabulary (`vybe_runtime::namespaces`) because it is a
/// PLATFORM's declaration and the platform crates cannot see this one.
pub use vybe_runtime::namespaces::DOM_ELEMENT_TYPE;

// The live statement of what a control IS lives below, in the emit helpers and
// the property-role tables. Those emit to `web:*`.


// ─── Emit helpers ────────────────────────────────────────────────────────────
//
// Canonical patterns. Every language frontend uses these directly or via a
// framework-specific resolver (`dotnet.rs`, etc.) that calls these.
//
// All emit functions are pure WASM bytecode + standard host imports — no
// custom opcodes, no language-specific knowledge. They define the calling
// convention so call sites are uniform across compilers.
//
// IMPORTANT: an import must be registered on the chunk that EMITS the call.
// `Compiler::import()` does exactly that (`chunks[self.current].add_import`).
// Registering on chunks[0] instead yields an index out of range of the current
// chunk's table, which the normalize pass's script-table fallback resolves
// correctly only by luck — see `dispatch.rs`'s `object.new` arm, where it
// silently mis-resolved to `js-string.concat` in nested function-expression
// contexts. Callers pass a pre-resolved `import_idx`, which keeps gui.rs
// chunk-agnostic — it does not need to know which chunk it is emitting into.
//
// (This comment previously stated the opposite rule and described `import()`
// as delegating to `chunks[0]`. It does not, and the chunk-0 form is the bug.)

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
/// VCL menus are the case that needs it: `FMainMenu.Items.Add(MenuFile)`. A
/// menu that does not enter the DOCUMENT cannot render, be hit-tested, or be
/// listed by `widgets`.
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

/// `Application.Run` — and it emits NOTHING, deliberately.
///
/// A document is not told to run. It runs because it HAS content, which is the
/// condition the launch gate already reads, so there is no DOM call left for
/// this to lower to. Emitting nothing is the honest answer; the alternative was
/// a host function whose entire effect was setting a `should_run` flag on a
/// host that is being deleted (guiplan.md, "There is no `runApplication`, for
/// anybody").
///
/// It stays a NAMED emit rather than being dropped from the tree because the
/// frontends still spell it — `Application.Run`, `Application.Terminate` — and a
/// declared leaf that answers "nothing" is what stops each of them inventing a
/// private answer.
pub const APP_RUN_EMIT: &str = "gui.app.run";


/// `Application.Terminate` — closes the browsing context.
pub const APP_EXIT_EMIT: &str = "gui.app.exit";

/// The DOM operation a property role IS. `(module, func, attribute-key)`.
///
/// The roles are the canonical property names every language already lowers
/// to — a shared vocabulary, not per-framework spellings. That is why dotnet
/// needs no mapping of its own: it emits these names already.
///
/// Pascal never learns any of this. It calls with the same intent it always
/// had; `widgets` is HTML underneath, which is not its business.
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
        // **A control that is BORN WITH CHILDREN does not paint its `Text`.**
        //
        // `textContent =` REPLACES a node's children (DOM §4.4 "string replace
        // all"), so a composite control whose chrome is its subtree loses that
        // subtree the moment anyone writes its caption — and a designer file
        // writes one on every control it generates. A `BindingNavigator` built
        // its five standard items and then became the single text node
        // `bnav1`; so did a `SplitContainer` and its two panes.
        //
        // Not painting it is also what the control does: WinForms' navigator,
        // split container and month calendar all INHERIT `Text` from `Control`
        // and none of them draws it — `widgets` agrees, showing `0 of 0`
        // where the caption would be. So the write is kept, off the text node
        // and on an attribute, where a property with no visual counterpart
        // belongs and where it still round-trips.
        //
        // ⚠ The LATE-BOUND read path (`emit_late_bound_property_get`) resolves
        // its role from the property name at runtime, with no type to consult,
        // so a `.Text` read reached that way still reads `textContent` and
        // answers with the chrome's own words. A typed read is correct.
        "unpaintedtext" => (
            DOM_MODULE,
            if setting {
                "setAttribute"
            } else {
                "getAttribute"
            },
            Some("data-text"),
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
        // A FONT IS A STYLE PROPERTY, reached exactly like `left` or
        // `backcolor`. It was the one role named as explicitly unmapped, so
        // `X.Font.Size := 16` became `setAttribute("font", …)` and reached
        // nothing — which is also why the CSS inheritance in `widgets`
        // measured neutral: `font_family`/`font_size`/`font_weight`/
        // `font_style` are all in its inherited set, and every `.dfm` in the
        // corpus declares `Font.Name`, so nothing could exercise it.
        //
        // The value is a Font OBJECT, so the write composes CSS's own `font`
        // shorthand from it (`italic bold 12px Arial`) — the same shape as the
        // `columncount` arm building `repeat(N, 1fr)`. One property carries all
        // four axes, so the cascade inherits them together as CSS specifies.
        "font" => (
            CSSOM_MODULE,
            if setting {
                "setStyleProperty"
            } else {
                "getStyleProperty"
            },
            Some("font"),
        ),
        // Box spacing, the pointer, and the background image — all CSS
        // properties CSS already names, reached exactly like `backcolor`.
        // Every one was declared by the frontends and mapped by nobody, so it
        // fell through to `setAttribute` under its toolkit spelling.
        //
        // ⚠ `padding`/`margin` take a LENGTH and so join the `px` list below:
        // `Padding := 8` is `padding: 8px`, and a browser drops `"8"`.
        "padding" => (
            CSSOM_MODULE,
            if setting {
                "setStyleProperty"
            } else {
                "getStyleProperty"
            },
            Some("padding"),
        ),
        "margin" => (
            CSSOM_MODULE,
            if setting {
                "setStyleProperty"
            } else {
                "getStyleProperty"
            },
            Some("margin"),
        ),
        "cursor" => (
            CSSOM_MODULE,
            if setting {
                "setStyleProperty"
            } else {
                "getStyleProperty"
            },
            Some("cursor"),
        ),
        "backgroundimage" => (
            CSSOM_MODULE,
            if setting {
                "setStyleProperty"
            } else {
                "getStyleProperty"
            },
            Some("background-image"),
        ),
        // `TabIndex` is an HTML ATTRIBUTE, not a style — `tabindex` is how the
        // document orders focus, which is exactly what the toolkit means.
        "tabindex" => (
            DOM_MODULE,
            if setting {
                "setAttribute"
            } else {
                "getAttribute"
            },
            Some("tabindex"),
        ),
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
        // **A table layout IS a grid.** `TableLayoutPanel.ColumnCount` (WinForms)
        // and `TGridPanel`'s column count (VCL) mean exactly
        // `grid-template-columns`, and unmapped they fell to the attribute
        // fallback: `columncount="7"` landed on the element and NOTHING read it,
        // so a 7x6 board rendered at whatever the widget was constructed with.
        // The attribute fallback is right for decorative unknowns; a layout
        // property with an exact CSS counterpart is not one.
        //
        // The count becomes `repeat(n, 1fr)` — equal tracks, which is what a
        // table layout with no explicit column styles is. Explicit styles
        // (`ColumnStyles.Add(Percent, …)`) refine it to real tracks later and
        // overwrite this, exactly as a later CSS declaration would.
        "columncount" => (
            CSSOM_MODULE,
            if setting {
                "setStyleProperty"
            } else {
                "getStyleProperty"
            },
            Some("grid-template-columns"),
        ),
        "rowcount" => (
            CSSOM_MODULE,
            if setting {
                "setStyleProperty"
            } else {
                "getStyleProperty"
            },
            Some("grid-template-rows"),
        ),
        // `dock` joins these because it is geometry too, just expressed as a
        // rule instead of a number: the container computes the rect from it.
        // A frontend that spells it `Align` (VCL) or `Dock` (WinForms) reaches
        // the same style property, and `widgets` owns the result.
        // ⚠ The READ is `getComputedStyle`, not `element.style`. A frontend
        // asking for `Left` wants the pixel the control OCCUPIES, which is a
        // resolved value; `element.style.getPropertyValue` answers what was
        // declared, so `Left := 0` followed by a layout would read back `0`
        // forever. Writing is unchanged — a write always sets a declaration.
        "left" | "top" | "width" | "height" => (
            CSSOM_MODULE,
            if setting {
                "setStyleProperty"
            } else {
                "getComputedStyleProperty"
            },
            Some(""),
        ),
        // `dock` is geometry too, but expressed as a RULE rather than a number:
        // the container computes a rect from it. So it splits from the four
        // above — there is no resolved `dock` to read, and asking the computed
        // style for one would answer with a value that has no meaning. A
        // frontend spelling it `Align` (VCL) or `Dock` (WinForms) reaches the
        // same declaration, and `widgets` owns the result.
        "dock" => (
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
/// The field names are the ones the value type actually STORES, not the
/// framework's property spelling — `Point` declares `X`/`Y` and stores
/// `x`/`y`.
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
        nest_coerce: Option<&str>,
        line: u32,
    ) -> Result<(), String> {
        use vybe_runtime::namespaces::FieldGui;
        match field {
            // A child widget nests; a scalar sets the property. Both spellings
            // reach an operation that already exists, so neither needs a new
            // one.
            // ⚠ The NEST half of this used to be missing entirely — the comment
            // above described both operations and the code only ever emitted the
            // property one. So `MaterialApp(home: CalculatorPage())` set an
            // ATTRIBUTE from a widget, which stringifies: the document came out
            // as `<flowlayoutpanel home="[object]">` with the whole page below
            // the root collapsed into seven characters, and one control on
            // screen.
            //
            // The test is `ref.test` against the DOM element type, not a probe
            // for a marker field. A marker read is answered through the element
            // attribute path — that is the same failure that made `_vfConcrete`
            // over-run — whereas the rtt is what the object IS. It only became
            // askable once a control was allocated with its own declared type;
            // before that every element shared one type id and this question had
            // no honest answer.
            //
            // `emit_is_instance_of` rather than a bare `REF_TEST`: it unions the
            // rtt with the `__types` chain, so a frontend that has NOT declared
            // the DOM tail still answers correctly from strings and keeps
            // working.
            FieldGui::NestOrProp(key) => {
                let role = key.to_ascii_lowercase();
                let value = self.define_local("__gui_nest_value");
                let parent = self.define_local("__gui_nest_parent");
                self.emit_u16(Op::LOCAL_SET, value);
                self.emit_u16(Op::LOCAL_SET, parent);
                self.emit_gui_nest_coerce(nest_coerce, value, line)?;

                let current = self.current;
                crate::primitives::reflection::emit_is_instance_of(
                    &mut self.chunks,
                    current,
                    value,
                    vybe_runtime::namespaces::DOM_ELEMENT_TYPE,
                    line,
                );
                // **Every arm of this match must leave the stack EXACTLY as it
                // found it.** `appendChild` and the property setters are host
                // calls that answer a value, and `Children` answers none — so
                // without these drops the effect differed per field, and a
                // constructor applied N fields left N values behind.
                //
                // Harmless while a constructor was a STATEMENT. Fatal in
                // expression position: `[Text('a'), Text('b')]` builds its
                // elements onto the stack for `array.new_fixed`, which then took
                // the residue instead of the widgets. Bound to a variable first
                // it worked, which is exactly what made it look like a list bug
                // rather than a stack one.
                self.chunk().emit_if(line);
                self.emit_u16(Op::LOCAL_GET, parent);
                self.emit_u16(Op::LOCAL_GET, value);
                self.emit_gui_append_child(line);
                self.emit(Op::DROP);
                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, parent);
                self.emit_u16(Op::LOCAL_GET, value);
                self.emit_gui_property_set(&role, line);
                self.emit(Op::DROP);
                self.chunk().emit_end(line);
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
                let item = self.define_local("__gui_children_item");
                self.emit_u16(Op::LOCAL_SET, list);
                self.emit_u16(Op::LOCAL_SET, parent);

                // The length is asked the way `.length` is asked EVERYWHERE
                // else — `collections::emit_len`, which dispatches on the shape
                // at run time.
                //
                // The raw `array.len` opcode was here, and it answers 0 for any
                // value that is not a packed `ObjectKind::Array` — no trap, no
                // diagnostic, just zero. So a `children:` list that arrived in
                // any other collection shape made this loop run zero times:
                // every child widget was constructed, sized and laid out, and
                // then attached to nothing. The document held a Scaffold with an
                // empty Column while 70+ live elements sat orphaned beside it.
                //
                // A widget that was given no list at all has NO children, and
                // that is not an error — `AppBar()` without `actions:` is
                // ordinary. `emit_len`'s last fallback is `Object.keys(value)`,
                // which THROWS on null, so the absent case has to be answered
                // before asking.
                let current = self.current;
                self.emit_u16(Op::LOCAL_GET, list);
                self.emit(Op::REF_IS_NULL);
                self.chunk().emit_if_value(line);
                self.emit_const(Value::I32(0));
                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, list);
                collections::emit_len(&mut self.chunks, current, line);
                self.chunk().emit_end(line);
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

                // Coerce per ITEM, not per list: a `children:` list mixes
                // concrete widgets with composites, and only the composite needs
                // asking.
                self.emit_u16(Op::LOCAL_GET, list);
                self.emit_u16(Op::LOCAL_GET, index);
                collections::emit_get(&mut self.chunks, current, line);
                self.emit_u16(Op::LOCAL_SET, item);
                self.emit_gui_nest_coerce(nest_coerce, item, line)?;
                self.emit_u16(Op::LOCAL_GET, parent);
                self.emit_u16(Op::LOCAL_GET, item);
                self.emit_gui_append_child(line);
                // **Inside the loop, once per child.** `appendChild` answers a
                // value, so a list of N children leaves N of them — this is not
                // one residue per ARM, it is one per ITERATION. Dropping once
                // outside the loop would under-drop by N-1, which is how a
                // constructor in expression position ended up feeding
                // `array.new_fixed` the leftovers instead of its widgets.
                self.emit(Op::DROP);

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
                self.emit(Op::DROP);
            }
            // The child's text IS this control's caption, not a nested control.
            FieldGui::Caption => {
                self.emit_gui_property_set("text", line);
                self.emit(Op::DROP);
            }
        }
        Ok(())
    }

    /// Replace `slot` with the node the value in it contributes, per the
    /// platform's [`CtorSpec::nest_coerce`]. No-op when the platform declares
    /// none — a control that IS its element has nothing to answer.
    ///
    /// The coercion is a plain call to a guest function of one argument, so the
    /// platform authors it in the target language and the shared path stays
    /// free of any framework's inflation rules. It must be TOTAL: it is applied
    /// to every nested value, including the scalars that end up as properties,
    /// so anything it does not recognise it returns unchanged.
    ///
    /// A coercion the program does not contain is skipped, and that is not a
    /// silent shim — it is decidable. The function is part of the platform's
    /// render runtime, which a frontend injects only into a program that
    /// RENDERS. Absent ⟹ the program never attaches anything to a document ⟹
    /// no nesting it performs is observable. What would be silent is the
    /// opposite: coercing when the runtime is present but the widget's platform
    /// forgot to declare it, which is why declaring it is the platform's job and
    /// not something inferred here.
    fn emit_gui_nest_coerce(
        &mut self,
        nest_coerce: Option<&str>,
        slot: u16,
        line: u32,
    ) -> Result<(), String> {
        let Some(name) = nest_coerce else {
            return Ok(());
        };
        let Some(chunk_idx) = self.chunks.iter().position(|c| c.name == name) else {
            return Ok(());
        };
        self.emit_u16(Op::REF_FUNC, chunk_idx as u16);
        self.chunk().emit(0u8, line); // upvalue count
        self.emit_u16(Op::LOCAL_GET, slot);
        crate::primitives::callable::emit_direct_invoke_chunk(self.chunk(), 1, line);
        self.emit_u16(Op::LOCAL_SET, slot);
        Ok(())
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

        // A FORM's own verbs. `HTMLDialogElement` names all three outright, so
        // there is nothing to invent: `ShowModal` IS `showModal()`, `Close` IS
        // `close()`, and the modal result is `returnValue`.
        //
        // These are the verbs a toolkit has and a plain control does not, which
        // is why they are their own arm rather than more `visible` writes:
        // showing a form MODALLY is a statement about input to everything else,
        // and no attribute on the element expresses that.
        //
        // ⚠ `show`/`hide` above stay the `visible` role — a control appearing is
        // not a dialog opening, and collapsing the two would make `Button.Show`
        // try to open a dialog.
        if let Some(func) = match verb {
            "show_modal" => Some("showModal"),
            "show_form" => Some("show"),
            "close" => Some("close"),
            _ => None,
        } {
            self.chunk().emit_call(doc_idx, 0, line);
            self.emit_u16(Op::LOCAL_GET, ctrl);
            let idx = self.import(DOCUMENT_MODULE, func);
            self.emit_host_call(idx, 2);
            self.emit_null();
            return;
        }

        // `Dispose` — DESTROY the control, which in a document means detaching
        // the node. `ChildNode.remove()` (DOM §4.2.9) is exactly that, and it
        // takes the handlers with it: listeners live on the node, so a removed
        // node stops receiving input without anything hunting them down.
        //
        // ⚠ This is the verb that must NOT be confused with `Hide`. Hiding is
        // the `visible` role — the node stays in the tree and can be shown
        // again, which is what a multi-window application relies on. Disposing
        // is final: there is nothing left to show. The old host fn conflated
        // the two (it set `Visible=false` and dropped handlers), so a disposed
        // control could be resurrected by writing `Visible` back.
        //
        // A node with no parent removes nothing and does not raise, per spec —
        // so disposing twice is safe without a guard here.
        if verb == "dispose" {
            self.chunk().emit_call(doc_idx, 0, line);
            self.emit_u16(Op::LOCAL_GET, ctrl);
            let idx = self.import(DOM_MODULE, "remove");
            self.emit_host_call(idx, 2);
            return;
        }

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
            // and `widgets` is right to store it that way. A control's
            // `Left` is a number, so parse the unit back off here.
            "left" | "top" | "width" | "height" => {
                let parse_float = self.import("ecma:number", "parseFloat");
                self.emit_host_call(parse_float, 1);
            }
            // The inverse of the write: the store holds the EXPANDED track
            // list (`widgets` normalises `repeat(7, 1fr)` on the way in),
            // so the count is how many tracks there are. Splitting on a space
            // and taking the length is the same trick `Lines` uses, and it
            // makes the round trip an identity — write 7, read 7.
            "columncount" | "rowcount" => {
                emit_string_const(self.chunk(), " ", line);
                strings::emit_split(self.chunk(), line);
                strings::emit_length(self.chunk(), line);
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

    /// `RemoveHandler ctrl.Click, AddressOf H` IS `removeEventListener` —
    /// stack in `[control, handler]`, out `[_]`.
    ///
    /// The counterpart of the `on<type>` branch in `emit_gui_property_set`.
    /// ⚠ It did not exist for a while: `compile_remove_handler_stmt` had no
    /// element branch at all, so subscribing reached the document and
    /// UNsubscribing wrote to a registry nothing read. Removing a handler
    /// quietly did nothing.
    ///
    /// The handler is passed UNBOUND, deliberately. `addEventListener` stores
    /// the wrapper `bind` produced, which the program has never seen and cannot
    /// name; the host matches a bound listener by delegate equality precisely
    /// for this case. Re-binding here would build a third object equal to
    /// neither.
    pub fn emit_remove_event_listener(&mut self, event: &str, line: u32) {
        let handler = self.define_local("__gui_event_handler");
        let ctrl = self.define_local("__gui_event_target");
        self.emit_u16(Op::LOCAL_SET, handler);
        self.emit_u16(Op::LOCAL_SET, ctrl);
        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
        self.chunk().emit_call(doc_idx, 0, line);
        self.emit_u16(Op::LOCAL_GET, ctrl);
        emit_string_const(self.chunk(), &event.to_ascii_lowercase(), line);
        self.emit_u16(Op::LOCAL_GET, handler);
        let idx = self.import(DOM_MODULE, "removeEventListener");
        self.emit_host_call(idx, 4);
    }

    /// Lower `gui.prop_set.<role>` — stack in `[control, value]`, out `[_]`.
    pub fn emit_gui_property_set(&mut self, role: &str, line: u32) {
        // The WINDOW title is the DOCUMENT's, not an element's.
        //
        // It is the one role whose target is not the control it was written on:
        // `MaterialApp(title:)`, a VCL form's `Caption` and a WinForms form's
        // `Text` all name the string the window manager shows, and HTML holds
        // exactly one of those — `document.title`, the `<title>` in the head.
        // Routed through the ordinary attribute path it became `title=""` on
        // the element, which is the hover TOOLTIP: the value was written, was
        // readable, and appeared nowhere a user looks.
        if role == "windowtitle" {
            let value = self.define_local("__gui_window_title");
            self.emit_u16(Op::LOCAL_SET, value);
            self.emit(Op::DROP); // the control — the title is not its property
            let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
            self.chunk().emit_call(doc_idx, 0, line);
            self.emit_u16(Op::LOCAL_GET, value);
            let set_idx = self.import(DOCUMENT_MODULE, "setTitle");
            self.emit_host_call(set_idx, 2);
            return;
        }
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
                let key = crate::primitives::class_slots::resolve_interned(
                    self.chunk(),
                    &crate::primitives::class_slots::ClassSlot::internal(field),
                    &crate::primitives::class_slots::PlainNames,
                );
                crate::primitives::class_slots::emit_class_get(
                    self.chunk(),
                    crate::primitives::class_slots::ObjSource::Stack,
                    &key,
                    crate::primitives::class_slots::Dest::Stack,
                    line,
                );
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
        // **A box placed by COORDINATES says so.** `left`/`top` are inert on a
        // `position: static` box in CSS, so a frontend writing them already
        // means `absolute` — it simply had not said it. Declaring it here, at
        // the write, is what lets containers stop claiming it at birth: a
        // control that IS positioned becomes a containing block for its own
        // children, and one that is not stays in flow.
        //
        // Written before the coordinate itself so the box is already positioned
        // when the value lands, and only for the two roles that mean placement
        // — `width`/`height` are a size, not a position.
        if matches!(role, "left" | "top") {
            let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
            self.chunk().emit_call(doc_idx, 0, line);
            self.emit_u16(Op::LOCAL_GET, ctrl);
            emit_string_const(self.chunk(), "position", line);
            emit_string_const(self.chunk(), "absolute", line);
            let set_idx = self.import(CSSOM_MODULE, "setStyleProperty");
            self.emit_host_call(set_idx, 4);
            self.emit(Op::DROP);
        }
        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
        self.chunk().emit_call(doc_idx, 0, line);
        self.emit_u16(Op::LOCAL_GET, ctrl);
        let argc = match key {
            Some(k) => {
                emit_string_const(self.chunk(), if k.is_empty() { role } else { k }, line);
                // A track COUNT is not a track list. `ColumnCount := 7` means
                // seven equal columns, which CSS spells `repeat(7, 1fr)` — and
                // the count is a runtime value, so the string is built here
                // rather than folded. Same shape as the `px` suffix below, for
                // the same reason: the DOM takes CSS text, not a toolkit
                // number.
                if matches!(role, "columncount" | "rowcount") {
                    emit_string_const(self.chunk(), "repeat(", line);
                    self.emit_u16(Op::LOCAL_GET, value);
                    strings::emit_to_string(self.chunk(), line);
                    ops::emit_dyn_add(self.chunk(), line);
                    emit_string_const(self.chunk(), ", 1fr)", line);
                    ops::emit_dyn_add(self.chunk(), line);
                    let idx = self.import(module, func);
                    self.emit_host_call(idx, 4);
                    return;
                }
                // `Font` arrives as the value OBJECT the language built
                // (`{name, size, bold, italic}`) and leaves as CSS's own
                // shorthand, `<style> <weight> <size>px <family>`. Composed
                // here for the same reason `columncount` composes
                // `repeat(N, 1fr)`: the DOM takes CSS TEXT, and the fields are
                // runtime values, so there is nothing to fold.
                if role == "font" {
                    let font = self.define_local("__gui_font");
                    self.emit_u16(Op::LOCAL_SET, font);
                    for (field, css) in [("italic", "italic "), ("bold", "bold ")] {
                        self.emit_u16(Op::LOCAL_GET, font);
                        self.class_get(class_slots::ObjSource::Stack, &class_slots::ClassSlot::internal(field));
                        ops::emit_dyn_to_bool(self.chunk(), line);
                        self.chunk().emit_if_value(line);
                        emit_string_const(self.chunk(), css, line);
                        self.chunk().emit_else(line);
                        emit_string_const(self.chunk(), "", line);
                        self.chunk().emit_end(line);
                    }
                    ops::emit_dyn_add(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_GET, font);
                    self.class_get(class_slots::ObjSource::Stack, &class_slots::ClassSlot::internal("size"));
                    strings::emit_to_string(self.chunk(), line);
                    ops::emit_dyn_add(self.chunk(), line);
                    emit_string_const(self.chunk(), "px ", line);
                    ops::emit_dyn_add(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_GET, font);
                    self.class_get(class_slots::ObjSource::Stack, &class_slots::ClassSlot::internal("name"));
                    ops::emit_dyn_add(self.chunk(), line);
                    let idx = self.import(module, func);
                    self.emit_host_call(idx, 4);
                    return;
                }
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
                // `widgets`' own lenient `parse_px` hid it. That is a
                // defect no capture can show, because both ends of OUR
                // pipeline agreed on the wrong thing.
                //
                // `dock` shares this arm and must NOT get a unit: it carries
                // an edge keyword, not a length.
                //
                // **A CSS property value is TEXT, whatever it started as.**
                // `setStyleProperty` takes a string, and a value that is not
                // one has to be asked what it says — which is the guest's own
                // `to_string`, so a value type answers with its own spelling.
                //
                // This used to run only for the LENGTH roles, as a side effect
                // of needing somewhere to append `px`. Every other property
                // handed the host a raw object, and the host formatted it the
                // only way it can from outside the guest: the literal
                // `[object]`. That is why `EdgeInsets` reached CSS correctly
                // and `Alignment` did not — one is a length and the other is
                // not, and nothing else about them differs.
                //
                // Only for the CSSOM. The other modules on this path take
                // typed arguments — `enabled`/`visible` are negated to an i32
                // just above — and stringifying those would break them.
                if module == CSSOM_MODULE {
                    strings::emit_to_string(self.chunk(), line);
                }
                if matches!(
                    role,
                    "left" | "top" | "width" | "height" | "padding" | "margin"
                ) {
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
    /// CSS the control is BORN with — `Row` is a `div` that declares
    /// `display: flex; flex-direction: row`.
    ///
    /// The element alone cannot say this: `Panel`, `Row` and `Column` are all
    /// `div`, and what separates them is a display mode, which is a STYLE and
    /// not a tag. guiplan's conversion table has carried a "Declares" column
    /// for exactly this and there was no mechanism behind it, so the only way
    /// to express a flex container was to invent a `vybe-*` tag — a pseudo-tag
    /// naming no widget kind, which renders as a 120x20 label.
    ///
    /// Declared CSS, not a layout flag: it lands through the same
    /// `setStyleProperty` a program would use, so it cascades, serializes into
    /// the `style` attribute, and a browser would do the same thing with it.
    pub declares: Vec<(String, String)>,
    /// Content ATTRIBUTES the control is born with, written `@name=value` in
    /// the declaration (`@multiple` alone for a boolean one).
    ///
    /// CSS could not express these and they are not children, so neither of the
    /// other two channels reached them. The case that forced it: HTML's list
    /// box IS `<select>` — with `size` above one or `multiple`, which is the
    /// only thing separating a list from a dropdown (HTML §4.10.7). `ListBox`
    /// was a `<ul>`, an element with no selection model at all, which is why it
    /// rendered its items and could not select one while `ComboBox` — the same
    /// control, one attribute apart — worked.
    pub attributes: Vec<(String, String)>,
    /// The children the control is BORN with — see `CtorSpec::inner_html`.
    ///
    /// Not parsed out of the declaration string beside the CSS: markup is full
    /// of `;` and `:`, which are exactly the two characters that grammar splits
    /// on, so it travels in its own field instead of behind an escape.
    pub inner_html: Option<String>,
}

impl ControlElement {
    /// Parse a platform's declaration of what its control is.
    ///
    /// A platform declares the ELEMENT (`"button"`, `"input:checkbox"`,
    /// `"body"`) because it owns the vocabulary — plib knows `TEdit` is a text
    /// input, and nothing in a shared crate should have to. That is the whole
    /// point: no per-language table lives here.
    ///
    /// `;` separates the element from the CSS it is born with, and each
    /// declaration is ordinary `prop: value`:
    ///
    /// ```text
    /// "button"                                  a plain element
    /// "input:checkbox"                          element + input type
    /// "div;display:flex;flex-direction:column"  element + declared CSS
    /// ```
    ///
    /// The `:` split for an input type happens on the FIRST segment only, so a
    /// value containing a colon cannot be mistaken for one.
    fn parse(decl: &str) -> ControlElement {
        let mut parts = decl.split(';');
        let head = parts.next().unwrap_or("");
        let (tag, input_type) = head.split_once(':').unwrap_or((head, ""));
        // `@name=value` is a content ATTRIBUTE, anything else a CSS declaration.
        // `@` cannot begin a property name, so the two never collide, and a
        // bare `@name` is a boolean attribute — present, empty value, which is
        // how HTML spells `multiple` and `disabled`.
        let (attribute_parts, style_parts): (Vec<&str>, Vec<&str>) =
            parts.partition(|part| part.trim_start().starts_with('@'));
        let attributes = attribute_parts
            .into_iter()
            .map(|attribute| {
                let attribute = attribute.trim().trim_start_matches('@');
                let (name, value) = attribute.split_once('=').unwrap_or((attribute, ""));
                (name.trim().to_ascii_lowercase(), value.trim().to_string())
            })
            .collect();
        let declares = style_parts
            .into_iter()
            .filter_map(|d| d.split_once(':'))
            .map(|(prop, value)| {
                (
                    prop.trim().to_ascii_lowercase(),
                    value.trim().to_ascii_lowercase(),
                )
            })
            .collect();
        ControlElement {
            tag: tag.trim().to_ascii_lowercase(),
            input_type: input_type.trim().to_ascii_lowercase(),
            declares,
            attributes,
            inner_html: None,
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
            declares: Vec::new(),
            attributes: Vec::new(),
            inner_html: None,
        }
    }
}

/// What the REGISTRY says this type's control is — the same authority
/// `is_framework_control_parent` consults, so the two can never disagree.
pub fn registered_control_element(
    type_scopes: &[String],
    type_name: &str,
    fold: vybe_runtime::namespaces::Fold,
) -> Option<ControlElement> {
    let spec = vybe_runtime::namespaces::lookup_type_ctor_spec(type_scopes, type_name, fold)?;
    let inner_html = spec.inner_html.clone();
    let decl = spec.control_fn?;
    let mut element = if decl.starts_with("new_") {
        ControlElement::custom(type_name)
    } else {
        ControlElement::parse(&decl)
    };
    element.inner_html = inner_html;
    Some(element)
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
        // **A user's own class shadows a platform type of the same name.**
        //
        // The construction path has always said so; this one did not, and the
        // asymmetry is the bug: `Class Point` constructed correctly as a user
        // object and then had its member READS routed to the GUI property axis,
        // because `Point` is also a registered platform type. Every instance
        // then answered from the same host getter, so `a.X` and `b.X` read the
        // same value — the objects were never aliased, the reads were.
        //
        // The shadow skips the DIRECT lookup only, never the parent walk.
        // `TForm1 = class(TForm)` is user-declared too and must still route:
        // it is not a control because of its own name but because of what it
        // derives from, and that is exactly the difference between the two
        // cases. Bailing outright here would take every Pascal form with it.
        // `shadows_builtin_type` is THE user-declaration question — see its doc.
        // A raw `defined_classes` probe here answers differently from the other
        // shadow sites, which is the split that let `Class Point` win in one
        // path and lose in another.
        //
        // The shadow must be asked with the SAME fold as the lookup it guards.
        // `registered_control_element` resolves through the namespace tree,
        // which matches case-insensitively — but callers reach here with a
        // `normalize_type_hint`-lowercased spelling, so in a case-SENSITIVE
        // language `shadows_builtin_type("label")` cannot see the user's
        // `class Label` and the registry hit went unguarded: the ctor's
        // `this.text = text` compiled as a DOM property write and the field
        // stayed null. One lookup, one fold, on both sides of the guard.
        let user_owns_spelling = self.user_owns_type_spelling(type_name);
        if !user_owns_spelling {
            if let Some(element) =
                registered_control_element(
                    &self.profile.namespaces.type_scopes,
                    type_name,
                    self.tree_fold(),
                )
            {
                return Some(element);
            }
        }
        // Walk the declared parents. Bounded by the chain itself, and a cycle
        // in it would already have broken construction long before here.
        let mut current = self.pending_class_parent(type_name);
        while let Some(parent) = current {
            if let Some(element) =
                registered_control_element(
                    &self.profile.namespaces.type_scopes,
                    &parent,
                    self.tree_fold(),
                )
            {
                return Some(element);
            }
            current = self.pending_class_parent(&parent);
        }
        // **A control is a control by NAME too — the same answer construction
        // gives.**
        //
        // `emit_control_element` resolves the registry and falls back to
        // `ControlElement::custom`, so an unregistered control still becomes
        // `<vybe-button>`. This function stopped at `None`, so the very same
        // type constructed as an element and then was not recognised as one
        // for its events or its properties. One fact, two answers.
        //
        // The gap between them — a type with a control NAME but no registered
        // element — was where a separate host-side event path used to live, and
        // it was not a fallback worth preserving: it unsubscribed against a
        // registry nothing wrote to.
        //
        // WinForms is the case that made it visible. Those classes are
        // adapters, the same way the VB ones are, and they reach the DOM
        // through here.
        //
        // `canonical_control_name` is the gate, so this claims nothing about
        // an arbitrary type — only about names the shared table already calls
        // controls. The user shadow still applies: `Class Button` is the
        // user's class, not a control.
        // Same predicate as the direct lookup above — `user_owns_spelling`,
        // not a second probe. Asking the shadow question two ways here is the
        // very split that let `Class Point` win in one path and lose in
        // another — and `canonical_control_name` folds case just like the
        // registry, so this arm needs the same case-blind shadow.
        if !user_owns_spelling && !canonical_control_name(type_name).is_empty() {
            return Some(ControlElement::custom(type_name));
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
        // A method call on a VALUE is never a construction. The receiver of
        // `new Tag("core").Label()` flattens to nothing (only name chains
        // flatten), so without this the receiver evaporates and the method
        // name masquerades as a bare `Label(...)` — which then constructs an
        // element out of an instance method call.
        if let vybe_ast::ExprKind::Member { object, .. } = &callee.kind {
            if self.flatten_member_chain(object).is_empty() {
                return None;
            }
        }
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
        // Same predicate as every other shadow site. A declared FUNCTION of the
        // name also wins here — a call is being classified, so a user function
        // owns the callee spelling as much as a user class owns a type name.
        if self.shadows_builtin_type(last) || self.defined_functions.contains(&canon_last) {
            return None;
        }
        vybe_runtime::namespaces::is_registered_type(&self.profile.namespaces.type_scopes, last, self.tree_fold())
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
        let fold = self.tree_fold();
        let declared = |name: &str| {
            let target = if setting {
                vybe_runtime::namespaces::lookup_type_property_setter_target(scopes, name, prop, fold)
            } else {
                vybe_runtime::namespaces::lookup_type_property_target(scopes, name, prop, fold)
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
        // A USER-DECLARED class owns its own members: a platform type that
        // happens to share its name must not answer for them.
        //
        // `constructed_control_type_name` above already refuses to treat a
        // user-declared name as a control; this is the same rule for property
        // ACCESS, which did not have it. Without it, `Class Point` with a field
        // `X` had every `p.X` read compiled to a host property getter — the
        // platform `System.Drawing.Point.X` role — instead of a struct field
        // read, so two distinct Points answered with one shared value and it
        // looked exactly like object aliasing. Exact name match, so `MyPoint`
        // and `Pointer` were always fine and only the collision bit.
        //
        // ONLY the direct lookup is skipped. The parent walk below still
        // consults platform roles, which is what keeps `Class MyForm Inherits
        // Form` working — the user name owns nothing of its own, but it
        // inherits `Form`'s roles exactly as before.
        // Case-blind, like the role tables this guards — see
        // `user_owns_type_spelling`.
        if !self.user_owns_type_spelling(type_name) {
            if let Some(role) = declared(type_name) {
                return Some(role);
            }
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
                let class_name = Self::tree_type_key(&hint);
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
                        self.tree_fold(),
                    )
                    .is_none()
                    && vybe_runtime::namespaces::lookup_type_property_setter_target(
                        scopes,
                        &class_name,
                        field,
                        self.tree_fold(),
                    )
                    .is_none()
                    && vybe_runtime::namespaces::lookup_type_instance_member(
                        scopes,
                        &class_name,
                        field,
                        self.tree_fold(),
                    )
                    .is_none()
                    && !self.is_declared_instance_field(&class_name, field)
            }
        }
    }

    /// `x.<prop> = value` where the compiler does not know what `x` IS.
    ///
    /// **This restores a decision the conversion moved.** The retired
    /// property host took the OBJECT at runtime, so a late-bound receiver
    /// reached the widget no matter what the compiler knew.
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
        self.class_get(class_slots::ObjSource::Stack, &class_slots::ClassSlot::internal(CONTROL_TYPE_FIELD));
        let undef_idx = self.import("wasm:js-undefined", "test");
        self.emit_host_call(undef_idx, 1);
        self.chunk().emit_if(line);
        // Not a control — the ordinary object property, exactly as before.
        self.emit_u16(Op::LOCAL_GET, obj_tmp);
        self.emit_u16(Op::LOCAL_GET, value_tmp);
        let prop_key = self.resolve_slot_interned(&class_slots::ClassSlot::internal(&self.canon(prop)));
        self.class_set_resolved(
            class_slots::ObjSource::Stack,
            &prop_key,
            class_slots::ValueSource::Stack,
        );
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
        self.class_get(class_slots::ObjSource::Stack, &class_slots::ClassSlot::internal(CONTROL_TYPE_FIELD));
        let undef_idx = self.import("wasm:js-undefined", "test");
        self.emit_host_call(undef_idx, 1);
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::LOCAL_GET, obj_tmp);
        self.class_get(class_slots::ObjSource::Stack, &class_slots::ClassSlot::internal(&self.canon(prop)));
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
        let mut role = self
            .declared_property_role(type_name, &prop, true)
            .unwrap_or_else(|| prop.clone());
        // A caption written onto a control that is born with children would
        // replace them — see the `unpaintedtext` arm in `property_op`. Asked of
        // the ELEMENT rather than of a control list, so a control acquires the
        // behaviour by declaring chrome and nothing has to be kept in step.
        if matches!(role.as_str(), "text" | "caption")
            && registered_control_element(
                &self.profile.namespaces.type_scopes,
                type_name,
                self.tree_fold(),
            )
            .is_some_and(|element| element.inner_html.is_some())
        {
            role = "unpaintedtext".to_string();
        }
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
            registered_control_element(
                    &self.profile.namespaces.type_scopes,
                    type_name,
                    self.tree_fold(),
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
        let element = registered_control_element(
                    &self.profile.namespaces.type_scopes,
                    type_name,
                    self.tree_fold(),
                )
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
        // **A container is NOT positioned just for being a container.**
        //
        // This used to stamp `position: absolute` on every container element at
        // creation, so that a child's `Left`/`Top` would resolve against it.
        // But `absolute` takes the box OUT OF FLOW, and a frontend whose
        // containers carry no coordinates — Flutter, where `Scaffold`/`Column`/
        // `Row` are pure flex boxes — had its whole tree removed from layout.
        // Every container sat at 0,0 at its default size, one on top of
        // another: an app that rendered as a single label.
        //
        // Positioning is now declared where coordinates actually are, by
        // `emit_gui_property_set`'s `left`/`top` arm. A box that IS placed by
        // coordinates says `position: absolute` about itself — which also makes
        // it a containing block for its own children, so VCL and WinForms keep
        // exactly the nesting they had, and a box that is not placed stays in
        // flow where it belongs.
        for (name, value) in &element.attributes {
            self.emit_declared_attribute(name, value, line);
        }
        for (prop, value) in &element.declares {
            self.emit_declared_style(prop, value, line);
        }
        if let Some(html) = &element.inner_html {
            self.emit_declared_markup(html, line);
        }
    }

    /// One content attribute a control is BORN with
    /// (`ControlElement::attributes`).
    ///
    /// A content attribute, not a property, so it is in the markup a
    /// serialization would show and a selector could match — the same thing
    /// authoring `<select size=4>` gives you.
    fn emit_declared_attribute(&mut self, name: &str, value: &str, line: u32) {
        let element = self.define_local("__gui_declared_attribute");
        self.emit_u16(Op::LOCAL_TEE, element);
        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
        self.chunk().emit_call(doc_idx, 0, line);
        self.emit_u16(Op::LOCAL_GET, element);
        emit_string_const(self.chunk(), name, line);
        emit_string_const(self.chunk(), value, line);
        let set_idx = self.import(DOM_MODULE, "setAttribute");
        self.emit_host_call(set_idx, 4);
        self.chunk().emit_op(Op::DROP, line);
    }

    /// The children a control is BORN with (`ControlElement::inner_html`).
    ///
    /// Set here, at creation, for the same reason the declared CSS is: a
    /// control whose chrome only appeared once something else ran would be
    /// empty for every program that just constructs it — which is precisely
    /// what a designer file does.
    ///
    /// Goes through `setInnerHtml`, so the markup is PARSED, by the same
    /// parser and tree-builder a script's `innerHTML =` uses. The subtree is
    /// therefore ordinary elements: addressable by selector, stylable by the
    /// cascade, and able to carry listeners — not an opaque widget.
    fn emit_declared_markup(&mut self, html: &str, line: u32) {
        let element = self.define_local("__gui_declared_markup");
        self.emit_u16(Op::LOCAL_TEE, element);
        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
        self.chunk().emit_call(doc_idx, 0, line);
        self.emit_u16(Op::LOCAL_GET, element);
        emit_string_const(self.chunk(), html, line);
        let set_idx = self.import(DOM_MODULE, "setInnerHtml");
        self.emit_host_call(set_idx, 3);
        // Same contract as the declared-style emit: the host call's one result
        // is dropped so this leaves exactly the element.
        self.chunk().emit_op(Op::DROP, line);
    }

    /// One piece of CSS a control is BORN with (`ControlElement::declares`).
    ///
    /// Emitted here, at creation, so it is in place before any constructor
    /// argument is applied — a program's own `style` write then simply
    /// overrides it, which is the cascade behaving normally rather than a
    /// precedence rule invented for controls.
    fn emit_declared_style(&mut self, prop: &str, value: &str, line: u32) {
        let element = self.define_local("__gui_declared_style");
        self.emit_u16(Op::LOCAL_TEE, element);
        let doc_idx = self.import(DOCUMENT_MODULE, HOST_FN_ACTIVE_DOCUMENT);
        self.chunk().emit_call(doc_idx, 0, line);
        self.emit_u16(Op::LOCAL_GET, element);
        emit_string_const(self.chunk(), prop, line);
        emit_string_const(self.chunk(), value, line);
        let set_idx = self.import(CSSOM_MODULE, "setStyleProperty");
        self.emit_host_call(set_idx, 4);
        // Same contract as `emit_container_is_positioned`: the host call's
        // result is dropped so this leaves exactly the element.
        self.chunk().emit_op(Op::DROP, line);
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
        self.class_set(
            class_slots::ObjSource::Stack,
            &class_slots::ClassSlot::internal(CONTROL_TYPE_FIELD),
            class_slots::ValueSource::Stack,
        );

        // A control has a NAME whether the program gives it one or not. Every
        // framework here guarantees it — WinForms' designer assigns `Button1`,
        // the VCL assigns `Button1` — and programs read it: the calculator does
        // `displayName := display.Name` and addresses the control by that
        // string afterwards. A bare `createElement` generates no name, so
        // without this `.Name` answered `""` and every by-name write went to a
        // control called nothing.
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

    // `try_emit_gui_property_by_name` is GONE, and with it the last reason for
    // `GUI_MODULE` to be matched against a program's own source.
    //
    // It lowered `vybe.gui.setProperty(name, prop, value)` onto
    // `getElementById` + the property ROLE — correct as far as it went, but it
    // existed to keep an INVENTED spelling working. Real VB.NET has no such
    // call: a control is a variable and you write `display.Text = v`. The two
    // samples that spelled it this way (`examples/vb/calculator.vb`,
    // `form_contacts.vb`) are real VB.NET now, and nothing in the tree names
    // `vybe.gui.*`, so the arm had no input left.

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

// Construction goes through `CtorSpec::control_fn` → `emit_control_element`,
// which creates the ELEMENT (`web:html.activeDocument` + `web:dom.createElement`).
//
// Event subscription is `web:dom`'s `addEventListener`, reached through
// `StmtKind::AddHandler` — which is what every frontend's spelling (`Handles`,
// `+=`, `OnClick :=`) normalizes to. Unsubscription is `removeEventListener`
// via the `on<event>` role, appending a child is `appendChild` on the document,
// and a document is never told to run.

/// Push a string constant onto the stack (helper used when assembling
/// arguments for the DOM calls above).
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
    let key = class_slots::resolve(
        &class_slots::ClassSlot::internal(CONTROL_NAME_FIELD),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &key,
        class_slots::Dest::Stack,
        line,
    );
}

/// Emit a struct_get to read the control's type tag field.
/// Stack on entry: [control_obj]
/// Stack on exit: [type_string]
pub fn emit_get_control_type(chunk: &mut Chunk, line: u32) {
    let key = class_slots::resolve(
        &class_slots::ClassSlot::internal(CONTROL_TYPE_FIELD),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        chunk,
        class_slots::ObjSource::Stack,
        &key,
        class_slots::Dest::Stack,
        line,
    );
}


// ── Linkable chunk builders ──────────────────────────────────────────────────
//
// Linkable chunk builders — the standalone-chunk packaging of what the
// `emit_*` forms splice inline. A language prefix in a name records which
// frontend first needed a linkable chunk, not a language-specific meaning.

// ── rgb(r, g, b) → i32 — pack 24-bit color (VB RGB / GDI 0x00BBGGRR) ─
//
// VB stores RGB color as 0x00BBGGRR (little-endian) — blue in high byte,
// red in low byte. Pack: `(b << 16) | (g << 8) | r`. Pure i32 ops.
pub fn build_rgb(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_rgb");
    c.arity = 3;
    c.local_count = 3;

    // (b & 0xFF) << 16
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::I32_FROM_F64, 0);
    let mask = c.add_constant(Value::I32(0xFF));
    crate::primitives::expressions::emit_const_index(&mut c, mask, 0);
    c.emit_op(Op::I32_AND, 0);
    let sh16 = c.add_constant(Value::I32(16));
    crate::primitives::expressions::emit_const_index(&mut c, sh16, 0);
    c.emit_op(Op::I32_SHL, 0);

    // (g & 0xFF) << 8
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::I32_FROM_F64, 0);
    crate::primitives::expressions::emit_const_index(&mut c, mask, 0);
    c.emit_op(Op::I32_AND, 0);
    let sh8 = c.add_constant(Value::I32(8));
    crate::primitives::expressions::emit_const_index(&mut c, sh8, 0);
    c.emit_op(Op::I32_SHL, 0);
    c.emit_op(Op::I32_OR, 0);

    // (r & 0xFF)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::I32_FROM_F64, 0);
    crate::primitives::expressions::emit_const_index(&mut c, mask, 0);
    c.emit_op(Op::I32_AND, 0);
    c.emit_op(Op::I32_OR, 0);

    c.emit_op(Op::RETURN, 0);
    c
}

// ── qbcolor(c) → i32 — QBasic 16-color palette → packed RGB ──
//
// QBasic's COLOR statement uses the EGA/VGA 16-color palette. Map
// 0-15 to standard palette entries; out-of-range returns black.
pub fn build_qbcolor(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_qbcolor");
    c.arity = 1;
    c.local_count = 1;

    // QBasic palette in 0x00BBGGRR (VB RGB) layout:
    // 0=black, 1=blue, 2=green, 3=cyan, 4=red, 5=magenta, 6=brown,
    // 7=lightgray, 8=darkgray, 9=lightblue, 10=lightgreen,
    // 11=lightcyan, 12=lightred, 13=lightmagenta, 14=yellow, 15=white.
    let palette: [i32; 16] = [
        0x000000, 0x800000, 0x008000, 0x808000, 0x000080, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
        0xFF0000, 0x00FF00, 0xFFFF00, 0x0000FF, 0xFF00FF, 0x00FFFF, 0xFFFFFF,
    ];

    // Build the palette as a constant array, then ARRAY_GET by index.
    // Compile-time pack: emit ARRAY_NEW + 16 push-style emits → array.
    // Simpler: chain SELECTs for the 16 entries — but that's 15 selects
    // and bloats the chunk. Use a small array literal instead.
    let arr_locals_start = 1u16;
    c.local_count = 2;
    crate::primitives::collections::emit_array_new_into(_imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, arr_locals_start, 0);
    for &val in palette.iter() {
        let v_const = c.add_constant(Value::I32(val));
        c.emit_op_u16(Op::LOCAL_GET, arr_locals_start, 0);
        crate::primitives::expressions::emit_const_index(&mut c, v_const, 0);
        crate::primitives::collections::emit_push_into(_imports, &mut c, 0);
        c.emit_op(Op::DROP, 0);
    }
    // ARRAY_GET(arr, idx & 0xF) — clamp via mask so out-of-range wraps.
    c.emit_op_u16(Op::LOCAL_GET, arr_locals_start, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::I32_FROM_F64, 0);
    let mask = c.add_constant(Value::I32(0xF));
    crate::primitives::expressions::emit_const_index(&mut c, mask, 0);
    c.emit_op(Op::I32_AND, 0);
    crate::primitives::collections::emit_get_into(_imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
