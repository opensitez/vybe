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
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

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
pub fn host_fn_new_control(canonical: &str) -> String {
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

// ─── Component Model Registration ────────────────────────────────────────────

/// Register all `vybe:gui` host functions as component module exports.
/// This is called by the compiler/Linker to populate the component descriptor
/// so all languages automatically get GUI functions without per-language
/// profile duplication (similar to WASI registration).
pub fn gui_component_exports() -> Vec<vybe_bytecode::component_model::ComponentExport> {
    use vybe_bytecode::component::{FuncSig, ValType};
    use vybe_bytecode::component_model::{ComponentExport, ComponentItemKind};

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
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
}

/// Emit a struct_get to read the control's type tag field.
/// Stack on entry: [control_obj]
/// Stack on exit: [type_string]
pub fn emit_get_control_type(chunk: &mut Chunk, line: u32) {
    let key = chunk.add_constant(Value::String(Arc::from(CONTROL_TYPE_FIELD)));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
}
