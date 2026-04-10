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
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

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
        "button"          => "Button",
        "checkbox"        => "CheckBox",
        "radiobutton"     => "RadioButton",
        "togglebutton"    => "ToggleButton",
        "linkbutton" | "linklabel" => "LinkLabel",

        // ── Text ──
        "label"           => "Label",
        "textbox" | "entry" | "textfield" => "TextBox",
        "richtextbox"     => "RichTextBox",
        "maskedtextbox"   => "MaskedTextBox",

        // ── Selection ──
        "combobox" | "dropdown" => "ComboBox",
        "listbox"         => "ListBox",
        "listview"        => "ListView",
        "treeview"        => "TreeView",

        // ── Containers ──
        "panel" | "container" => "Panel",
        "groupbox"        => "GroupBox",
        "tabcontrol" | "tabbedpane" => "TabControl",
        "tabpage"         => "TabPage",
        "splitcontainer"  => "SplitContainer",
        "flowlayoutpanel" => "FlowLayoutPanel",
        "tablelayoutpanel" => "TableLayoutPanel",

        // ── Date / time / numeric ──
        "datetimepicker"  => "DateTimePicker",
        "monthcalendar"   => "MonthCalendar",
        "numericupdown"   => "NumericUpDown",

        // ── Progress / indicators ──
        "progressbar"     => "ProgressBar",
        "trackbar" | "slider" => "TrackBar",

        // ── Images / media ──
        "picturebox" | "image" => "PictureBox",
        "webbrowser"      => "WebBrowser",

        // ── Data / grids ──
        "datagridview" | "datagrid" => "DataGridView",

        // ── Menus / strips ──
        "menustrip" | "menubar" => "MenuStrip",
        "toolstrip" | "toolbar" => "ToolStrip",
        "statusstrip" | "statusbar" => "StatusStrip",
        "contextmenustrip" | "contextmenu" => "ContextMenuStrip",

        // ── Scrollbars ──
        "hscrollbar"      => "HScrollBar",
        "vscrollbar"      => "VScrollBar",

        // ── Dialogs ──
        "openfiledialog"  => "OpenFileDialog",
        "savefiledialog"  => "SaveFileDialog",
        "folderbrowserdialog" => "FolderBrowserDialog",
        "colordialog"     => "ColorDialog",
        "fontdialog"      => "FontDialog",

        // ── Non-visual / lifecycle ──
        "timer"           => "Timer",
        "tooltip"         => "ToolTip",
        "imagelist"       => "ImageList",
        "notifyicon"      => "NotifyIcon",
        "errorprovider"   => "ErrorProvider",
        "helpprovider"    => "HelpProvider",
        "backgroundworker" => "BackgroundWorker",
        "bindingsource"   => "BindingSource",
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
    chunk.emit_op_u16(Op::call_import, import_idx, line);
    chunk.emit(argc, line);
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
    chunk.emit_op_u16(Op::call_import, import_idx, line);
    chunk.emit(3, line);
}

/// Emit `vybe:gui::removeEvent(control_name, event_name, handler_fn)`.
/// `import_idx` must be the result of `compiler.import("vybe:gui", HOST_FN_UNBIND_EVENT)`.
/// Caller drops the result if used as a statement.
pub fn emit_unbind_event(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_op_u16(Op::call_import, import_idx, line);
    chunk.emit(3, line);
}

/// Emit `vybe:gui::controlsAdd(parent, child)`.
/// `import_idx` must be `compiler.import("vybe:gui", HOST_FN_ADD_CHILD)`.
/// Stack on entry: [parent, child]
/// Stack on exit:  [host_call_result]   (caller drops if statement)
pub fn emit_add_child(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_op_u16(Op::call_import, import_idx, line);
    chunk.emit(2, line);
}

/// Emit `vybe:gui::runApplication(form)`.
/// `import_idx` must be `compiler.import("vybe:gui", HOST_FN_RUN_APPLICATION)`.
pub fn emit_run_application(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_op_u16(Op::call_import, import_idx, line);
    chunk.emit(1, line);
}

/// Emit `vybe:gui::appExit()`.
/// `import_idx` must be `compiler.import("vybe:gui", HOST_FN_APP_EXIT)`.
pub fn emit_app_exit(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_op_u16(Op::call_import, import_idx, line);
    chunk.emit(0, line);
}

/// Emit `vybe:gui::raiseEvent(arg0, ..., argN, event_name)`.
/// `import_idx` must be `compiler.import("vybe:gui", HOST_FN_RAISE_EVENT)`.
/// Stack on entry: [arg0, arg1, ..., event_name_string]
/// Stack on exit:  [host_call_result]   (caller drops if statement)
pub fn emit_raise_event(chunk: &mut Chunk, import_idx: u16, total_args: u8, line: u32) {
    chunk.emit_op_u16(Op::call_import, import_idx, line);
    chunk.emit(total_args, line);
}

/// Emit a control-property assignment that mirrors to the GUI host registry.
///
/// This is the canonical pattern for `Me.Text = "X"` (and any
/// `<control>.<Property> = value` inside a class method): writes to BOTH
/// the in-memory struct AND the GUI host's property table so that
/// `gui.get_property(control_name, "Text")` reflects the new value.
///
/// Caller arranges:
///   - The OBJECT in `obj_slot` (a temp local already populated)
///   - The VALUE in `val_slot` (a temp local already populated)
///
/// Caller passes:
///   - `field_lower`  — canonical (lowercased) field key for the in-memory
///     struct_set, e.g. `"text"`
///   - `field_pascal` — PascalCase form for the host fn (matches widget
///     property names), e.g. `"Text"`
///   - `set_property_import_idx` — pre-resolved index for
///     `vybe:gui::controlSetProperty` from the compiler's import table
///
/// Stack on entry: []
/// Stack on exit:  [] (both side effects performed)
///
/// All language frontends (.NET WinForms, future MAUI, Flutter, Tkinter, …)
/// share this single emit path so a property write looks identical to the
/// host regardless of the source language.
pub fn emit_set_control_property(
    chunk: &mut Chunk,
    obj_slot: u16,
    val_slot: u16,
    field_lower: &str,
    field_pascal: &str,
    set_property_import_idx: u16,
    line: u32,
) {
    // 1. In-memory: obj.<field> = value
    let field_key = chunk.add_constant(Value::String(Arc::from(field_lower)));
    chunk.emit_op_u16(Op::local_get, obj_slot, line);
    chunk.emit_op_u16(Op::local_get, val_slot, line);
    chunk.emit_op_u16(Op::struct_set, field_key, line);
    chunk.emit_op(Op::drop, line);

    // 2. Host mirror: vybe:gui::controlSetProperty(obj, "Field", value)
    //    The host fn keys the GUI registry by obj.__control_name, so user
    //    classes whose `emit_new_typed_object` stamped __control_name = the
    //    lowercased class name are reachable via `gui.get_property(name, ...)`.
    chunk.emit_op_u16(Op::local_get, obj_slot, line);
    let prop_str = chunk.add_constant(Value::String(Arc::from(field_pascal)));
    chunk.emit_op_u16(Op::r#const, prop_str, line);
    chunk.emit_op_u16(Op::local_get, val_slot, line);
    chunk.emit_op_u16(Op::call_import, set_property_import_idx, line);
    chunk.emit(3, line);
    chunk.emit_op(Op::drop, line);
}

/// Push a string constant onto the stack (helper used when assembling
/// arguments for the GUI host calls above).
pub fn emit_string_const(chunk: &mut Chunk, s: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(s)));
    chunk.emit_op_u16(Op::r#const, idx, line);
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
    chunk.emit_op_u16(Op::struct_get, key, line);
}

/// Emit a struct_get to read the control's type tag field.
/// Stack on entry: [control_obj]
/// Stack on exit: [type_string]
pub fn emit_get_control_type(chunk: &mut Chunk, line: u32) {
    let key = chunk.add_constant(Value::String(Arc::from(CONTROL_TYPE_FIELD)));
    chunk.emit_op_u16(Op::struct_get, key, line);
}
