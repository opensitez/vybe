//! Chunk-building helpers for `.NET` class wrappers.
//!
//! Every helper here returns a fully-formed `Chunk` that the orchestrator
//! (in the user's compiler) appends to its `chunks` vec. The helpers don't
//! touch the compiler's `defined_globals` / `pending_classes` bookkeeping —
//! that's the orchestrator's job.
//!
//! ## Stack discipline
//!
//! - Setter chunks have arity 2: `[this=slot 0, value=slot 1]`. They call
//!   `vybe:gui::controlSetProperty(this, "PropName", value)`, drop the
//!   host return value, and return null.
//!
//! - Constructor chunks have arity 0 (the .NET BCL classes have no
//!   user-visible constructor params at this layer — user code that wants
//!   to pass args sits in a child class that overrides `New()` and calls
//!   the parameterless base ctor). Slot 0 is `this`. The child-class flow
//!   in `compile_class` already supports this shape by emitting
//!   `global_get parent; call(0); local_set this_slot`.
//!
//! ## Import indices
//!
//! `vybe:gui::controlSetProperty` and the `vybe:gui::new_<Type>` host fns
//! must be added to `chunks[0].imports` by the orchestrator. The resulting
//! `u16` import index is passed into every helper that needs it. This
//! matches how `compiler_common::gui::emit_*` already works and how the VM
//! resolves `call_import` indices through the script chunk's import table.

use std::sync::Arc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

use crate::functions::create_function_chunk;
use super::DotnetClass;

// ─── Setter chunk ───────────────────────────────────────────────────────────

/// Build the setter chunk for one property.
///
/// The chunk implements:
///
/// ```text
/// fn __set_<prop>(this, value):
///     vybe:gui::controlSetProperty(this, "<PropName>", value)
///     return null
/// ```
///
/// `prop_pascal` is the .NET PascalCase property name passed as the second
/// argument to `controlSetProperty` so the host's gui state registry sees
/// the canonical key (`"Text"`, `"FormBorderStyle"`). The setter is bound
/// on the object under `__set_<lowercased>` by [`super::DotnetClass`]
/// orchestration so the VM's `struct_set → __set_<field>` dispatch finds
/// it for `Me.Text = "X"` (which lowercases to `text`).
///
/// `set_property_import_idx` must already point to the
/// `vybe:gui::controlSetProperty` import in `chunks[0]`.
///
/// ## Local layout
///
/// VM call frames reserve slot 0 for the closure ref of the function being
/// called; user-visible locals (params and additional locals) start at
/// slot 1. So for a setter with `arity = 2 (this, value)`:
/// - slot 0 = closure ref (reserved by the VM)
/// - slot 1 = `this`
/// - slot 2 = `value`
pub fn build_setter_chunk(
    class_name: &str,
    prop_pascal: &str,
    set_property_import_idx: u16,
) -> Chunk {
    let chunk_name = format!("{}::__set_{}", class_name, prop_pascal.to_lowercase());
    let mut chunk = create_function_chunk(&chunk_name, 2); // (this, value)
    let line = 0u32;

    // [this]
    chunk.emit_op_u16(Op::local_get, 1, line);
    // [this, "PropName"]
    let prop_const = chunk.add_constant(Value::String(Arc::from(prop_pascal)));
    chunk.emit_op_u16(Op::r#const, prop_const, line);
    // [this, "PropName", value]
    chunk.emit_op_u16(Op::local_get, 2, line);
    // [this, "PropName", value] → call_import controlSetProperty(3) → [result]
    chunk.emit_op_u16(Op::call_import, set_property_import_idx, line);
    chunk.emit(3, line);
    // drop the host return value
    chunk.emit_op(Op::drop, line);
    // return null
    chunk.emit_op(Op::null, line);
    chunk.emit_op(Op::r#return, line);

    chunk.local_count = 2; // this + value
    chunk
}

// ─── Constructor chunk ──────────────────────────────────────────────────────

/// Per-property setter binding info supplied to [`build_constructor_chunk`].
///
/// `prop_pascal` is the .NET property name; `setter_chunk_idx` is the chunk
/// index returned when the orchestrator pushed the setter chunk into the
/// compiler's `chunks` vec.
#[derive(Debug, Clone, Copy)]
pub struct SetterBinding<'a> {
    pub prop_pascal: &'a str,
    pub setter_chunk_idx: usize,
}

/// Build the constructor chunk for one .NET class.
///
/// The chunk implements:
///
/// ```text
/// fn <ClassName>():                              # arity 0
///     this = <Parent>()                          # if parent is Some
///     this.__type = "<ClassName>"                # re-stamp
///     this.__set_<prop1> = ref_func(setter1)     # for each property
///     this.__set_<prop2> = ref_func(setter2)
///     ...
///     # If concrete leaf:
///     widget = vybe:gui::new_<ClassName>()
///     this.__control_name = widget.__control_name
///     this.__control_type = widget.__control_type
///     this.name           = widget.name
///     this.show           = widget.show
///     this.close          = widget.close
///     this.focus          = widget.focus
///     this.hide           = widget.hide
///     return this
/// ```
///
/// For the root class (`Object`, `parent = None`) the body starts with
/// `struct_new 0; local_set this; drop` instead of a parent ctor call.
///
/// ## Local layout
///
/// VM call frames reserve slot 0 for the closure ref; locals start at
/// slot 1. So for an arity-0 ctor:
/// - slot 0 = closure ref (reserved by the VM)
/// - slot 1 = `this`
/// - slot 2 = `widget` (only when wiring a concrete widget host fn)
///
/// `widget_new_import_idx` (when class is concrete) is the chunk[0] import
/// index for `vybe:gui::new_<ClassName>`.
pub fn build_constructor_chunk(
    class: &DotnetClass,
    setter_bindings: &[SetterBinding],
    widget_new_import_idx: Option<u16>,
) -> Chunk {
    let mut chunk = create_function_chunk(class.name, 0);
    let line = 0u32;
    let this_slot: u16 = 1;
    let widget_slot: u16 = 2;

    // ── Step 1: get `this` ──────────────────────────────────────────────────
    if let Some(parent_name) = class.parent {
        // this = <Parent>()  — global_get parent ; call(0) ; local_set this ; drop
        let parent_const = chunk.add_constant(Value::String(Arc::from(parent_name)));
        chunk.emit_op_u16(Op::global_get, parent_const, line);
        chunk.emit_op_u8(Op::call, 0, line);
        chunk.emit_op_u16(Op::local_set, this_slot, line);
        chunk.emit_op(Op::drop, line);
    } else {
        // Root class (Object): this = struct_new 0
        chunk.emit_op_u16(Op::struct_new, 0, line);
        chunk.emit_op_u16(Op::local_set, this_slot, line);
        chunk.emit_op(Op::drop, line);
    }

    // ── Step 2: re-stamp __type with this class's name ──────────────────────
    {
        chunk.emit_op_u16(Op::local_get, this_slot, line);
        let type_str = chunk.add_constant(Value::String(Arc::from(class.name)));
        chunk.emit_op_u16(Op::r#const, type_str, line);
        let type_key = chunk.add_constant(Value::String(Arc::from("__type")));
        chunk.emit_op_u16(Op::struct_set, type_key, line);
        chunk.emit_op(Op::drop, line);
    }

    // ── Step 3: bind setters for THIS class's properties ────────────────────
    //
    // Each setter chunk was pre-built and pushed by the orchestrator; the
    // orchestrator passes the resulting chunk indices in `setter_bindings`.
    for binding in setter_bindings {
        let set_name = format!("__set_{}", binding.prop_pascal.to_lowercase());
        // local_get this ; ref_func setter ; struct_set "__set_<prop>" ; drop
        chunk.emit_op_u16(Op::local_get, this_slot, line);
        chunk.emit_op_u16(Op::ref_func, binding.setter_chunk_idx as u16, line);
        chunk.emit(0, line); // 0 upvalues
        let key = chunk.add_constant(Value::String(Arc::from(set_name.as_str())));
        chunk.emit_op_u16(Op::struct_set, key, line);
        chunk.emit_op(Op::drop, line);
    }

    // ── Step 4: concrete leaf — wire vybe_widgets backing ──────────────────
    //
    // Calls `vybe:gui::new_<ClassName>()`, then copies the widget identity
    // fields onto `this`. The inherited setters stay intact because we
    // never overwrite them.
    if let Some(import_idx) = widget_new_import_idx {
        // widget = vybe:gui::new_<ClassName>()
        chunk.emit_op_u16(Op::call_import, import_idx, line);
        chunk.emit(0, line); // 0 args
        chunk.emit_op_u16(Op::local_set, widget_slot, line);
        chunk.emit_op(Op::drop, line);

        // Copy each identity field: this.<f> = widget.<f>
        //
        // NB: `name` is intentionally NOT copied here. The vybe_widgets
        // host fn pre-stamps `name` with an auto-generated id (e.g.
        // "Form_3"), and `Control` has `__set_name` bound — so a
        // `struct_set "name"` would dispatch to `__set_name` which calls
        // `controlSetProperty(this, "Name", widget.name)`. That writes the
        // widget's auto-id into the gui state registry under the user's
        // not-yet-set canonical control name (which is empty at this
        // point), polluting the registry with a `("", "Name")` entry. The
        // canonical control name is stamped later by user code or by the
        // VB walker normalization (`Me.__control_name = "<lower class
        // name>"`), and any subsequent `Me.Name = "..."` does the right
        // thing through the same setter dispatch.
        for field in &[
            "__control_name",
            "__control_type",
            "show",
            "close",
            "focus",
            "hide",
        ] {
            let key_idx = chunk.add_constant(Value::String(Arc::from(*field)));
            // [this]
            chunk.emit_op_u16(Op::local_get, this_slot, line);
            // [this, widget]
            chunk.emit_op_u16(Op::local_get, widget_slot, line);
            // [this, widget_field_value]
            chunk.emit_op_u16(Op::struct_get, key_idx, line);
            // struct_set this.<field> = widget_field_value → [field_value]
            chunk.emit_op_u16(Op::struct_set, key_idx, line);
            chunk.emit_op(Op::drop, line);
        }
    }

    // ── Step 5: return this ─────────────────────────────────────────────────
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    chunk.emit_op(Op::r#return, line);

    // Locals beyond the closure ref slot 0:
    //   slot 1 = this  (always present)
    //   slot 2 = widget (only when wiring a concrete widget host fn)
    chunk.local_count = if widget_new_import_idx.is_some() { 2 } else { 1 };
    chunk
}

// ─── Class-global installer ─────────────────────────────────────────────────

/// Emit code into the script chunk that installs a .NET class as a callable
/// global named `class_name`. After this runs, user code can write
/// `Inherits ClassName` and the existing `compile_class` machinery will
/// resolve the parent via `global_get class_name; call(0)`.
///
/// Stack: unchanged. Emits:
///
/// ```text
/// ref_func ctor_chunk_idx ; global_set "<ClassName>" ; drop
/// # case-insensitive alias for VB / Pascal
/// ref_func ctor_chunk_idx ; global_set "<classname>" ; drop
/// ```
///
/// Two aliases are emitted because VB resolves identifiers case-
/// insensitively (`Inherits Form` lowercases to `form` in the canonical
/// AST), while C# / Dart / Python keep PascalCase.
pub fn emit_install_class_global(
    script_chunk: &mut Chunk,
    class_name: &str,
    ctor_chunk_idx: usize,
    line: u32,
) {
    // Original case
    script_chunk.emit_op_u16(Op::ref_func, ctor_chunk_idx as u16, line);
    script_chunk.emit(0, line); // 0 upvalues
    let name_const = script_chunk.add_constant(Value::String(Arc::from(class_name)));
    script_chunk.emit_op_u16(Op::global_set, name_const, line);
    script_chunk.emit_op(Op::drop, line);

    // Lowercase alias (skip if already lowercase)
    let lower = class_name.to_lowercase();
    if lower != class_name {
        script_chunk.emit_op_u16(Op::ref_func, ctor_chunk_idx as u16, line);
        script_chunk.emit(0, line);
        let lower_const = script_chunk.add_constant(Value::String(Arc::from(lower.as_str())));
        script_chunk.emit_op_u16(Op::global_set, lower_const, line);
        script_chunk.emit_op(Op::drop, line);
    }
}
