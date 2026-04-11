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
use super::{DotnetClass, DotnetMethod, MethodTarget};

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

// ─── Method thunk chunk ─────────────────────────────────────────────────────

/// Build the thunk chunk that bridges a `.NET` instance method call to
/// either a host import or a dotnet class constructor, depending on
/// `method.target`.
///
/// **`MethodTarget::Host`** — the chunk forwards `(this, arg0, ..., argN-1)`
/// to the host import:
///
/// ```text
/// fn <ClassName>::<MethodName>(this, arg0, ..., argN-1):
///     call_import <host_module>::<host_fn> with [this, arg0, ..., argN-1]
///     return result
/// ```
///
/// Host implementations that don't care about `this` (conceptually static
/// methods) just ignore `args[0]`. Implementations that DO care (like
/// `Graphics::DrawLine`) read `args[0]` to find the target and `args[1..]`
/// for the user's arguments.
///
/// **`MethodTarget::DotnetCtor`** — the chunk discards `this` and calls
/// the target dotnet class's ctor with the user args:
///
/// ```text
/// fn <ClassName>::<MethodName>(this, arg0, ..., argN-1):
///     global_get <target_class>
///     local_get arg0 ; local_get arg1 ; ... ; local_get argN-1
///     call (N-1)
///     return
/// ```
///
/// Used by methods like `Control.CreateGraphics()` which return a fresh
/// `Graphics` instance — going through the dotnet ctor (rather than the
/// raw `vybe:drawing::graphicsNew` host fn) ensures the returned object
/// has all `Graphics` methods bound on it via the standard registration
/// flow.
///
/// `arity` is the total arity including `this` (so `Show()` has
/// `arity = 1`; `DrawLine(p, x1, y1, x2, y2)` is `arity = 6`).
///
/// `import_idx` is only consulted for the `Host` variant — the
/// orchestrator pre-resolves the import via `chunks[0].add_import(...)`
/// and passes the resulting index in. For `DotnetCtor` the chunk uses
/// `global_get` instead and `import_idx` is ignored.
///
/// ## Local layout
///
/// - slot 0 = closure ref (reserved by the VM call frame)
/// - slot 1 = `this`
/// - slot 2..=arity = user args
pub fn build_method_thunk_chunk(
    class_name: &str,
    method: &DotnetMethod,
    import_idx: u16,
) -> Chunk {
    let chunk_name = format!("{}::{}", class_name, method.name);
    let mut chunk = create_function_chunk(&chunk_name, method.arity);
    let line = 0u32;

    match method.target {
        MethodTarget::Host { .. } => {
            // Push this + each user arg in order, then call_import.
            for slot in 1..=method.arity as u16 {
                chunk.emit_op_u16(Op::local_get, slot, line);
            }
            chunk.emit_op_u16(Op::call_import, import_idx, line);
            chunk.emit(method.arity, line);
            // Result of the host call is on the stack — return it. For
            // void methods (most setters / `Show` / `DrawLine`) the host
            // fn returns `Value::Null`, which is fine.
            chunk.emit_op(Op::r#return, line);
        }
        MethodTarget::DotnetCtor { class: target_class } => {
            // Discard `this` (slot 1) — factory-style methods don't pass
            // it to the target ctor. Push the target class global, then
            // the user args (slots 2..=arity), then call.
            let target_const = chunk.add_constant(Value::String(Arc::from(target_class)));
            chunk.emit_op_u16(Op::global_get, target_const, line);
            // User args only — skip slot 1 (this).
            for slot in 2..=method.arity as u16 {
                chunk.emit_op_u16(Op::local_get, slot, line);
            }
            // arity - 1 because we dropped `this`.
            chunk.emit_op_u8(Op::call, method.arity - 1, line);
            chunk.emit_op(Op::r#return, line);
        }
    }

    chunk.local_count = method.arity as u16;
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

/// Per-method thunk binding info supplied to [`build_constructor_chunk`].
///
/// `method_name` is the lowercased instance-side key (`"createGraphics"`
/// → `"creategraphics"`) — written by the user as `obj.MethodName(...)`,
/// looked up by the VM via the lowercased canonical AST.
#[derive(Debug, Clone, Copy)]
pub struct MethodBinding<'a> {
    pub method_name: &'a str,
    pub thunk_chunk_idx: usize,
}

/// Build the constructor chunk for one .NET class.
///
/// The chunk implements:
///
/// ```text
/// fn <ClassName>(arg0, arg1, ..., argN-1):        # arity = class.ctor_arity
///     this = <Parent>()                            # arity-0 parent call
///     this.__type = "<ClassName>"                  # re-stamp
///     this.__set_<prop1> = ref_func(setter1)       # for each property
///     ...
///     this.<method1> = ref_func(thunk1)            # for each method
///     ...
///     # If concrete leaf:
///     widget = <host_module>::<host_fn>(arg0, ..., argN-1)
///     this.__control_name = widget.__control_name
///     this.__control_type = widget.__control_type
///     return this
/// ```
///
/// For the root class (`Object`, `parent = None`) the body starts with
/// `struct_new 0; local_set this; drop` instead of a parent ctor call.
///
/// The parent ctor is ALWAYS called with 0 args. When `class.ctor_arity > 0`
/// the user-supplied args are forwarded to the leaf widget host fn — they
/// are NOT propagated up the inheritance chain (abstract bases like
/// `Object` / `Component` / `Brush` take no args at this layer).
///
/// ## Local layout
///
/// VM call frames reserve slot 0 for the closure ref; locals start at
/// slot 1. For a `ctor_arity = N` ctor:
/// - slot 0 = closure ref (reserved by the VM)
/// - slots 1..=N = user-supplied ctor args
/// - slot N+1 = `this`
/// - slot N+2 = `widget` (only when wiring a concrete widget host fn)
///
/// `widget_new_import_idx` (when class is concrete) is the chunk[0] import
/// index for the configured `widget_host_module::widget_host_fn`.
pub fn build_constructor_chunk(
    class: &DotnetClass,
    setter_bindings: &[SetterBinding],
    method_bindings: &[MethodBinding],
    widget_new_import_idx: Option<u16>,
) -> Chunk {
    let mut chunk = create_function_chunk(class.name, class.ctor_arity);
    let line = 0u32;
    let arity = class.ctor_arity as u16;
    let this_slot: u16 = arity + 1;
    let widget_slot: u16 = arity + 2;

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

    // ── Step 4: bind methods for THIS class ────────────────────────────────
    //
    // Each method thunk was pre-built and pushed by the orchestrator. The
    // method is bound under its lowercased name to match the canonical AST
    // shape produced by every walker.
    //
    // Inheritance order matters: parents are registered first, so a child
    // re-binding the same method name overwrites the parent's binding via
    // the same `struct_set` — exactly how virtual override works in .NET.
    for binding in method_bindings {
        chunk.emit_op_u16(Op::local_get, this_slot, line);
        chunk.emit_op_u16(Op::ref_func, binding.thunk_chunk_idx as u16, line);
        chunk.emit(0, line); // 0 upvalues
        let key = chunk.add_constant(Value::String(Arc::from(binding.method_name)));
        chunk.emit_op_u16(Op::struct_set, key, line);
        chunk.emit_op(Op::drop, line);
    }

    // ── Step 5: concrete leaf — wire backing object ────────────────────────
    //
    // Calls `<widget_host_module>::<widget_host_fn>(args...)`, then copies
    // the backing object's identity fields onto `this`. The inherited
    // setters and methods stay intact because we never overwrite their
    // keys here.
    //
    // Args: for `class.ctor_arity > 0` we forward the user-supplied ctor
    // args (slots 1..=arity) to the host fn. For arity-0 classes the host
    // fn is called with no args.
    if let Some(import_idx) = widget_new_import_idx {
        for i in 1..=arity {
            chunk.emit_op_u16(Op::local_get, i, line);
        }
        chunk.emit_op_u16(Op::call_import, import_idx, line);
        chunk.emit(arity as u8, line);
        chunk.emit_op_u16(Op::local_set, widget_slot, line);
        chunk.emit_op(Op::drop, line);

        // Copy backing identity fields: this.<f> = widget.<f>.
        //
        // NB: `name` is intentionally NOT copied. The backing host fn
        // pre-stamps `name` with an auto-generated id (e.g. "Form_3"), and
        // `Control` has a `__set_name` setter bound — so a
        // `struct_set "name"` would dispatch to that setter which calls
        // `controlSetProperty(this, "Name", widget.name)`. That writes the
        // auto-id into the gui state registry under whatever
        // `__control_name` `this` had at that moment, polluting the
        // registry. The canonical control name is stamped later by user
        // code or by the walker normalization (`Me.__control_name = "..."`).
        //
        // We also no longer copy `show`/`close`/`focus`/`hide` here. Those
        // used to be host-stamped on the backing object, but they're now
        // installed via real method thunks bound at this class's level
        // (or inherited from `Control`).
        for field in &["__control_name", "__control_type"] {
            let key_idx = chunk.add_constant(Value::String(Arc::from(*field)));
            chunk.emit_op_u16(Op::local_get, this_slot, line);
            chunk.emit_op_u16(Op::local_get, widget_slot, line);
            chunk.emit_op_u16(Op::struct_get, key_idx, line);
            chunk.emit_op_u16(Op::struct_set, key_idx, line);
            chunk.emit_op(Op::drop, line);
        }
    }

    // ── Step 6: return this ─────────────────────────────────────────────────
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    chunk.emit_op(Op::r#return, line);

    // Locals beyond the closure ref slot 0:
    //   slots 1..=ctor_arity = ctor args
    //   slot ctor_arity+1     = this   (always present)
    //   slot ctor_arity+2     = widget (only when wiring a backing host fn)
    chunk.local_count = if widget_new_import_idx.is_some() {
        arity + 2
    } else {
        arity + 1
    };
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
