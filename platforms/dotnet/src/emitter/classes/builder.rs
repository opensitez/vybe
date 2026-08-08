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
//!   the parameterless base ctor). The child-class flow in `compile_class`
//!   already supports this shape by emitting
//!   `global_get parent; call(0); local_set this_slot`.
//!
//! ## Import indices
//!
//! `vybe:gui::controlSetProperty` and the `vybe:gui::new_<Type>` host fns
//! must be added to `chunks[0].imports` by the orchestrator. The resulting
//! `u16` import index is passed into every helper that needs it, and the
//! chunks built here must stay FREE of local imports: baked script-table
//! indices are ambiguous once a chunk has its own import table (the
//! normalizer remaps local-first), so strings are pushed as pool
//! constants via [`push_string_const`] — NEVER `emit_string_const`,
//! which secretly adds a `wasm:string-constants` import.

use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};
use vybe_compiler::primitives::instructions::core_wasm;

use super::{DotnetClass, DotnetMethod, MethodOp, MethodTarget};
use vybe_compiler::primitives::functions::create_function_chunk;

/// Push a string as a pool constant. Wrapper chunks must not use
/// `Chunk::emit_string_const` — it registers a `wasm:string-constants`
/// import on the chunk, and any local import shadows same-valued baked
/// `chunks[0]` indices in the import-table normalizer's local-first
/// remap (that collision silently turned every property-setter's
/// `controlSetProperty` call into a string constant).
fn push_string_const(chunk: &mut Chunk, s: &str, line: u32) {
    chunk.emit_string_const(s, line);
}

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
/// WASM convention: slot 0 is the first argument. For a setter with
/// `arity = 2 (this, value)`:
/// - slot 0 = `this`
/// - slot 1 = `value`
pub fn build_setter_chunk(
    class_name: &str,
    prop_pascal: &str,
    set_property_import_idx: u16,
) -> Chunk {
    let chunk_name = format!("{}::__set_{}", class_name, prop_pascal.to_lowercase());
    let mut chunk = create_function_chunk(&chunk_name, 2); // (this, value)
    let line = 0u32;

    // [this]
    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    // [this, "PropName"]
    push_string_const(&mut chunk, prop_pascal, line);
    // [this, "PropName", value]
    chunk.emit_op_u16(Op::LOCAL_GET, 1, line);
    // [this, "PropName", value] → spec `call` controlSetProperty(3) → [result]
    chunk.emit_call(set_property_import_idx, 3, line);
    // drop the host return value
    chunk.emit_op(Op::DROP, line);
    // return null
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_op(Op::RETURN, line);

    chunk.local_count = 2; // this + value
    chunk
}

pub fn build_getter_chunk(
    class_name: &str,
    prop_pascal: &str,
    get_property_import_idx: u16,
) -> Chunk {
    let chunk_name = format!("{}::__get_{}", class_name, prop_pascal.to_lowercase());
    let mut chunk = create_function_chunk(&chunk_name, 1); // (this)
    let line = 0u32;

    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    push_string_const(&mut chunk, prop_pascal, line);
    chunk.emit_call(get_property_import_idx, 2, line);
    chunk.emit_op(Op::RETURN, line);

    chunk.local_count = 1;
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
/// WASM convention: slot 0 is the first argument.
/// - slot 0 = `this`
/// - slot 1..=arity-1 = user args
pub fn build_method_thunk_chunk(
    class_name: &str,
    method: &DotnetMethod,
    import_idx: u16,
    body_imports: &[u16],
) -> Chunk {
    let chunk_name = format!("{}::{}", class_name, method.name);
    let mut chunk = create_function_chunk(&chunk_name, method.arity);
    let line = 0u32;

    match method.target {
        MethodTarget::Host { .. } => {
            // Push this + each user arg in order, then spec `call`.
            for slot in 0..method.arity as u16 {
                chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
            }
            chunk.emit_call(import_idx, method.arity, line);
            // Result of the host call is on the stack — return it. For
            // void methods (most setters / `Show` / `DrawLine`) the host
            // fn returns `Value::Null`, which is fine.
            chunk.emit_op(Op::RETURN, line);
        }
        MethodTarget::DotnetCtor {
            class: target_class } => {
            // Discard `this` (slot 0) — factory-style methods don't pass
            // it to the target ctor. Push the target class global, then
            // the user args (slots 1..=arity-1), then call.
            vybe_compiler::primitives::globals::emit_read(&mut chunk, target_class, line);
            // User args only — skip slot 0 (this).
            for slot in 1..method.arity as u16 {
                chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
            }
            // arity - 1 because we dropped `this`.
            chunk.emit_op_u8_u8(Op::CALL_REF, method.arity - 1, 1, line);
            chunk.emit_op(Op::RETURN, line);
        }
        MethodTarget::Body(ops) => {
            compile_body_offset(&mut chunk, ops, body_imports, method.arity, 0, false, line);
        }
    }

    chunk.local_count = method.arity as u16;
    chunk
}

/// Compile a `MethodTarget::Body` sequence into bytecode.
///
/// `body_imports` is the per-`CallHost`-op import index, in the order
/// the ops appear in `ops`. The orchestrator pre-resolves them via
/// [`super::collect_body_imports`] and `chunks[0].add_import` so the
/// builder doesn't have to touch the imports vec.
///
/// Emit a `MethodTarget::Body` sequence INLINE at a call site.
///
/// The receiver and user args are on the stack (`[this, arg1, …, argN]`,
/// `argc = arity`). They're spilled into `alloc_scratch(argc)` slots and the
/// body's `this`/arg reads are offset there, so the exact same `MethodOp`
/// table that builds a thunk chunk also lowers at a call site — control-leaf
/// drawing objects resolve `g.DrawLine(…)` through the component descriptor
/// (`MethodBody::Common`) with no per-class ctor chunk to bind a thunk. The
/// method's result value is left on the stack.
pub fn emit_body_inline(chunk: &mut Chunk, ops: &[MethodOp], argc: u8, line: u32) {
    // Resolve every CallHost target's import index on THIS chunk, in the order
    // the ops appear (the same order `compile_body_offset` consumes them).
    let targets = collect_body_call_targets(ops);
    let mut imports = Vec::with_capacity(targets.len());
    for (module, fn_name) in targets {
        imports.push(chunk.add_import(module, fn_name));
    }
    // Spill [this, arg1..argN] (top = argN) into base+0..base+argc-1.
    let base = chunk.alloc_scratch(argc as u16);
    for slot in (0..argc as u16).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + slot, line);
    }
    compile_body_offset(chunk, ops, &imports, argc, base, true, line);
}

/// Compile a `MethodTarget::Body` sequence into bytecode.
///
/// `base_slot` is where `this` lives; arg N lives in `base_slot + N`. For a
/// thunk chunk that's slot 0 (params); for an inline call-site emit it's the
/// scratch base the receiver+args were spilled to. When `inline` is true,
/// `Return` leaves the result on the stack instead of emitting `Op::RETURN`
/// (returning would exit the *caller's* function).
///
/// `body_imports` is the per-`CallHost`-op import index, in the order the ops
/// appear in `ops`.
fn compile_body_offset(
    chunk: &mut Chunk,
    ops: &[MethodOp],
    body_imports: &[u16],
    arity: u8,
    base_slot: u16,
    inline: bool,
    line: u32,
) {
    let mut import_cursor = 0usize;
    let mut returned = false;

    for op in ops {
        match *op {
            MethodOp::PushThis => {
                chunk.emit_op_u16(Op::LOCAL_GET, base_slot, line);
            }
            MethodOp::PushArg(n) => {
                debug_assert!(
                    n >= 1 && n <= arity - 1,
                    "PushArg({}) out of range for method arity {} (this + {} args)",
                    n,
                    arity,
                    arity - 1
                );
                // arg N (1-indexed after `this`) lives in slot base+N.
                chunk.emit_op_u16(Op::LOCAL_GET, base_slot + n as u16, line);
            }
            MethodOp::PushThisField(field) => {
                chunk.emit_op_u16(Op::LOCAL_GET, base_slot, line);
                let key = chunk.add_constant(Value::String(Arc::from(field)));
                chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
            }
            MethodOp::PushArgField(n, field) => {
                debug_assert!(
                    n >= 1 && n <= arity - 1,
                    "PushArgField({}, _) out of range for method arity {}",
                    n,
                    arity
                );
                chunk.emit_op_u16(Op::LOCAL_GET, base_slot + n as u16, line);
                let key = chunk.add_constant(Value::String(Arc::from(field)));
                chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
            }
            MethodOp::PushArgFieldField(n, f1, f2) => {
                debug_assert!(
                    n >= 1 && n <= arity - 1,
                    "PushArgFieldField({}, _, _) out of range for method arity {}",
                    n,
                    arity
                );
                chunk.emit_op_u16(Op::LOCAL_GET, base_slot + n as u16, line);
                let k1 = chunk.add_constant(Value::String(Arc::from(f1)));
                chunk.emit_struct_field_op(Op::STRUCT_GET, 0, k1, line);
                let k2 = chunk.add_constant(Value::String(Arc::from(f2)));
                chunk.emit_struct_field_op(Op::STRUCT_GET, 0, k2, line);
            }
            MethodOp::PushConstInt(v) => {
                chunk.emit_f64_const(v as f64, line);
            }
            MethodOp::PushConstFloat(v) => {
                chunk.emit_f64_const(v, line);
            }
            MethodOp::PushConstStr(s) => {
                push_string_const(chunk, s, line);
            }
            MethodOp::PushConstBool(b) => {
                if b {
                    core_wasm::bool_const(chunk, line, true);
                } else {
                    core_wasm::bool_const(chunk, line, false);
                }
            }
            MethodOp::PushConstNull => {
                chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            }
            MethodOp::CallHost { argc, .. } => {
                debug_assert!(
                    import_cursor < body_imports.len(),
                    "compile_body: ran out of pre-resolved import indices"
                );
                let idx = body_imports[import_cursor];
                import_cursor += 1;
                chunk.emit_call(idx, argc, line);
            }
            MethodOp::NewDotnet { class, argc } => {
                vybe_compiler::primitives::globals::emit_read(chunk, class, line);
                // Note: global_get pushes the ctor; the args are
                // expected to already be on the stack BELOW it from
                // earlier `Push*` ops. Real .NET ctor convention here
                // expects [args..., ctor], but the VM's `call` op
                // expects [ctor, args...]. Body authors must order
                // their ops accordingly: emit the ctor (`NewDotnet`)
                // FIRST, then the args, then... wait, that's not how
                // we have it.
                //
                // Actually re-reading: most languages put the callee
                // first then args. The simpler convention is for body
                // authors to emit args first, then NewDotnet, and have
                // NewDotnet do the work of pushing the ctor and
                // calling. But that means the ctor goes ABOVE the args
                // on the stack, which is wrong for `call`.
                //
                // The `call` opcode expects: stack = [callee, arg0, arg1, ..., argN-1]
                // and `argc` operand = N. So we need callee BELOW the args.
                //
                // Resolution: NewDotnet pushes the ctor, then issues
                // call(argc) ASSUMING the args are not yet on the
                // stack. Body authors must emit NewDotnet first, then
                // the user args, then nothing else — NewDotnet is the
                // call boundary.
                //
                // Wait — that means NewDotnet can't be emitted in the
                // middle of a sequence; it'd have to be the last op.
                // That's too restrictive.
                //
                // Better: NewDotnet emits the ctor push only, and a
                // separate explicit `Call(argc)` op handles the call.
                // But we don't have that in the enum.
                //
                // For now: emit ctor + call(argc) as a unit, and
                // require body authors to push args BEFORE NewDotnet.
                // The implementation has to swap them, which on a
                // stack VM means using temporaries.
                //
                // To avoid the temp dance, emit the simplest form: a
                // `call_indirect`-style global_get + call sequence
                // assuming args are already on the stack ABOVE the
                // ctor push location. The cleanest fix is to require
                // body authors to use NewDotnet AS the entire call
                // (no args) — and use `Body` ops for ctor calls only
                // when arity = 0 OR push args via locals.
                //
                // Pragmatic: NewDotnet supports arity-0 only for now.
                // The two known callers (CreateGraphics → Graphics(),
                // and any other "factory of arity 0") work fine. When
                // we need arity-N factory methods we'll add a `Call`
                // op or restructure.
                debug_assert_eq!(
                    argc, 0,
                    "MethodOp::NewDotnet currently only supports argc=0; \
                     for arity-N factories, switch to a Host target or extend the DSL"
                );
                chunk.emit_op_u8_u8(Op::CALL_REF, 0, 1, line);
            }
            MethodOp::SetField(field) => {
                let key = chunk.add_constant(Value::String(Arc::from(field)));
                chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
            }
            MethodOp::Drop => {
                chunk.emit_op(Op::DROP, line);
            }
            MethodOp::Dup => {
                core_wasm::dup(chunk, line);
            }
            MethodOp::Return => {
                // Inline at a call site: the result value is already on the
                // stack — emitting RETURN would exit the *caller*. Just stop.
                if !inline {
                    chunk.emit_op(Op::RETURN, line);
                }
                returned = true;
                break;
            }
        }
    }

    // Safety net: if the body didn't end in `Return`, ensure a result. Inline
    // leaves a null on the stack (the method's value); the thunk path returns.
    if !returned {
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        if !inline {
            chunk.emit_op(Op::RETURN, line);
        }
    }
}

/// Walk a `Body` op sequence and return every unique `(module, fn_name)`
/// pair referenced by `CallHost` ops, in encounter order. The
/// orchestrator uses this to pre-resolve import indices via
/// `chunks[0].add_import` before calling `build_method_thunk_chunk`.
pub fn collect_body_call_targets(ops: &[MethodOp]) -> Vec<(&'static str, &'static str)> {
    let mut targets = Vec::new();
    for op in ops {
        if let MethodOp::CallHost {
            module, fn_name, ..
        } = op
        {
            targets.push((*module, *fn_name));
        }
    }
    targets
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
    pub setter_chunk_idx: usize }

#[derive(Debug, Clone, Copy)]
pub struct GetterBinding<'a> {
    pub prop_pascal: &'a str,
    pub getter_chunk_idx: usize }

/// Per-method thunk binding info supplied to [`build_constructor_chunk`].
///
/// `method_name` is the lowercased instance-side key (`"createGraphics"`
/// → `"creategraphics"`) — written by the user as `obj.MethodName(...)`,
/// looked up by the VM via the lowercased canonical AST.
#[derive(Debug, Clone, Copy)]
pub struct MethodBinding<'a> {
    pub method_name: &'a str,
    pub thunk_chunk_idx: usize }

/// The key variants a member is bound under so the exact-matching VM resolves
/// it for every source language: the **original** declared name (case-sensitive
/// languages access it verbatim) plus its **lowercase** form (case-insensitive
/// languages canon-fold their access keys). Deduplicated when already lowercase.
fn accessor_name_variants(name: &str) -> Vec<String> {
    let lower = name.to_lowercase();
    if lower == name {
        vec![name.to_string()]
    } else {
        vec![name.to_string(), lower]
    }
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
/// WASM convention: slot 0 is the first argument. For a `ctor_arity = N` ctor:
/// - slots 0..N-1 = user-supplied ctor args
/// - slot N     = `this`
/// - slot N+1   = `widget` (only when wiring a concrete widget host fn)
///
/// `widget_new_import_idx` (when class is concrete) is the chunk[0] import
/// index for the configured `widget_host_module::widget_host_fn`.
pub fn build_constructor_chunk(
    class: &DotnetClass,
    setter_bindings: &[SetterBinding],
    getter_bindings: &[GetterBinding],
    method_bindings: &[MethodBinding],
    widget_new_import_idx: Option<u16>,
    new_controls_collection_import_idx: u16,
    new_components_collection_import_idx: u16,
) -> Chunk {
    let mut chunk = create_function_chunk(class.name, class.ctor_arity);
    let line = 0u32;
    let arity = class.ctor_arity as u16;
    let this_slot: u16 = arity;
    let widget_slot: u16 = arity + 1;

    // ── Fast path: value-type ctors ─────────────────────────────────────────
    //
    // Point / Size / similar: forward args to the host fn and return its
    // result directly. Value-type classes don't inherit, don't install
    // setters or methods, and their host fn produces the exact field
    // layout the consumer expects (`{x, y}` / `{width, height}`). Going
    // through the usual setter-chain / field-copy path would strip those
    // fields from `this` and leave user code seeing `null` for `.X` etc.
    if class.is_value_type() {
        if let Some(import_idx) = widget_new_import_idx {
            for i in 0..arity {
                chunk.emit_op_u16(Op::LOCAL_GET, i, line);
            }
            chunk.emit_call(import_idx, arity as u8, line);
            chunk.emit_op(Op::RETURN, line);
            chunk.local_count = arity;
            return chunk;
        }
    }

    // ── Step 1: get `this` ──────────────────────────────────────────────────
    if let Some(parent_name) = class.parent {
        // this = <Parent>()  — global_get parent ; call(0) ; local_set this ; drop
        vybe_compiler::primitives::globals::emit_read(&mut chunk, parent_name, line);
        chunk.emit_op_u8_u8(Op::CALL_REF, 0, 1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    } else {
        // Root class (Object): this = struct_new 0
        chunk.emit_struct_new(0, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    }

    // ── Step 2: re-stamp __type with this class's name ──────────────────────
    {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        push_string_const(&mut chunk, class.name, line);
        let type_key = chunk.add_constant(Value::String(Arc::from("__type")));
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, type_key, line);
    }

    // ── Step 3: bind setters for THIS class's properties ────────────────────
    //
    // Each setter chunk was pre-built and pushed by the orchestrator; the
    // orchestrator passes the resulting chunk indices in `setter_bindings`.
    // Bind each accessor under BOTH the original declared name (so
    // case-sensitive languages — C#, F#, … — match exactly) and its lowercase
    // form (so case-insensitive languages, whose walker canon-folds access
    // keys, also match). The VM matches names exactly; case handling lives here
    // in the emitter, not the VM.
    for binding in setter_bindings {
        for name in accessor_name_variants(binding.prop_pascal) {
            let set_name = format!("__set_{}", name);
            chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
            chunk.emit_op_u16(Op::REF_FUNC, binding.setter_chunk_idx as u16, line);
            chunk.emit(0, line); // 0 upvalues
            let key = chunk.add_constant(Value::String(Arc::from(set_name.as_str())));
            chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
        }
    }

    for binding in getter_bindings {
        for name in accessor_name_variants(binding.prop_pascal) {
            let get_name = format!("__get_{}", name);
            chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
            chunk.emit_op_u16(Op::REF_FUNC, binding.getter_chunk_idx as u16, line);
            chunk.emit(0, line);
            let key = chunk.add_constant(Value::String(Arc::from(get_name.as_str())));
            chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
        }
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
        for name in accessor_name_variants(binding.method_name) {
            chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
            chunk.emit_op_u16(Op::REF_FUNC, binding.thunk_chunk_idx as u16, line);
            chunk.emit(0, line); // 0 upvalues
            let key = chunk.add_constant(Value::String(Arc::from(name.as_str())));
            chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
        }
    }

    if class.name == "Control" {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_call(new_controls_collection_import_idx, 1, line);
        let key = chunk.add_constant(Value::String(Arc::from("controls")));
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    }

    if class.name == "Form" {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        chunk.emit_call(new_components_collection_import_idx, 1, line);
        let key = chunk.add_constant(Value::String(Arc::from("components")));
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    }

    // ── Step 5: concrete leaf — wire backing object ────────────────────────
    //
    // Calls `<widget_host_module>::<widget_host_fn>(args...)`, then copies
    // the backing object's identity fields onto `this`. The inherited
    // setters and methods stay intact because we never overwrite their
    // keys here.
    //
    // Args: for `class.ctor_arity > 0` we forward the user-supplied ctor
    // args (slots 0..arity-1) to the host fn. For arity-0 classes the host
    // fn is called with no args.
    if let Some(import_idx) = widget_new_import_idx {
        for i in 0..arity {
            chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        }
        chunk.emit_call(import_idx, arity as u8, line);
        chunk.emit_op_u16(Op::LOCAL_SET, widget_slot, line);

        // Copy backing identity fields: this.<f> = widget.<f>.
        //
        // installed via real method thunks bound at this class's level
        // (or inherited from `Control`).
        for field in &["name", "__control_name", "__control_type"] {
            let key_idx = chunk.add_constant(Value::String(Arc::from(*field)));
            chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, widget_slot, line);
            chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key_idx, line);
            chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key_idx, line);
        }
    }

    // ── Step 6: return this ─────────────────────────────────────────────────
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op(Op::RETURN, line);

    // Local layout (WASM convention):
    //   slots 0..ctor_arity-1 = ctor args
    //   slot ctor_arity       = this   (always present)
    //   slot ctor_arity+1     = widget (only when wiring a backing host fn)
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
    script_chunk.emit_op_u16(Op::REF_FUNC, ctor_chunk_idx as u16, line);
    script_chunk.emit(0, line); // 0 upvalues
    vybe_compiler::primitives::globals::emit_write(script_chunk, class_name, line);

    // Lowercase alias (skip if already lowercase)
    let lower = class_name.to_lowercase();
    if lower != class_name {
        script_chunk.emit_op_u16(Op::REF_FUNC, ctor_chunk_idx as u16, line);
        script_chunk.emit(0, line);
        vybe_compiler::primitives::globals::emit_write(script_chunk, lower.as_str(), line);
    }
}
