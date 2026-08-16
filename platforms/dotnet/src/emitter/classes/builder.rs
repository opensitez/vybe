//! Inline lowering for a `.NET` class method declared as a `MethodTarget::Body`
//! op sequence.
//!
//! This file used to also hand-build per-class ctor / setter / getter / method
//! thunk chunks and install each class as a callable global — a second,
//! adapter-local implementation of what `primitives/classes.rs` already does
//! (parent chaining, `__type` stamping, accessor binding, prototype wiring).
//! Every one of those helpers went unreferenced once construction moved onto
//! the shared tree `CtorSpec` path and control members onto the namespace-tree
//! resolver, so they are gone rather than left as a template for the next
//! adapter to copy. `plib` still carries a sibling copy of the same four
//! functions (also uncalled).
//!
//! What remains is the one thing that is genuinely dotnet's: a `MethodOp` table
//! is DATA the descriptor declares (`Graphics::DrawLine`, the `Pen`/`Brush`
//! transforms), and `emit_body_inline` lowers it at the call site so the
//! drawing objects need no ctor-bound thunk.
//!
//! ## Import indices
//!
//! Chunks built here must stay FREE of local imports: baked script-table
//! indices are ambiguous once a chunk has its own import table (the normalizer
//! remaps local-first), so strings are pushed as pool constants via
//! [`push_string_const`] — NEVER `emit_string_const`, which secretly adds a
//! `wasm:string-constants` import.

use std::sync::Arc;
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::MethodOp;

/// Push a string as a pool constant. Wrapper chunks must not use
/// `Chunk::emit_string_const` — it registers a `wasm:string-constants`
/// import on the chunk, and any local import shadows same-valued baked
/// `chunks[0]` indices in the import-table normalizer's local-first
/// remap (that collision silently turned every property-setter's
/// `controlSetProperty` call into a string constant).
fn push_string_const(chunk: &mut Chunk, s: &str, line: u32) {
    chunk.emit_string_const(s, line);
}

/// Emit a `MethodTarget::Body` sequence INLINE at a call site.
///
/// The receiver and user args are on the stack (`[this, arg1, …, argN]`,
/// `argc = arity`). They're spilled into `alloc_scratch(argc)` slots and the
/// body's `this`/arg reads are offset there, so control-leaf drawing objects
/// resolve `g.DrawLine(…)` through the component descriptor
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
/// `base_slot` is where `this` lives; arg N lives in `base_slot + N` — the
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
                // The `call` opcode expects stack = [callee, arg0, …, argN-1]
                // with `argc` = N, so the callee has to sit BELOW its args.
                // A `Body` sequence pushes its args before reaching this op,
                // which puts the ctor ABOVE them — so this op is the call
                // boundary and supports arity 0 only. Both known callers are
                // arity-0 factories (`CreateGraphics` → `Graphics()`); an
                // arity-N factory needs a Host target or a separate `Call` op
                // in the DSL rather than a temporaries dance here.
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
            // Every coordinate in this DSL is already an f64 — `PushConstInt`
            // widens, and the drawing arguments arrive as numbers — so these
            // are the plain f64 opcodes with no coercion ladder in front.
            MethodOp::Add => chunk.emit_op(Op::F64_ADD, line),
            MethodOp::Sub => chunk.emit_op(Op::F64_SUB, line),
            MethodOp::Mul => chunk.emit_op(Op::F64_MUL, line),
            MethodOp::Div => chunk.emit_op(Op::F64_DIV, line),
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

/// Walk a `Body` op sequence and return every `(module, fn_name)` pair
/// referenced by `CallHost` ops, in encounter order, so the call site can
/// pre-resolve one import index per op.
fn collect_body_call_targets(ops: &[MethodOp]) -> Vec<(&'static str, &'static str)> {
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
