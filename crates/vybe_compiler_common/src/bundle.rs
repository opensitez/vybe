//! Bundle stdlib into compiled output.
//!
//! Appends stdlib chunks to program. Emits a preamble in the script chunk
//! that creates function refs and stores them in globals (`__vybe_*`).
//!
//! Call sites do: `global_get "__vybe_range"` + `call_ref 3`
//!
//! On Vybe VM, `register_all` overwrites these globals with host fn objects.
//! On any other runtime, the globals hold the bundled stdlib function refs.
//! One binary, works everywhere, no unresolvable imports.

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use crate::stdlib::build_stdlib;
use std::rc::Rc;

/// Mapping from stdlib chunk name to the global name used at call sites.
const MAPPINGS: &[(&str, &str)] = &[
    ("__stdlib_range",      "__vybe_range"),
    ("__stdlib_sorted",     "__vybe_sorted"),
    ("__stdlib_reversed",   "__vybe_reversed"),
    ("__stdlib_enumerate",  "__vybe_enumerate"),
    ("__stdlib_zip",        "__vybe_zip"),
    ("__stdlib_sum",        "__vybe_sum"),
    ("__stdlib_min",        "__vybe_min"),
    ("__stdlib_max",        "__vybe_max"),
    ("__stdlib_pow",        "__vybe_pow"),
    ("__stdlib_tostring",   "__vybe_tostring"),
    ("__stdlib_count",      "__vybe_count"),
    ("__stdlib_isnumeric",  "__vybe_isnumeric"),
    ("__stdlib_splice",     "__vybe_splice"),
    ("__stdlib_floor",      "__vybe_floor"),
    ("__stdlib_slice",      "__vybe_slice"),
    ("__stdlib_keys",       "__vybe_keys"),
    ("__stdlib_hasproperty","__vybe_hasproperty"),
    ("__stdlib_assign",     "__vybe_assign"),
    ("__stdlib_instanceof", "__vybe_instanceof"),
    ("__stdlib_deleteproperty","__vybe_deleteproperty"),
    ("__stdlib_from",       "__vybe_from"),
    ("__stdlib_redim",      "__vybe_redim"),
    ("__stdlib_slicestep",  "__vybe_slicestep"),
    ("__stdlib_dynmul",     "__vybe_dynmul"),
];

/// Additional alias mappings: import name → __vybe_* global name.
/// Used by VM import resolution to find stdlib fallbacks for host calls.
pub const IMPORT_ALIASES: &[(&str, &str, &str)] = &[
    ("vybe:convert", "toString", "__vybe_tostring"),
    ("vybe:string",  "count",    "__vybe_count"),
    ("vybe:math",    "pow",      "__vybe_pow"),
    ("vybe:array",   "range",    "__vybe_range"),
    ("vybe:array",   "sorted",   "__vybe_sorted"),
    ("vybe:array",   "reversed", "__vybe_reversed"),
    ("vybe:array",   "enumerate","__vybe_enumerate"),
    ("vybe:array",   "zip",      "__vybe_zip"),
    ("vybe:array",   "sum",      "__vybe_sum"),
    ("vybe:array",   "pymin",    "__vybe_min"),
    ("vybe:array",   "pymax",    "__vybe_max"),
    ("vybe:convert", "isNumeric","__vybe_isnumeric"),
    ("vybe:array",   "splice",   "__vybe_splice"),
    ("vybe:math",    "floor",    "__vybe_floor"),
    ("vybe:array",   "slice",    "__vybe_slice"),
    ("vybe:object",  "keys",     "__vybe_keys"),
    ("vybe:object",  "hasProperty", "__vybe_hasproperty"),
    ("vybe:object",  "assign",   "__vybe_assign"),
    ("vybe:object",  "instanceOf", "__vybe_instanceof"),
    ("vybe:object",  "deleteProperty", "__vybe_deleteproperty"),
    ("vybe:array",   "from",     "__vybe_from"),
    ("vybe:array",   "redim",    "__vybe_redim"),
    ("vybe:array",   "sliceStep","__vybe_slicestep"),
    ("vybe:math",    "dynMul",   "__vybe_dynmul"),
];

/// Emit the stdlib preamble at the START of a script chunk.
/// This must be called BEFORE any user code is emitted.
///
/// The preamble emits `ref_func + global_set` for each stdlib function,
/// storing function refs in globals that call sites can use.
///
/// Returns (stdlib_base_idx, mapping) so the caller can append stdlib chunks later.
pub fn emit_stdlib_preamble(script: &mut Chunk, stdlib_base: usize) {
    for (i, &(chunk_name, global_name)) in MAPPINGS.iter().enumerate() {
        let ci = stdlib_base + i;
        // ref_func creates a Function object from chunk index
        script.emit_op_u16(Op::ref_func, ci as u16, 0);
        script.emit(0, 0); // 0 upvalues
        // global_set stores it under the __vybe_* name
        let name_c = script.add_constant(Value::String(Rc::from(global_name)));
        script.emit_op_u16(Op::global_set, name_c, 0);
        script.emit_op(Op::drop, 0);
    }
}

/// Append stdlib chunks to program chunks. Call AFTER compilation is done.
pub fn append_stdlib_chunks(program_chunks: &mut Vec<Chunk>) {
    let stdlib = build_stdlib();
    program_chunks.extend(stdlib.chunks);
}

/// Emit a call to a stdlib/vybe function.
/// IMPORTANT: push the function ref FIRST (via global_get), then push args, then call_ref.
/// The call convention is: [func_ref, arg0, arg1, ...] on stack.
///
/// Usage in compiler:
///   emit_call_push_func(chunk, "__vybe_range", 0);  // push func ref
///   compile_expr(arg0);                               // push start
///   compile_expr(arg1);                               // push stop
///   compile_expr(arg2);                               // push step
///   emit_call_invoke(chunk, 3, 0);                    // call_ref 3
pub fn emit_call_push_func(chunk: &mut Chunk, global_name: &str, line: u32) {
    let name_c = chunk.add_constant(Value::String(Rc::from(global_name)));
    chunk.emit_op_u16(Op::global_get, name_c, line);
}

/// Emit call_ref after func + args are on stack.
pub fn emit_call_invoke(chunk: &mut Chunk, argc: u8, line: u32) {
    chunk.emit_op_u8(Op::call_ref, argc, line);
}

/// Append stdlib chunks to a compiled program and register them as global_inits.
/// Call this at the END of compilation, after all user chunks are finalized.
/// The script chunk (chunks[0]) gets global_inits with RefFunc entries.
pub fn finalize_with_stdlib(chunks: &mut Vec<Chunk>) {
    use vybe_bytecode::chunk::{GlobalInit, ConstExpr};

    let stdlib = crate::stdlib::build_stdlib();
    let stdlib_base = chunks.len();

    for (i, &(chunk_name, global_name)) in MAPPINGS.iter().enumerate() {
        if stdlib.exports.iter().any(|&n| n == chunk_name) {
            chunks[0].global_inits.push(GlobalInit {
                name: global_name.to_string(),
                init: ConstExpr::RefFunc(stdlib_base + i),
            });
        }
    }
    chunks.extend(stdlib.chunks);
}

/// Convenience: emit func_ref push + args already on stack + call_ref.
/// Args must already be on stack BEFORE calling this.
/// This function inserts the func ref below the args using a temp local approach.
/// Actually — we can't easily insert below stack items without a local.
/// Instead, callers should use emit_call_push_func + args + emit_call_invoke.
///
/// For backward compatibility, this emits global_get + call_ref, but callers
/// MUST push args AFTER calling this (which means this only works for 0-arg calls).
/// For multi-arg calls, use the push_func/invoke pair.
pub fn emit_call(chunk: &mut Chunk, global_name: &str, argc: u8, line: u32) {
    let name_c = chunk.add_constant(Value::String(Rc::from(global_name)));
    chunk.emit_op_u16(Op::global_get, name_c, line);
    chunk.emit_op_u8(Op::call_ref, argc, line);
}
