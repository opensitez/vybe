//! .NET `Console.Write` / `Console.WriteLine` — Rust inline emitters.
//!
//! Maps to `wasi:cli.log` like the previous direct host-call pattern,
//! but inserts a .NET-style stringifier first:
//!
//! - `bool` → `"True"` / `"False"` (capitalised per .NET spec, vs.
//!   JS-style lowercase `true`/`false` that the default Display impl
//!   emits).
//! - `null` → `""` (matches `Console.WriteLine((string)null)`).
//! - everything else → `String(v)` via `ecma:string.String`.
//!
//! Without this conversion, `Console.WriteLine(true)` prints `true`
//! and `is_constant_pattern` etc. fail their .NET-shaped assertions.

use std::sync::Arc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

/// `Console.WriteLine(v)` / `Console.Write(v)` — emit the bool/null
/// fixup then dispatch to `wasi:cli.log`. Stack: [v] → [null].
///
/// Single outer block as the structured-control-flow exit; each arm
/// stages its converted string in `result_local` and `br exit`. One
/// log call at the end. Avoids emitting RETURN inside a structured
/// block (which would leak the block label to the caller's
/// `label_stack` — the same trap iter_drain hit).
pub fn emit_console_writeline(chunks: &mut [Chunk], current: usize, line: u32) {
    let log_idx = chunks[0].add_import("wasi:cli", "log");
    let chunk = &mut chunks[current];
    // `__vybe_tostring` (stdlib chunk wired by `bundle::finalize_with_stdlib`)
    // dispatches to user-defined `ToString` / `tostring` methods first,
    // falling back to ECMA `String(value)` for primitives. Without this,
    // `Console.WriteLine(person)` would print `[object]` even when the
    // class declares `public override string ToString() { ... }`.
    let tostring_global = chunk.add_constant(Value::String(Arc::from("__vybe_tostring")));
    let v_local = alloc_local(chunk);
    let result_local = alloc_local(chunk);

    // Stash v.
    chunk.emit_op_u16(Op::LOCAL_SET, v_local, line);
    chunk.emit_op(Op::DROP, line);

    let bool_str = chunk.add_constant(Value::String(Arc::from("boolean")));
    let true_str = chunk.add_constant(Value::String(Arc::from("True")));
    let false_str = chunk.add_constant(Value::String(Arc::from("False")));
    let empty_str = chunk.add_constant(Value::String(Arc::from("")));

    let exit_block = chunk.emit_block(line);

    // Bool branch
    let not_bool = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, v_local, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    chunk.emit_op_u16(Op::CONST, bool_str, line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, v_local, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    let false_path = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.emit_op_u16(Op::CONST, true_str, line);
    let after_true = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(false_path);
    chunk.emit_op_u16(Op::CONST, false_str, line);
    chunk.patch_jump(after_true);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_br(1, line); // exit
    chunk.emit_end(line); chunk.patch_block(not_bool);

    // Null branch — `Console.WriteLine((string)null)` prints "".
    let not_null = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, v_local, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op_u16(Op::CONST, empty_str, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_br(1, line);
    chunk.emit_end(line); chunk.patch_block(not_null);

    // Default: __vybe_tostring(v) — handles primitive coercion and
    // canonical runtime-shape stringification (Date, Map, Set, …).
    // User-defined `ToString` overrides on .NET-shape classes are NOT
    // picked up here yet — that requires routing through
    // `ecma:value.invokeMethod` which is itself a method-dispatch
    // problem (the receiver might not be a class instance). For now,
    // tests that need `Console.WriteLine(p)` to call `p.ToString()`
    // fail with "[object]" — call `Console.WriteLine(p.ToString())`
    // explicitly to get the override.
    chunk.emit_op_u16(Op::GLOBAL_GET, tostring_global, line);
    chunk.emit_op_u16(Op::LOCAL_GET, v_local, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line); chunk.patch_block(exit_block);

    // Single log call with the staged string. Push null after so the
    // call site (which DROPs print results uniformly) sees a value.
    chunk.emit_op_u16(Op::LOCAL_GET, result_local, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, log_idx, line);
    chunk.emit(1, line);
    chunk.emit_op(Op::NULL, line);
}

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let s = chunk.local_count;
    chunk.local_count = s + 1;
    s
}
