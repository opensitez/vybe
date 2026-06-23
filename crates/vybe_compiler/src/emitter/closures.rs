//! Closure environment emission — shared across all languages.
//!
//! Closures are function pointer + environment pointer (WASM GC style).
//! Captured variables live in a GC array. The inner function receives
//! the env as its first upvalue and accesses captured vars by index
//! via `array.get` / `array.set` — O(1) with no hash overhead.
//!
//! The name→index mapping is compile-time only. At runtime, the env
//! is a plain array: env[0] = first captured var, env[1] = second, etc.

use vybe_bytecode::opcode::Op;
use vybe_bytecode::Chunk;

/// Read a captured variable from the environment array by index.
pub fn emit_env_get(chunk: &mut Chunk, env_slot: u16, index: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, env_slot, line);
    chunk.emit_i32_const(index as i32, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

/// Write a captured variable into the environment array by index.
///
/// Stack before: [new_value]
/// Stack after: [] (value consumed)
pub fn emit_env_set(chunk: &mut Chunk, env_slot: u16, index: u16, line: u32) {
    let tmp = if let Some(s) = chunk.dup_slot {
        s
    } else {
        let s = chunk.local_count;
        chunk.local_count += 1;
        chunk.dup_slot = Some(s);
        s
    };
    chunk.emit_op_u16(Op::LOCAL_SET, tmp, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, env_slot, line);
    chunk.emit_i32_const(index as i32, line);
    chunk.emit_op_u16(Op::LOCAL_GET, tmp, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);
}

/// Create a new environment array and populate with captured variable values.
///
/// `slots` contains the outer function's local slots for each captured var.
/// The index in the vec IS the env index.
///
/// Stack after: [env_array]
pub fn emit_env_new(chunk: &mut Chunk, slots: &[u16], line: u32) {
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, slots.len() as u16, line);
    for (i, slot) in slots.iter().enumerate() {
        chunk.emit_dup(line);
        chunk.emit_i32_const(i as i32, line);
        chunk.emit_op_u16(Op::LOCAL_GET, *slot, line);
        chunk.emit_op(Op::ARRAY_SET, line);
        chunk.emit_op(Op::DROP, line);
    }
}
