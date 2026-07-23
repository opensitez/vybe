//! Ruby `common:ruby.<name>` emit dispatch.
//!
//! Mirrors the per-language dispatch pattern (php/dotnet/fortran/...):
//! `emitter::dispatch::emit_common` delegates every `ruby.*` name here.
//! Arms run as side-effects; `_ => return false` signals "not mine" so
//! the caller can fall through. Returns `true` once an arm matched.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

/// Route a `ruby.<op>` name to its emitter. Returns `true` if handled.
pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        "ruby.dig" => emit_dig(&mut chunks[current], argc, line),
        name if crate::emitter::runtime_adapter::emit_helper(name, chunks, current, argc, line) => {
        }
        _ => return false,
    }
    true
}

/// Emit Ruby `obj.dig(k1, k2, ..., kN)` — variadic property walk with
/// null short-circuit. Stack: `[receiver, k1, k2, ..., kN]` where
/// `argc == N + 1` (receiver + N keys). Stack on exit: `[value_or_null]`.
///
/// Strategy: stash all keys + receiver into temps. Walk one key at a
/// time using `Op::ARRAY_GET` (polymorphic Map/Array/Object). Between
/// each step, check if current value is null and short-circuit out of
/// the wrapping block if so.
fn emit_dig(chunk: &mut Chunk, argc: u8, line: u32) {
    if argc == 0 {
        chunk.emit_op(Op::NULL, line);
        return;
    }
    if argc == 1 {
        // Just the receiver, no keys — return it as-is.
        return;
    }
    let nkeys = argc - 1;
    // Allocate temps: `cur` slot + N key slots.
    let cur_slot = chunk.alloc_scratch(1 + nkeys as u16);
    // Stash keys back-to-front (last key first, ends up in highest slot).
    for i in (0..nkeys).rev() {
        let slot = cur_slot + 1 + i as u16;
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    // Stash receiver as initial `cur`.
    chunk.emit_op_u16(Op::LOCAL_SET, cur_slot, line);

    // Wrapping block: `br_if(0)` exits early when `cur` becomes null.
    let exit_block = chunk.emit_block(line);
    for i in 0..nkeys {
        let key_slot = cur_slot + 1 + i as u16;
        // if cur is null: exit
        chunk.emit_op_u16(Op::LOCAL_GET, cur_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_br_if(0, line);
        // cur = ARRAY_GET(cur, key)
        chunk.emit_op_u16(Op::LOCAL_GET, cur_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op_u16(Op::LOCAL_SET, cur_slot, line);
    }
    chunk.emit_end(line);
    chunk.patch_block(exit_block);

    // Push final result.
    chunk.emit_op_u16(Op::LOCAL_GET, cur_slot, line);
}
