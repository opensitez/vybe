// Frozen values emitter before eliminating the entries/projection pass.
use vybe_runtime::{Chunk, Op};
pub fn emit_values(chunks: &mut [Chunk], current: usize, line: u32) {
    // [dict] → entries → [[k,v], ...]
    let entries_fn = chunks[current].add_import("ecma:object", "entries");
    chunks[current].emit_call(entries_fn, 1, line);

    // [[k,v], ...] → call __vybe_dict_values_from_entries → [v0, v1, ...]
    let entries_local = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries_local, line);
    vybe_compiler::primitives::globals::emit_read(
        &mut chunks[current],
        "__vybe_dict_values_from_entries",
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, entries_local, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(&mut chunks[current], 1, line);
}
