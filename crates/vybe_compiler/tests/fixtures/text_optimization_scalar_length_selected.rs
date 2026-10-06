use vybe_runtime::{Chunk, Op};
pub fn emit(chunks: &mut [Chunk], current: usize, line: u32) {
    let from = chunks[current].add_import("ecma:array", "from");
    chunks[current].emit_call(from, 1, line);
    // Existing Array.from always materializes a fresh private Array; keep it
    // exactly and bypass generic String/Map/Object length dispatch only.
    // Preserve the old private array's lifetime (it can retain caller objects).
    let array = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array, line);
    let length = chunks[current].add_import("ecma:array", "length");
    chunks[current].emit_call(length, 1, line);
}
