use vybe_compiler::primitives::class_slots::{Dest, ObjSource, ResolvedSlot};
use vybe_compiler::primitives::{class_slots, instructions::host, pointers};
use vybe_runtime::{Chunk, Op};

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn field(chunk: &mut Chunk, source: u16, key: &str, dest: u16, line: u32) {
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(source),
        &ResolvedSlot::Key(key.into()),
        Dest::Local(dest),
        line,
    );
}

/// Byte-buffer memset: preserve the pointer, not the backing array's fill result.
pub(super) fn emit_fill(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let first = c.alloc_scratch(6);
    let (pointer, value, length, base, offset, kind) =
        (first, first + 1, first + 2, first + 3, first + 4, first + 5);
    c.emit_op_u16(Op::LOCAL_SET, length, line);
    c.emit_i32_const(255, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_op_u16(Op::LOCAL_SET, value, line);
    c.emit_op_u16(Op::LOCAL_SET, pointer, line);
    get(c, pointer, line);
    c.emit_op_u16(Op::LOCAL_SET, base, line);
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, offset, line);
    get(c, pointer, line);
    host::emit(c, "wasm:js-number", "test", 1, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    field(c, pointer, pointers::REF_KIND_KEY, kind, line);
    get(c, kind, line);
    host::emit(c, "wasm:js-string", "test", 1, line);
    c.emit_if(line);
    get(c, kind, line);
    c.emit_string_const(pointers::CARRAY_KIND, line);
    host::emit(c, "wasm:js-string", "equals", 2, line);
    c.emit_if(line);
    field(c, pointer, pointers::CARRAY_BASE_KEY, base, line);
    field(c, pointer, pointers::CARRAY_IDX_KEY, offset, line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_end(line);

    get(c, base, line);
    host::emit(c, "wasm:js-number", "test", 1, line);
    c.emit_if(line);
    get(c, base, line);
    get(c, offset, line);
    c.emit_op(Op::I32_ADD, line);
    get(c, value, line);
    get(c, length, line);
    pointers::emit_linear_memory_fill(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op(Op::DROP, line);
    c.emit_else(line);
    get(c, base, line);
    get(c, value, line);
    get(c, offset, line);
    get(c, offset, line);
    get(c, length, line);
    c.emit_op(Op::I32_ADD, line);
    host::emit(c, "ecma:array", "fill", 4, line);
    c.emit_op(Op::DROP, line);
    c.emit_end(line);
    get(c, pointer, line);
}
