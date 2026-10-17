use super::*;

// Previous split layout retained only as a size oracle, never in production.
fn split_read(chunk: &mut Chunk, normalize: bool, line: u32) {
    let index = chunk.alloc_scratch(1);
    let pointer = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    chunk.emit_op_u16(Op::LOCAL_SET, pointer, line);
    if normalize { unwrap_pointer_cell(chunk, pointer, line); }
    local(chunk, pointer, line);
    recipes::is_object(chunk, line);
    chunk.emit_if_i32(line);
    let kind = chunk.alloc_scratch(1);
    class_slots::emit_class_get(chunk, ObjSource::Local(pointer),
        &ResolvedSlot::Key("__ref_kind".into()), Dest::Local(kind), line);
    test(chunk, "wasm:js-string", kind, line);
    chunk.emit_if_i32(line);
    local(chunk, kind, line);
    chunk.emit_string_const("carray", line);
    let eq = chunk.add_import("wasm:js-string", "equals");
    chunk.emit_call(eq, 2, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
    chunk.emit_if_value(line);
    let base = chunk.alloc_scratch(1);
    let adjusted = chunk.alloc_scratch(1);
    class_slots::emit_class_get(chunk, ObjSource::Local(pointer),
        &ResolvedSlot::Key("__base".into()), Dest::Local(base), line);
    class_slots::emit_class_get(chunk, ObjSource::Local(pointer),
        &ResolvedSlot::Key("__idx".into()), Dest::Stack, line);
    local(chunk, index, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, adjusted, line);
    emit_storage_read(chunk, base, adjusted, line);
    chunk.emit_else(line);
    emit_storage_read(chunk, pointer, index, line);
    chunk.emit_end(line);
}

#[test]
fn common_storage_tail_reduces_emission_and_scratch_locals() {
    for normalize in [false, true] {
        let mut before = Chunk::new("split");
        let mut after = Chunk::new("shared");
        split_read(&mut before, normalize, 1);
        emit_read_with_cell_normalization(&mut after, 1, normalize);
        assert!(after.code.len() < before.code.len());
        assert!(after.local_count < before.local_count);
        eprintln!("char read normalize={normalize}: bytecode {} -> {}, locals {} -> {}",
            before.code.len(), after.code.len(), before.local_count, after.local_count);
    }
}
