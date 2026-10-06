use vybe_compiler::primitives::canon_marshal::{
    emit_alloc, emit_append_bytes, emit_new_bytes, emit_store_byte_array, emit_store_utf8,
};
use vybe_runtime::{Chunk, Op};

#[test]
fn canonical_allocator_emits_valid_wasm() {
    let mut chunk = Chunk::new("<canonical-allocator-validation>");
    chunk.memory_min_pages.push(1);
    let size = chunk.alloc_scratch(1);
    chunk.emit_i32_const(256 * 1024 + 17, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, size, 0);
    let ptr = emit_alloc(&mut chunk, 0, size);
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut chunks = vec![chunk];
    vybe_compiler::primitives::globals::declare_free_globals(&mut chunks);
    vybe_compiler::primitives::globals::normalize_global_table(&mut chunks);
    let wasm = vybe_platform_wasm::write_wasm(&chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm)
        .unwrap();
}

#[test]
fn canonical_byte_append_emits_valid_wasm() {
    let mut chunk = Chunk::new("<canonical-byte-append-validation>");
    chunk.memory_min_pages.push(1);
    let array = chunk.alloc_scratch(1);
    let ptr = chunk.alloc_scratch(1);
    let len = chunk.alloc_scratch(1);
    emit_new_bytes(&mut chunk, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, array, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, ptr, 0);
    chunk.emit_i32_const(255, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, len, 0);
    emit_append_bytes(&mut chunk, 0, array, ptr, len);
    chunk.emit_op_u16(Op::LOCAL_GET, array, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut chunks = vec![chunk];
    vybe_compiler::primitives::globals::declare_free_globals(&mut chunks);
    vybe_compiler::primitives::globals::normalize_global_table(&mut chunks);
    let wasm = vybe_platform_wasm::write_wasm(&chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm)
        .unwrap();
}

#[test]
fn canonical_utf8_store_emits_valid_wasm() {
    let mut chunk = Chunk::new("<canonical-utf8-store-validation>");
    chunk.memory_min_pages.push(1);
    let (ptr, _) = emit_store_utf8(&mut chunk, 0, |chunk| {
        chunk.emit_string_const("\u{e9}\u{1f642}", 0);
    });
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut chunks = vec![chunk];
    vybe_compiler::primitives::globals::declare_free_globals(&mut chunks);
    vybe_compiler::primitives::globals::normalize_global_table(&mut chunks);
    let wasm = vybe_platform_wasm::write_wasm(&chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm)
        .unwrap();
}

#[test]
fn canonical_byte_array_store_emits_valid_wasm() {
    let mut chunk = Chunk::new("<canonical-byte-array-store-validation>");
    chunk.memory_min_pages.push(1);
    let array = chunk.alloc_scratch(1);
    emit_new_bytes(&mut chunk, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, array, 0);
    let (ptr, _) = emit_store_byte_array(&mut chunk, 0, |chunk| {
        chunk.emit_op_u16(Op::LOCAL_GET, array, 0);
    });
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut chunks = vec![chunk];
    vybe_compiler::primitives::globals::declare_free_globals(&mut chunks);
    vybe_compiler::primitives::globals::normalize_global_table(&mut chunks);
    let wasm = vybe_platform_wasm::write_wasm(&chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm)
        .unwrap();
}
