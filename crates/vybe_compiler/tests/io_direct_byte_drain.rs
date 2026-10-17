use vybe_runtime::{Chunk, Op};

#[test]
fn direct_canonical_byte_drain_is_valid_without_staging_copies() {
    let mut chunk = Chunk::new("direct-canonical-byte-drain");
    chunk.memory_min_pages.push(1);
    chunk.emit_i32_const(0, 0);
    vybe_compiler::primitives::io::emit_read_stream_to_bytes(&mut chunk, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut chunks = vec![chunk];
    vybe_compiler::primitives::globals::normalize_global_table(&mut chunks);
    let binary = vybe_platform_wasm::write_wasm(&chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&binary).unwrap();
    let body = wasmparser::Parser::new(0).parse_all(&binary)
        .find_map(|payload| match payload.unwrap() {
            wasmparser::Payload::CodeSectionEntry(body) => Some(body),
            _ => None,
        }).unwrap();
    for operator in body.get_operators_reader().unwrap() {
        assert!(!matches!(operator.unwrap(), wasmparser::Operator::MemoryCopy { .. }),
            "canonical byte drain must not copy through a staging buffer");
    }
}
