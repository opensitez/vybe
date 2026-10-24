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
    vybe_platform_wasm::validate_binary(&binary).unwrap();
    let mut first = true;
    vybe_platform_wasm::inspection::CoreModule::new(&binary).unwrap().function_bodies(|body| {
        if first {
            first = false;
            for operator in body.instructions()? {
                let operator = operator?;
                assert!(!(operator.opcode == 0xfc && operator.prefixed_opcode() == Some(10)),
                    "canonical byte drain must not copy through a staging buffer");
            }
        }
        Ok(())
    }).unwrap();
    assert!(!first, "canonical byte drain has a function body");
}
