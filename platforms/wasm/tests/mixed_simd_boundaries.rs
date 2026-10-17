use vybe_runtime::{Chunk, Op};

#[test]
fn mixed_value_body_bridges_standard_simd_address_and_mask() {
    for load in [Op::V128_LOAD, Op::V128_LOAD32_ZERO] {
    let mut probe = Chunk::new("mixed_simd_mask");
    probe.arity = 1;
    probe.param_count = 1;
    probe.local_count = 1;
    probe.func_sig = Some(("externref".into(), "externref".into()));
    probe.memory_min_pages = vec![1];
    probe.emit_op_u16(Op::LOCAL_GET, 0, 1);
    probe.emit_op(load, 1);
    probe.emit_leb_u32(0x80, 1);
    probe.emit_leb_u32(0, 1);
    probe.emit_op(Op::V128_CONST, 1);
    for _ in 0..16 { probe.emit(0, 1); }
    probe.emit_op(Op::I8X16_EQ, 1);
    probe.emit_op(Op::I8X16_BITMASK, 1);
    probe.emit_op(Op::RETURN, 1);
    let native = vybe_platform_wasm::write_wasm(&[probe]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&native).unwrap();
    }
}
