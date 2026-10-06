//! Standalone comparisons must declare their unboxing spill local.
use vybe_runtime::{Chunk, Op};

#[test]
fn isolated_i32_comparisons_validate_without_other_temp_reserving_ops() {
    for op in [Op::I32_LT_S, Op::I32_LT_U, Op::I32_GT_S, Op::I32_GT_U,
               Op::I32_LE_S, Op::I32_LE_U, Op::I32_GE_S, Op::I32_GE_U] {
        for (a, b) in [(i32::MIN, i32::MAX), (-1, 0), (0, 0), (1, -1)] {
            let mut chunk = Chunk::new("isolated_i32_comparison");
            chunk.emit_i32_const(a, 1);
            chunk.emit_i32_const(b, 1);
            chunk.emit_op(op, 1);
            chunk.emit_op(Op::RETURN, 1);
            let bytes = vybe_platform_wasm::write_wasm(&[chunk]);
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes)
                .unwrap_or_else(|error| panic!("{op:?}({a}, {b}): {error}"));
        }
    }
}
