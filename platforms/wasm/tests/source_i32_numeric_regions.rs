//! Validate Source value-ABI expressions across stores, calls and barriers.
use vybe_runtime::{Chunk, Op};

fn expression(c: &mut Chunk, depth: u32, seed: u32, barrier: bool) {
    if depth == 0 {
        if seed & 1 == 0 {
            c.emit_i32_const(seed as i32, 0);
        } else {
            c.emit_op_u16(Op::LOCAL_GET, (seed % 3) as u16, 0);
        }
        return;
    }
    expression(c, depth - 1, seed.wrapping_mul(17).wrapping_add(1), barrier);
    if seed % 3 == 0 {
        c.emit_op_u16(Op::LOCAL_TEE, 3, 0);
    }
    if seed % 5 == 0 {
        let import = c.add_import("wasm:js-number", "fromI32");
        c.emit_call(import, 1, 0);
    }
    if seed % 7 == 0 {
        if barrier { c.emit_op(Op::NOP, 0); }
        c.emit_op([Op::I32_EQZ, Op::I32_CLZ, Op::I32_CTZ, Op::I32_POPCNT][(seed % 4) as usize], 0);
    } else {
        expression(c, depth - 1, seed.wrapping_mul(23).wrapping_add(5), barrier);
        if barrier { c.emit_op(Op::NOP, 0); }
        let ops = [Op::I32_EQ, Op::I32_NE, Op::I32_LT_S, Op::I32_LT_U,
            Op::I32_GT_S, Op::I32_GT_U, Op::I32_LE_S, Op::I32_LE_U,
            Op::I32_GE_S, Op::I32_GE_U, Op::I32_ADD, Op::I32_SUB,
            Op::I32_MUL, Op::I32_DIV_S, Op::I32_DIV_U, Op::I32_REM_S,
            Op::I32_REM_U, Op::I32_AND, Op::I32_OR, Op::I32_XOR,
            Op::I32_SHL, Op::I32_SHR_S, Op::I32_SHR_U, Op::I32_ROTL, Op::I32_ROTR];
        c.emit_op(ops[(seed % ops.len() as u32) as usize], 0);
    }
}

#[test]
fn integer_regions_keep_valid_value_abi_at_expression_boundaries() {
    for seed in 0..128 {
        for barrier in [false, true] {
            let mut c = Chunk::new("source_numeric_expression");
            c.arity = 3;
            c.param_count = 3;
            c.local_count = 4;
            expression(&mut c, 4, seed, barrier);
            c.emit_op(Op::RETURN, 0);
            let bytes = vybe_platform_wasm::write_wasm_with_options(&[c],
                vybe_platform_wasm::WasmWriteOptions {
                    include_name_section: false,
                    include_vybe_metadata: false,
                });
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes)
                .unwrap_or_else(|e| panic!("seed={seed}, barrier={barrier}: {e}"));
        }
    }
}
