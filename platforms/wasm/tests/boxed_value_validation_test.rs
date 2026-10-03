//! Validate the actual exported binary, including value-ABI spill boundaries.
use vybe_runtime::{Chunk, Op};

fn validate(chunks: &[Chunk]) {
    let binary = vybe_platform_wasm::write_wasm(chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&binary)
        .expect("exported module must type-check");
}

#[test]
fn function_reference_survives_local_storage_and_argument_spills() {
    for (store, call) in [
        (false, Op::CALL_REF),
        (true, Op::CALL_REF),
        (true, Op::RETURN_CALL_REF),
        (true, Op::RETURN_CALL),
    ] {
        let mut script = Chunk::new("<script>");
        script.local_count = 1;
        script.emit_op_u16(Op::REF_FUNC, 1, 0);
        script.emit(0, 0);
        if store {
            script.emit_op_u16(Op::LOCAL_SET, 0, 0);
            script.emit_op_u16(Op::LOCAL_GET, 0, 0);
        }
        script.emit_f64_const(21.0, 0);
        script.emit_op_u8_u8(call, 1, 1, 0);
        script.emit_op(Op::RETURN, 0);
        let mut callee = Chunk::new("identity");
        callee.arity = 1;
        callee.local_count = 1;
        callee.emit_op_u16(Op::LOCAL_GET, 0, 0);
        callee.emit_op(Op::RETURN, 0);
        validate(&[script, callee]);
    }
}

#[test]
fn outlined_helper_global_survives_local_storage_and_argument_spills() {
    for store in [false, true] {
        let mut script = Chunk::new("<script>");
        script.local_count = 1;
        std::sync::Arc::make_mut(&mut script.globals).push("__vybe_dyneq".into());
        script.emit_op_u32(Op::GLOBAL_GET, 0, 0);
        if store {
            script.emit_op_u16(Op::LOCAL_SET, 0, 0);
            script.emit_op_u16(Op::LOCAL_GET, 0, 0);
        }
        script.emit_f64_const(21.0, 0);
        script.emit_f64_const(21.0, 0);
        script.emit_op_u8_u8(Op::CALL_REF, 2, 1, 0);
        script.emit_op(Op::RETURN, 0);
        let mut callee = Chunk::new("__stdlib_dyneq");
        callee.arity = 2;
        callee.local_count = 2;
        callee.emit_op_u16(Op::LOCAL_GET, 0, 0);
        callee.emit_op(Op::RETURN, 0);
        validate(&[script, callee]);
    }
}

#[test]
fn boxed_local_conditions_are_unboxed_for_branches() {
    for tee in [false, true] {
        let mut script = Chunk::new("<script>");
        script.local_count = 1;
        let block = script.emit_block(0);
        script.emit_i32_const(1, 0);
        if tee {
            script.emit_op_u16(Op::LOCAL_TEE, 0, 0);
        } else {
            script.emit_op_u16(Op::LOCAL_SET, 0, 0);
            script.emit_op_u16(Op::LOCAL_GET, 0, 0);
        }
        script.emit_br_if(0, 0);
        script.emit_end(0);
        script.patch_block(block);
        script.emit_f64_const(42.0, 0);
        script.emit_op(Op::RETURN, 0);
        validate(&[script]);
    }
}
