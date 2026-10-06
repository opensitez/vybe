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

#[test]
fn gc_array_allocation_and_copy_unbox_all_operands() {
    let mut script = Chunk::new("<script>");
    script.local_count = 2;
    script.emit_i32_const(4, 0);
    script.emit_op_u16(Op::ARRAY_NEW_DEFAULT, 0, 0);
    script.emit_op_u16(Op::LOCAL_SET, 0, 0);
    for value in [10, 20, 30] {
        script.emit_i32_const(value, 0);
    }
    script.emit_op_u16_u16(Op::ARRAY_NEW_FIXED, 0, 3, 0);
    script.emit_op_u16(Op::LOCAL_SET, 1, 0);
    script.emit_op_u16(Op::LOCAL_GET, 0, 0);
    script.emit_i32_const(1, 0);
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_i32_const(0, 0);
    script.emit_i32_const(3, 0);
    script.emit_op(Op::ARRAY_COPY, 0);
    script.emit_op_u16(Op::LOCAL_GET, 0, 0);
    script.emit_i32_const(2, 0);
    script.emit_op(Op::ARRAY_GET, 0);
    script.emit_op(Op::RETURN, 0);
    validate(&[script.clone()]);
    assert_eq!(
        vybe_runtime::VM::new().run(vec![script]).unwrap().as_i32(),
        20
    );
}

#[test]
fn unsigned_numeric_conversions_survive_boxed_locals_and_conditions() {
    for condition in [false, true] {
        let mut script = Chunk::new("<unsigned-conversion>");
        script.local_count = 1;
        script.emit_f64_const(4294967295.0, 0);
        script.emit_op(Op::I32_TRUNC_F64_U, 0);
        script.emit_op_u16(Op::LOCAL_SET, 0, 0);
        script.emit_op_u16(Op::LOCAL_GET, 0, 0);
        script.emit_op(Op::F64_CONVERT_I32_U, 0);
        script.emit_op(Op::I32_TRUNC_F64_U, 0);
        if condition {
            script.emit_if_value(0);
            script.emit_f64_const(1.0, 0);
            script.emit_else(0);
            script.emit_f64_const(0.0, 0);
            script.emit_end(0);
        }
        script.emit_op(Op::RETURN, 0);
        validate(&[script]);
    }
}
