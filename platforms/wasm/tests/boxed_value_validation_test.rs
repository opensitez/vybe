//! Validate the actual exported binary, including value-ABI spill boundaries.
use vybe_runtime::{Chunk, Op};

fn validate(chunks: &[Chunk]) {
    let binary = vybe_platform_wasm::write_wasm(chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&binary)
        .expect("exported module must type-check");
}

#[test]
fn void_if_followed_by_local_set_preserves_its_existing_stack_prefix() {
    for condition in [0, 1] {
        let mut chunk = Chunk::new("void-if-stack-prefix");
        chunk.local_count = 2;
        chunk.emit_f64_const(42.0, 0);
        chunk.emit_i32_const(condition, 0);
        chunk.emit_if(0);
        chunk.emit_f64_const(9.0, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
        chunk.emit_else(0);
        chunk.emit_f64_const(10.0, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
        chunk.emit_end(0);
        chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
        chunk.emit_op(Op::RETURN, 0);
        let bytes = vybe_platform_wasm::write_wasm_with_options(&[chunk.clone()],
            vybe_platform_wasm::WasmWriteOptions { include_vybe_metadata: false, include_name_section: true });
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes).unwrap();
        assert_eq!(vybe_runtime::VM::new().run(vec![chunk]).unwrap().as_f64(), 42.0);
    }
}

fn assert_raw_i32_body(binary: &[u8], expected_op: &str, expected_const: i32) {
    let mut calls = 0;
    let mut saw_const = false;
    let mut saw_op = false;
    for payload in wasmparser::Parser::new(0).parse_all(binary) {
        if let wasmparser::Payload::CodeSectionEntry(body) = payload.expect("code section") {
            let mut ops = body.get_operators_reader().expect("operators");
            while !ops.eof() {
                match ops.read().expect("operator") {
                    wasmparser::Operator::Call { .. } => calls += 1,
                    wasmparser::Operator::I32Const { value } if value == expected_const => {
                        saw_const = true;
                    }
                    wasmparser::Operator::I32Add if expected_op == "add" => saw_op = true,
                    wasmparser::Operator::I32Sub if expected_op == "sub" => saw_op = true,
                    wasmparser::Operator::I32Mul if expected_op == "mul" => saw_op = true,
                    _ => {}
                }
            }
        }
    }
    assert_eq!(calls, 0, "raw i32 body must not call box/unbox imports");
    assert!(saw_const, "missing raw i32 literal");
    assert!(saw_op, "missing raw i32 arithmetic operator");
}

#[test]
fn numeric_core_returns_validate_with_shared_fallthrough_epilogue() {
    for result in ["i32", "i64", "f32", "f64"] {
        let mut chunk = Chunk::new("typed-result");
        chunk.func_sig = Some((String::new(), result.into()));
        match result {
            "i32" => chunk.emit_i32_const(42, 0),
            "i64" => chunk.emit_i64_const(42, 0),
            "f32" | "f64" => chunk.emit_op(Op::UNREACHABLE, 0),
            _ => unreachable!(),
        }
        if matches!(result, "i32" | "i64") {
            chunk.emit_op(Op::RETURN, 0);
        }
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);
        validate(&[chunk]);
    }
}

#[test]
fn ordered_multi_integer_results_keep_raw_locals_and_validate() {
    for (signature, first, second) in [
        ("i32,i64", Op::I32_CONST, Op::I64_CONST),
        ("i64,i32", Op::I64_CONST, Op::I32_CONST),
    ] {
        let mut chunk = Chunk::new("multi-integer-result");
        chunk.local_count = 2;
        chunk.func_sig = Some((String::new(), signature.into()));
        if first == Op::I32_CONST {
            chunk.emit_i32_const(1, 0);
        } else {
            chunk.emit_i64_const(1, 0);
        }
        if second == Op::I32_CONST {
            chunk.emit_i32_const(2, 0);
        } else {
            chunk.emit_i64_const(2, 0);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
        chunk.emit_op(Op::RETURN, 0);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);
        let binary = vybe_platform_wasm::write_wasm(&[chunk]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .expect("ordered integer results must validate");
        for payload in wasmparser::Parser::new(0).parse_all(&binary) {
            if let wasmparser::Payload::CodeSectionEntry(body) = payload.expect("code section") {
                let mut ops = body.get_operators_reader().expect("operators");
                while !ops.eof() {
                    assert!(
                        !matches!(
                            ops.read().expect("operator"),
                            wasmparser::Operator::Call { .. }
                        ),
                        "multi-result body must not call box/unbox imports"
                    );
                }
            }
        }
    }
}

#[test]
fn numeric_core_parameters_survive_forwarded_local_copies() {
    for ty in ["i32", "i64", "f32", "f64"] {
        let mut chunk = Chunk::new("numeric-identity");
        chunk.arity = 1;
        chunk.local_count = 3;
        chunk.func_sig = Some((ty.into(), ty.into()));
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, 2, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 2, 0);
        chunk.emit_op(Op::RETURN, 0);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);
        validate(&[chunk]);
    }
}

#[test]
fn i32_core_arithmetic_keeps_numeric_locals_and_constants_raw() {
    for (op, name) in [
        (Op::I32_ADD, "add"),
        (Op::I32_SUB, "sub"),
        (Op::I32_MUL, "mul"),
    ] {
        let mut chunk = Chunk::new("numeric-arithmetic");
        chunk.arity = 1;
        chunk.local_count = 3;
        chunk.func_sig = Some(("i32".into(), "i32".into()));
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
        chunk.emit_i32_const(5, 0);
        chunk.emit_op(op, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, 2, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 2, 0);
        chunk.emit_op(Op::RETURN, 0);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);
        validate(&[chunk.clone()]);
        assert_raw_i32_body(&vybe_platform_wasm::write_wasm(&[chunk]), name, 5);
    }
}

#[test]
fn i32_core_literal_return_uses_the_declared_result_type() {
    let mut chunk = Chunk::new("numeric-literal");
    chunk.local_count = 1;
    chunk.func_sig = Some((String::new(), "i32".into()));
    chunk.emit_i32_const(42, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);
    validate(&[chunk]);
}

#[test]
fn i64_core_arithmetic_keeps_numeric_locals_raw() {
    let mut chunk = Chunk::new("numeric-i64-arithmetic");
    chunk.arity = 1;
    chunk.local_count = 3;
    chunk.func_sig = Some(("i64".into(), "i64".into()));
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    chunk.emit_i64_const(5, 0);
    chunk.emit_op(Op::I64_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 2, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 2, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);
    validate(&[chunk.clone()]);
    let binary = vybe_platform_wasm::write_wasm(&[chunk]);
    let mut calls = 0;
    for payload in wasmparser::Parser::new(0).parse_all(&binary) {
        if let wasmparser::Payload::CodeSectionEntry(body) = payload.expect("code section") {
            let mut ops = body.get_operators_reader().expect("operators");
            while !ops.eof() {
                if matches!(
                    ops.read().expect("operator"),
                    wasmparser::Operator::Call { .. }
                ) {
                    calls += 1;
                }
            }
        }
    }
    assert_eq!(calls, 0, "raw i64 body must not call box/unbox imports");
}

#[test]
fn straight_line_integer_division_and_bitwise_ops_remain_raw() {
    for (ty, op) in [
        ("i32", Op::I32_DIV_S),
        ("i32", Op::I32_AND),
        ("i64", Op::I64_DIV_S),
        ("i64", Op::I64_SHL),
    ] {
        let mut chunk = Chunk::new("numeric-binary-op");
        chunk.arity = 1;
        chunk.local_count = 2;
        chunk.func_sig = Some((ty.into(), ty.into()));
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        if ty == "i32" {
            chunk.emit_i32_const(2, 0);
        } else {
            chunk.emit_i64_const(2, 0);
        }
        chunk.emit_op(op, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
        chunk.emit_op(Op::RETURN, 0);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);
        let binary = vybe_platform_wasm::write_wasm(&[chunk]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .expect("raw integer operation must validate");
        for payload in wasmparser::Parser::new(0).parse_all(&binary) {
            if let wasmparser::Payload::CodeSectionEntry(body) = payload.expect("code section") {
                let mut ops = body.get_operators_reader().expect("operators");
                while !ops.eof() {
                    assert!(
                        !matches!(
                            ops.read().expect("operator"),
                            wasmparser::Operator::Call { .. }
                        ),
                        "{ty} operation must not call box/unbox imports"
                    );
                }
            }
        }
    }
}

#[test]
fn straight_line_integer_predicates_keep_typed_results_raw() {
    for (params, arity, op) in [
        ("i32", 1, Op::I32_EQZ),
        ("i32,i32", 2, Op::I32_GT_S),
        ("i64", 1, Op::I64_EQZ),
        ("i64,i64", 2, Op::I64_GT_S),
    ] {
        let mut chunk = Chunk::new("numeric-predicate");
        chunk.arity = arity;
        let result_slot = u16::from(arity);
        chunk.local_count = result_slot + 1;
        chunk.func_sig = Some((params.into(), "i32".into()));
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        if arity == 2 {
            chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
        }
        chunk.emit_op(op, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, result_slot, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, result_slot, 0);
        chunk.emit_op(Op::RETURN, 0);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);
        let binary = vybe_platform_wasm::write_wasm(&[chunk]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .expect("raw integer predicate must validate");
        for payload in wasmparser::Parser::new(0).parse_all(&binary) {
            if let wasmparser::Payload::CodeSectionEntry(body) = payload.expect("code section") {
                let mut ops = body.get_operators_reader().expect("operators");
                while !ops.eof() {
                    assert!(
                        !matches!(
                            ops.read().expect("operator"),
                            wasmparser::Operator::Call { .. }
                        ),
                        "{params} predicate must not call box/unbox imports"
                    );
                }
            }
        }
    }
}

#[test]
fn straight_line_i32_core_arithmetic_preserves_raw_parameters_and_literals() {
    for opcode in [Op::I32_ADD, Op::I32_SUB, Op::I32_MUL] {
        let mut chunk = Chunk::new("numeric-arithmetic");
        chunk.arity = 1;
        chunk.local_count = 3;
        chunk.func_sig = Some(("i32".into(), "i32".into()));
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
        chunk.emit_i32_const(5, 0);
        chunk.emit_op(opcode, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, 2, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 2, 0);
        chunk.emit_op(Op::RETURN, 0);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);
        validate(&[chunk]);
    }
}

#[test]
fn straight_line_i32_core_literal_result_validates() {
    let mut chunk = Chunk::new("numeric-literal");
    chunk.local_count = 1;
    chunk.func_sig = Some((String::new(), "i32".into()));
    chunk.emit_i32_const(42, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);
    validate(&[chunk]);
}

#[test]
fn declared_wasm_parameter_count_controls_local_declarations() {
    for result in ["", "i32"] {
        let mut chunk = Chunk::new("declared-zero-params");
        chunk.arity = 1;
        chunk.local_count = 2;
        chunk.func_sig = Some((String::new(), result.into()));
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
        chunk.emit_op(Op::DROP, 0);
        if result == "i32" {
            chunk.emit_i32_const(1, 0);
        }
        chunk.emit_op(Op::RETURN, 0);
        validate(&[chunk]);
    }
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
        script.globals = std::sync::Arc::new(vybe_runtime::chunk::global_index_space(
            &[], &[], &["__vybe_dyneq".into()],
        ).into());
        let helper_index = script.globals.iter().position(|name| name == "__vybe_dyneq").unwrap();
        script.emit_op_u32(Op::GLOBAL_GET, helper_index as u32, 0);
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

#[test]
fn native_i64_conversions_and_predicates_cross_value_abi() {
    for (truncate, convert, value) in [
        (Op::I64_TRUNC_F64_S, Op::F64_CONVERT_I64_S, -4294967296.0),
        (
            Op::I64_TRUNC_F64_U,
            Op::F64_CONVERT_I64_U,
            9223372036854775808.0,
        ),
    ] {
        for condition in [false, true] {
            let mut script = Chunk::new("<i64-conversion>");
            let word = script.i64_scratch();
            script.emit_f64_const(value, 0);
            script.emit_op(truncate, 0);
            script.emit_op_u16(Op::LOCAL_SET, word, 0);
            script.emit_op_u16(Op::LOCAL_GET, word, 0);
            script.emit_op(convert, 0);
            script.emit_op(truncate, 0);
            if condition {
                script.emit_op(Op::I64_EQZ, 0);
                script.emit_if_value(0);
                script.emit_f64_const(1.0, 0);
                script.emit_else(0);
                script.emit_f64_const(0.0, 0);
                script.emit_end(0);
            } else {
                script.emit_op(convert, 0);
            }
            script.emit_op(Op::RETURN, 0);
            validate(&[script.clone()]);
            let result = vybe_runtime::VM::new().run(vec![script]).unwrap();
            if condition {
                assert_eq!(result.as_i32(), 0);
            } else {
                assert_eq!(result, vybe_runtime::value::Value::F64(value));
            }
        }
    }
}

#[test]
fn exported_i64_truncation_retains_invalid_conversion_traps() {
    for (truncate, convert) in [
        (Op::I64_TRUNC_F64_S, Op::F64_CONVERT_I64_S),
        (Op::I64_TRUNC_F64_U, Op::F64_CONVERT_I64_U),
    ] {
        for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            let mut script = Chunk::new("<invalid-i64-conversion>");
            script.emit_f64_const(value, 0);
            script.emit_op(truncate, 0);
            script.emit_op(convert, 0);
            script.emit_op(Op::RETURN, 0);
            validate(&[script.clone()]);
            assert!(vybe_runtime::VM::new().run(vec![script]).is_err());
        }
    }
}

#[test]
fn i64_predicate_can_return_with_declared_i32_signature() {
    let mut script = Chunk::new("<native-i32-result>");
    script.func_sig = Some((String::new(), "i32".into()));
    script.emit_f64_const(0.0, 0);
    script.emit_op(Op::I64_TRUNC_F64_S, 0);
    script.emit_op(Op::I64_EQZ, 0);
    script.emit_op(Op::RETURN, 0);
    validate(&[script.clone()]);
    assert_eq!(
        vybe_runtime::VM::new().run(vec![script]).unwrap().as_i32(),
        1
    );
}
