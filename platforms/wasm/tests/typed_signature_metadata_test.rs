use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

fn typed_chunks() -> [Chunk; 2] {
    let mut i32_fn = Chunk::new("typed-add");
    i32_fn.arity = 1;
    i32_fn.local_count = 1;
    i32_fn.func_sig = Some(("i32".into(), "i32".into()));
    i32_fn.emit_op_u16(Op::LOCAL_GET, 0, 0);
    i32_fn.emit_i32_const(5, 0);
    i32_fn.emit_op(Op::I32_ADD, 0);
    i32_fn.emit_op(Op::RETURN, 0);
    i32_fn.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    i32_fn.emit_op(Op::RETURN, 0);

    let mut predicate = Chunk::new("typed-compare");
    predicate.arity = 2;
    predicate.local_count = 2;
    predicate.func_sig = Some(("i64,i64".into(), "i32".into()));
    predicate.emit_op_u16(Op::LOCAL_GET, 0, 0);
    predicate.emit_op_u16(Op::LOCAL_GET, 1, 0);
    predicate.emit_op(Op::I64_GT_S, 0);
    predicate.emit_op(Op::RETURN, 0);
    predicate.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    predicate.emit_op(Op::RETURN, 0);

    [i32_fn, predicate]
}

#[test]
fn typed_i32_parameter_enters_boxed_control_flow() {
    let mut chunk = Chunk::new("typed-if");
    chunk.arity = 1;
    chunk.local_count = 2;
    chunk.func_sig = Some(("i32".into(), "i32".into()));
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    chunk.emit_if(0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk.emit_else(0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk.emit_end(0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let binary = vybe_platform_wasm::write_wasm(&[chunk]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&binary)
        .expect("typed parameter must be boxed before entering value-ABI locals");
}

#[test]
fn boxed_values_can_write_typed_i32_parameters() {
    for write in [Op::LOCAL_SET, Op::LOCAL_TEE] {
        let mut chunk = Chunk::new("typed-param-write");
        chunk.arity = 1;
        chunk.local_count = 2;
        chunk.func_sig = Some(("i32".into(), "i32".into()));
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::DROP, 0);
        chunk.emit_i32_const(7, 0);
        chunk.emit_op_u16(write, 0, 0);
        if write == Op::LOCAL_SET {
            chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
        chunk.emit_op(Op::RETURN, 0);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);

        let binary = vybe_platform_wasm::write_wasm(&[chunk]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .expect("typed parameter writes must preserve the core Wasm type");
    }
}

#[test]
fn direct_i32_parameter_condition_needs_no_boxing_calls() {
    let mut chunk = Chunk::new("typed-condition");
    chunk.arity = 1;
    chunk.local_count = 1;
    chunk.func_sig = Some(("i32".into(), "i32".into()));
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_if(0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk.emit_else(0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk.emit_end(0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let binary = vybe_platform_wasm::write_wasm(&[chunk]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&binary)
        .expect("direct typed condition must remain valid Wasm");
    let calls = wasmparser::Parser::new(0)
        .parse_all(&binary)
        .filter_map(|payload| match payload.expect("module payload") {
            wasmparser::Payload::CodeSectionEntry(body) => Some(
                body.get_operators_reader()
                    .expect("operators")
                    .into_iter()
                    .filter(|op| matches!(op, Ok(wasmparser::Operator::Call { .. })))
                    .count(),
            ),
            _ => None,
        })
        .last()
        .expect("typed function body");
    assert_eq!(calls, 0, "direct i32 conditions need no host conversion");
}

#[test]
fn typed_f64_parameter_crosses_boxed_value_locals() {
    for write in [None, Some(Op::LOCAL_SET), Some(Op::LOCAL_TEE)] {
        let mut chunk = Chunk::new("typed-float-boundary");
        chunk.arity = 1;
        chunk.local_count = 2;
        chunk.func_sig = Some(("f64".into(), "f64".into()));
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::DROP, 0);
        if let Some(op) = write {
            chunk.emit_f64_const(7.0, 0);
            chunk.emit_op_u16(op, 0, 0);
            if op == Op::LOCAL_SET {
                chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
            }
        } else {
            chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
        chunk.emit_op(Op::RETURN, 0);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);

        let binary = vybe_platform_wasm::write_wasm(&[chunk]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .expect("typed f64 parameter and result must cross the boxed ABI");
    }
}

#[test]
fn direct_f64_parameter_return_needs_no_boxing_calls() {
    let mut chunk = Chunk::new("typed-float-return");
    chunk.arity = 1;
    chunk.local_count = 1;
    chunk.func_sig = Some(("f64".into(), "f64".into()));
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let binary = vybe_platform_wasm::write_wasm(&[chunk]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&binary)
        .expect("direct typed f64 return must remain valid Wasm");
}

#[test]
fn typed_f32_parameter_crosses_boxed_value_locals() {
    for write in [None, Some(Op::LOCAL_SET), Some(Op::LOCAL_TEE)] {
        let mut chunk = Chunk::new("typed-single-boundary");
        chunk.arity = 1;
        chunk.local_count = 2;
        chunk.func_sig = Some(("f32".into(), "f32".into()));
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::DROP, 0);
        if let Some(op) = write {
            chunk.emit_f32_const(7.0, 0);
            chunk.emit_op_u16(op, 0, 0);
            if op == Op::LOCAL_SET {
                chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
            }
        } else {
            chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
        chunk.emit_op(Op::RETURN, 0);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);

        let binary = vybe_platform_wasm::write_wasm(&[chunk]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .expect("typed f32 parameter and result must cross the boxed ABI");
    }
}

#[test]
fn direct_f32_return_and_constant_remain_raw() {
    for constant in [false, true] {
        let mut chunk = Chunk::new("typed-single-return");
        chunk.arity = u8::from(!constant);
        chunk.local_count = u16::from(!constant);
        chunk.func_sig = Some((
            if constant {
                String::new()
            } else {
                "f32".into()
            },
            "f32".into(),
        ));
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::DROP, 0);
        if constant {
            chunk.emit_f32_const(7.0, 0);
        } else {
            chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        }
        chunk.emit_op(Op::RETURN, 0);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);

        let binary = vybe_platform_wasm::write_wasm(&[chunk]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .expect("direct typed f32 result must remain raw Wasm");
    }
}

#[test]
fn floating_operations_cross_boxed_value_abi() {
    for (ty, arithmetic, comparison, unary) in [
        (
            "f32",
            [
                Op::F32_ADD,
                Op::F32_SUB,
                Op::F32_MUL,
                Op::F32_DIV,
                Op::F32_MIN,
                Op::F32_MAX,
                Op::F32_COPYSIGN,
            ],
            [
                Op::F32_EQ,
                Op::F32_NE,
                Op::F32_LT,
                Op::F32_GT,
                Op::F32_LE,
                Op::F32_GE,
            ],
            [
                Op::F32_ABS,
                Op::F32_NEG,
                Op::F32_CEIL,
                Op::F32_FLOOR,
                Op::F32_TRUNC,
                Op::F32_NEAREST,
                Op::F32_SQRT,
            ],
        ),
        (
            "f64",
            [
                Op::F64_ADD,
                Op::F64_SUB,
                Op::F64_MUL,
                Op::F64_DIV,
                Op::F64_MIN,
                Op::F64_MAX,
                Op::F64_COPYSIGN,
            ],
            [
                Op::F64_EQ,
                Op::F64_NE,
                Op::F64_LT,
                Op::F64_GT,
                Op::F64_LE,
                Op::F64_GE,
            ],
            [
                Op::F64_ABS,
                Op::F64_NEG,
                Op::F64_CEIL,
                Op::F64_FLOOR,
                Op::F64_TRUNC,
                Op::F64_NEAREST,
                Op::F64_SQRT,
            ],
        ),
    ] {
        let operations = arithmetic
            .into_iter()
            .map(|op| (op, true, ty))
            .chain(comparison.into_iter().map(|op| (op, true, "i32")))
            .chain(unary.into_iter().map(|op| (op, false, ty)));
        for (op, binary, result) in operations {
            for spill in [false, true] {
                let mut chunk = Chunk::new("typed-float-op");
                chunk.arity = if binary { 2 } else { 1 };
                chunk.local_count = u16::from(chunk.arity) + u16::from(spill);
                chunk.func_sig = Some((
                    if binary {
                        format!("{ty},{ty}")
                    } else {
                        ty.into()
                    },
                    result.into(),
                ));
                chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
                chunk.emit_op(Op::DROP, 0);
                chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
                if binary {
                    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
                }
                chunk.emit_op(op, 0);
                if spill {
                    chunk.emit_op_u16(Op::LOCAL_SET, u16::from(chunk.arity), 0);
                    chunk.emit_op_u16(Op::LOCAL_GET, u16::from(chunk.arity), 0);
                }
                chunk.emit_op(Op::RETURN, 0);
                chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
                chunk.emit_op(Op::RETURN, 0);

                let binary = vybe_platform_wasm::write_wasm(&[chunk]);
                wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                    .validate_all(&binary)
                    .unwrap_or_else(|error| {
                        panic!("typed {op} spill={spill} must cross the boxed ABI: {error}")
                    });
            }
        }
    }
}

#[test]
fn float32array_imports_cross_boxed_value_abi() {
    for (method, supplied) in [("get", 2), ("set", 3), ("set", 2)] {
        let mut chunk = Chunk::new(format!("float32array-{method}"));
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_i32_const(0, 0);
        if method == "set" && supplied == 3 {
            chunk.emit_f64_const(1.25, 0);
        }
        let import = chunk.add_import("ecma:float32array", method);
        chunk.emit_call(import, supplied, 0);
        if method == "set" {
            chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        }
        chunk.emit_op(Op::RETURN, 0);

        let binary = vybe_platform_wasm::write_wasm(&[chunk]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .unwrap_or_else(|error| {
                panic!("float32array.{method} argc={supplied} must validate: {error}")
            });
    }
}

#[test]
fn typed_float_import_result_can_return_without_boxing() {
    for (module, result) in [("ecma:float32array", "f32"), ("ecma:float64array", "f64")] {
        let mut chunk = Chunk::new(format!("direct-{result}-import-return"));
        chunk.func_sig = Some(("".into(), result.into()));
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_i32_const(0, 0);
        let import = chunk.add_import(module, "get");
        chunk.emit_call(import, 2, 0);
        chunk.emit_op(Op::RETURN, 0);
        if result == "f32" {
            chunk.emit_f32_const(0.0, 0);
        } else {
            chunk.emit_f64_const(0.0, 0);
        }
        chunk.emit_op(Op::RETURN, 0);

        let binary = vybe_platform_wasm::write_wasm(&[chunk]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .unwrap_or_else(|error| panic!("direct {result} import return: {error}"));
        let calls = wasmparser::Parser::new(0)
            .parse_all(&binary)
            .filter_map(|payload| match payload.expect("module payload") {
                wasmparser::Payload::CodeSectionEntry(body) => Some(
                    body.get_operators_reader()
                        .expect("operators")
                        .into_iter()
                        .filter(|op| matches!(op, Ok(wasmparser::Operator::Call { .. })))
                        .count(),
                ),
                _ => None,
            })
            .sum::<usize>();
        assert_eq!(calls, 3, "{result} direct return must not box and unbox");
    }
}

#[test]
fn float32array_imports_execute_after_wasm_round_trip() {
    let mut chunk = Chunk::new("float32array-round-trip");
    chunk.local_count = 1;
    chunk.emit_i32_const(1, 0);
    let new_array = chunk.add_import("ecma:float32array", "newWithLength");
    chunk.emit_call(new_array, 1, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_f64_const(1.25, 0);
    let set = chunk.add_import("ecma:float32array", "set");
    chunk.emit_call(set, 3, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_i32_const(0, 0);
    let get = chunk.add_import("ecma:float32array", "get");
    chunk.emit_call(get, 2, 0);
    chunk.emit_op(Op::RETURN, 0);

    let binary = vybe_platform_wasm::write_wasm(&[chunk]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&binary)
        .expect("Float32Array round trip must be valid Wasm");
    let decoded = vybe_platform_wasm::read_wasm(&binary).expect("decode Wasm");
    let mut vm = VM::new();
    vybe_runtime::js_builtins::register(&mut vm);
    vybe_platform_ecma::register(&mut vm);
    // The reader's script wrapper discards `_start`'s result; inspect the
    // exported function itself to verify the imported f32 value.
    assert_eq!(
        vm.run(vec![decoded[1].clone()]).expect("run decoded Wasm"),
        Value::F64(1.25)
    );
}

#[test]
fn declared_core_signatures_survive_metadata_round_trip() {
    let [i32_fn, predicate] = typed_chunks();

    let binary = vybe_platform_wasm::write_wasm_with_options(
        &[i32_fn, predicate],
        vybe_platform_wasm::writer::WasmWriteOptions {
            include_vybe_metadata: true,
            include_name_section: false,
        },
    );
    let decoded = vybe_platform_wasm::read_wasm(&binary).expect("decode metadata");
    assert_eq!(decoded[0].func_sig, Some(("i32".into(), "i32".into())));
    assert_eq!(decoded[1].func_sig, Some(("i64,i64".into(), "i32".into())));
    let rewritten = vybe_platform_wasm::write_wasm(&decoded);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&rewritten)
        .expect("round-tripped core functions must remain valid Wasm");
}

#[test]
fn standard_core_signatures_survive_without_metadata() {
    let binary = vybe_platform_wasm::write_wasm(&typed_chunks());
    let decoded = vybe_platform_wasm::read_wasm(&binary).expect("decode standard Wasm");
    assert_eq!(decoded[0].func_sig, Some((String::new(), String::new())));
    assert_eq!(decoded[1].func_sig, Some(("i32".into(), "i32".into())));
    assert_eq!(decoded[2].func_sig, Some(("i64,i64".into(), "i32".into())));
    let rewritten = vybe_platform_wasm::write_wasm(&decoded);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&rewritten)
        .expect("standard core functions must retain their declared types");
}

#[test]
fn standard_core_implicit_result_survives_without_metadata() {
    let binary = [
        0, 97, 115, 109, 1, 0, 0, 0, // core Wasm header
        1, 5, 1, 0x60, 0, 1, 0x7f, // type: () -> i32
        3, 2, 1, 0, // one defined function of that type
        7, 10, 1, 6, 95, 115, 116, 97, 114, 116, 0, 0, // export _start
        10, 6, 1, 4, 0, 0x41, 7, 0x0b, // i32.const 7; end
    ];
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&binary)
        .expect("fixture must be standard Wasm");
    let decoded = vybe_platform_wasm::read_wasm(&binary).expect("decode implicit return");
    assert_eq!(decoded[1].func_sig, Some((String::new(), "i32".into())));
    let rewritten = vybe_platform_wasm::write_wasm(&decoded);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&rewritten)
        .expect("implicit typed result must survive re-emission");
}

#[test]
fn standard_core_float_signature_survives_without_metadata() {
    let mut chunk = Chunk::new("float-id");
    chunk.arity = 1;
    chunk.local_count = 1;
    chunk.func_sig = Some(("f64".into(), "f64".into()));
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let binary = vybe_platform_wasm::write_wasm(&[chunk]);
    let decoded = vybe_platform_wasm::read_wasm(&binary).expect("decode float function");
    assert_eq!(decoded[1].func_sig, Some(("f64".into(), "f64".into())));
    let rewritten = vybe_platform_wasm::write_wasm(&decoded);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&rewritten)
        .expect("float signature must survive re-emission");
}

#[test]
fn standard_core_reference_signatures_survive_without_metadata() {
    for ty in ["funcref", "externref"] {
        let mut chunk = Chunk::new("ref-id");
        chunk.arity = 1;
        chunk.local_count = 1;
        chunk.func_sig = Some((ty.into(), ty.into()));
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        chunk.emit_op(Op::RETURN, 0);

        let binary = vybe_platform_wasm::write_wasm(&[chunk]);
        let decoded = vybe_platform_wasm::read_wasm(&binary).expect("decode reference function");
        assert_eq!(decoded[1].func_sig, Some((ty.into(), ty.into())));
        let rewritten = vybe_platform_wasm::write_wasm(&decoded);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&rewritten)
            .expect("reference signature must survive re-emission");
    }
}

#[test]
fn standard_core_multivalue_and_v128_signatures_survive_without_metadata() {
    let mut chunk = Chunk::new("typed-multivalue");
    chunk.arity = 2;
    chunk.local_count = 2;
    chunk.func_sig = Some(("v128,i64".into(), "i32,externref".into()));
    chunk.emit_op(Op::UNREACHABLE, 0);

    let binary = vybe_platform_wasm::write_wasm(&[chunk]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&binary)
        .expect("declared multivalue signature must be valid Wasm");
    let decoded = vybe_platform_wasm::read_wasm(&binary).expect("decode multivalue function");
    assert_eq!(
        decoded[1].func_sig,
        Some(("v128,i64".into(), "i32,externref".into()))
    );
    let rewritten = vybe_platform_wasm::write_wasm(&decoded);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&rewritten)
        .expect("round-tripped multivalue signature must remain valid Wasm");
}

#[test]
fn standard_core_abstract_reference_signatures_survive_without_metadata() {
    for ty in [
        "nullfuncref",
        "nullexternref",
        "nullref",
        "nullexnref",
        "anyref",
        "eqref",
        "i31ref",
        "structref",
        "arrayref",
        "exnref",
    ] {
        let mut chunk = Chunk::new("gc-ref-type");
        chunk.arity = 1;
        chunk.local_count = 1;
        chunk.func_sig = Some((ty.into(), ty.into()));
        chunk.emit_op(Op::UNREACHABLE, 0);

        let binary = vybe_platform_wasm::write_wasm(&[chunk]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .expect("abstract GC reference type must be valid Wasm");
        let decoded = vybe_platform_wasm::read_wasm(&binary).expect("decode GC reference type");
        assert_eq!(decoded[1].func_sig, Some((ty.into(), ty.into())));
        let rewritten = vybe_platform_wasm::write_wasm(&decoded);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&rewritten)
            .expect("round-tripped GC reference type must remain valid Wasm");
    }
}
