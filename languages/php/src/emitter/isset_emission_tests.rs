use super::*;
use vybe_runtime::{Value, VM};
use vybe_runtime::registry::emitter_results::{EmitOutcome, EmitRequest, ResultRepresentation};

fn check(values: Vec<Value>, expected: bool, raw: bool) {
    let argc = values.len() as u8;
    let mut chunk = Chunk::new("<script>");
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    for (index, value) in values.into_iter().enumerate() {
        let name = format!("argument{index}");
        let import = chunk.add_import("test:isset", &name);
        chunk.emit_call(import, 0, 0);
        vm.register_host_fn("test:isset", &name, Box::new(move |_, _| value.clone()));
    }
    let mut chunks = vec![chunk];
    if raw {
        assert!(matches!(dispatch_typed(
            "php.isset_all", &mut chunks, 0, argc, 0, EmitRequest::I32Condition,
        ), EmitOutcome::Handled(ResultRepresentation::RawI32Boolean)));
    } else {
        emit_isset_all(&mut chunks[0], argc, 0);
    }
    assert_eq!(chunks[0].local_count, u16::from(argc));
    chunks[0].emit_op(Op::RETURN, 0);
    let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false, include_name_section: true,
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    let result = vm.run(chunks).unwrap();
    if raw {
        assert!(matches!(result, Value::I32(value) if value == i32::from(expected)));
    } else {
        assert!(matches!(result, Value::Bool(value) if value == expected));
    }
}

#[test]
fn isset_condition_and_value_modes_preserve_nullish_semantics() {
    for (values, expected) in [
        (vec![], true),
        (vec![Value::Null], false),
        (vec![Value::Undefined], false),
        (vec![Value::Bool(false)], true),
        (vec![Value::F64(0.0)], true),
        (vec![Value::String("".into())], true),
        (vec![Value::F64(f64::NAN)], true),
        (vec![Value::Bool(false), Value::F64(0.0), Value::String("0".into())], true),
        (vec![Value::Null, Value::Bool(true)], false),
        (vec![Value::Bool(true), Value::Undefined], false),
    ] {
        check(values.clone(), expected, false);
        check(values, expected, true);
    }
}
