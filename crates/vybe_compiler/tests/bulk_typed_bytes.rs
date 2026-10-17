use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, VM, Value};
fn chunks(fast: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("hex");
    c.arity = 1;
    c.param_count = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let mut chunks = vec![root, c];
    if fast {
        primitives::base64::emit_binary_string_to_uint8_array(&mut chunks, 1, 0);
    } else {
        primitives::base64::emit_binary_string_to_byte_array(&mut chunks, 1, 0);
        let i = chunks[1].add_import("ecma:uint8array", "new");
        chunks[1].emit_call(i, 1, 0);
    }
    chunks[1].emit_op(Op::RETURN, 0);
    primitives::bundle::finalize_with_runtime_helpers(&mut chunks);
    primitives::globals::declare_free_globals(&mut chunks);
    primitives::globals::normalize_global_table(&mut chunks);
    chunks
}
fn worker(fast: bool) -> (VM, Value) {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    let f = vm.run(chunks(fast)).unwrap();
    (vm, f)
}

fn result(w: &mut (VM, Value), input: Value) -> Result<Vec<u8>, String> {
    match w.0.invoke(&w.1, &[input]) {
        Ok(value @ Value::Object(_)) => {
            let Value::Object(o) = &value else {
                unreachable!()
            };
            assert!(matches!(
                o.lock().unwrap().kind,
                vybe_runtime::value::ObjectKind::TypedArray(_)
            ));
            Ok(vybe_platform_ecma::typedarray::bytes_from_value(&value))
        }
        Ok(v) => panic!("unexpected {v:?}"),
        Err(e) => Err(e.message),
    }
}
#[test]
fn bulk_typed_bytes_preserves_utf16_wrapping_traps_and_fresh_arrays() {
    #[cfg(feature = "native-gates")]
    for fast in [false, true] {
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(&chunks(fast)))
            .unwrap();
    }
    let mut workers = [worker(false), worker(true)];
    for text in [
        String::new(),
        "Ā🙂東京\0".into(),
        "a".repeat(8192),
        (0..=65535).filter_map(char::from_u32).collect(),
    ] {
        let output = workers
            .each_mut()
            .map(|w| result(w, Value::String(text.clone().into())));
        assert_eq!(output[0], output[1]);
        assert_eq!(
            output[1],
            Ok(text.encode_utf16().map(|u| u as u8).collect())
        );
    }
    for unit in 0..=65535 {
        if let Some(ch) = char::from_u32(unit) {
            let text = ch.to_string();
            let output = workers
                .each_mut()
                .map(|w| result(w, Value::String(text.clone().into())));
            assert_eq!(output[0], output[1], "unit={unit:04x}");
        }
    }
    for input in [
        Value::Null,
        Value::Undefined,
        Value::Bool(true),
        Value::I32(42),
    ] {
        let output = workers.each_mut().map(|w| result(w, input.clone()));
        assert!(output[0].is_err());
        assert_eq!(output[0], output[1]);
    }
    for text in ["", "🙂"] {
        for fast in [false, true] {
            let (mut vm, f) = worker(fast);
            let a = vm.invoke(&f, &[Value::String(text.into())]).unwrap();
            let b = vm.invoke(&f, &[Value::String(text.into())]).unwrap();
            let (Value::Object(a), Value::Object(b)) = (a, b) else {
                panic!("typed arrays")
            };
            assert!(!std::sync::Arc::ptr_eq(&a, &b));
        }
    }
}
