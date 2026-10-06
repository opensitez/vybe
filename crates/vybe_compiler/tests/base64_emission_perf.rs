//! Shared binary-string conversion behavior and emission/execution checks.
use vybe_compiler::primitives::{base64, bundle, globals, platforms};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};
use vybex as _;

#[path = "fixtures/concatenating_base64.rs"]
mod baseline;

fn array(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}
fn uint8_array(values: Vec<Value>) -> Value {
    let mut script = Chunk::new("<byte-input>");
    globals::emit_read(&mut script, "input", 0);
    let ctor = script.add_import("ecma:uint8array", "new");
    script.emit_call(ctor, 1, 0);
    script.emit_op(Op::RETURN, 0);
    vm(array(values)).run(finish(vec![script])).unwrap()
}
fn vm(input: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("input", input);
    vm
}
fn finish(mut chunks: Vec<Chunk>) -> Vec<Chunk> {
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn program(old: bool, reverse: bool, slice: Option<(i32, i32)>, calls: usize) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<binary-conversion>")];
    // These locals belong to the caller, before any emitter scratch slots.
    chunks[0].local_count = 3;
    for call in 0..calls {
        globals::emit_read(&mut chunks[0], "input", 0);
        if reverse {
            if old {
                baseline::emit_binary_string_to_byte_array(&mut chunks, 0, 0);
            } else {
                base64::emit_binary_string_to_byte_array(&mut chunks, 0, 0);
            }
        } else if let Some((start, end)) = slice {
            chunks[0].emit_op_u16(Op::LOCAL_SET, 0, 0);
            chunks[0].emit_i32_const(start, 0);
            chunks[0].emit_op_u16(Op::LOCAL_SET, 1, 0);
            chunks[0].emit_i32_const(end, 0);
            chunks[0].emit_op_u16(Op::LOCAL_SET, 2, 0);
            if old {
                baseline::emit_byte_array_slot_to_binary_string(
                    &mut chunks,
                    0,
                    Some(0),
                    Some(1),
                    Some(2),
                    0,
                );
            } else {
                base64::emit_byte_array_slot_to_binary_string(
                    &mut chunks,
                    0,
                    Some(0),
                    Some(1),
                    Some(2),
                    0,
                );
            }
        } else if old {
            baseline::emit_byte_array_to_binary_string(&mut chunks, 0, 0);
        } else {
            base64::emit_byte_array_to_binary_string(&mut chunks, 0, 0);
        }
        if call + 1 != calls {
            chunks[0].emit_op(Op::DROP, 0);
        }
    }
    chunks[0].emit_op(Op::RETURN, 0);
    finish(chunks)
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}
fn items(value: Value) -> Vec<Value> {
    let Value::Object(object) = value else {
        panic!("expected array, got {value:?}")
    };
    let object = object.lock().unwrap();
    match &object.kind {
        ObjectKind::Array(items) => items.clone(),
        ObjectKind::TypedArray(ta) => {
            assert_eq!(ta.elem, vybe_runtime::value::TypedElemKind::U8);
            ta.buffer.lock().unwrap()[ta.byte_offset..ta.byte_offset + ta.length]
                .iter()
                .map(|byte| Value::I32(*byte as i32))
                .collect()
        }
        _ => panic!("expected array"),
    }
}
#[test]
fn byte_array_conversion_preserves_codes_slices_and_input() {
    for values in [
        vec![],
        (0..256).map(Value::I32).collect(),
        vec![
            Value::I32(-1),
            Value::I32(65536),
            Value::F64(f64::NAN),
            Value::I32(256),
            Value::I32(0x20ac),
        ],
    ] {
        for slice in [None, Some((0, 0)), Some((1, 4)), Some((4, 1))] {
            let input = array(values.clone());
            let old = program(true, false, slice, 1);
            let new = program(false, false, slice, 1);
            validate(&old);
            validate(&new);
            let expected = vm(input.clone()).run(old).unwrap();
            let actual = vm(input.clone()).run(new).unwrap();
            assert_eq!(actual.as_str(), expected.as_str(), "{values:?}, {slice:?}");
            // NaN is deliberately included in the conversion fixture.
            assert_eq!(format!("{:?}", items(input)), format!("{values:?}"));
        }
    }
}
#[test]
fn inverse_conversion_preserves_utf16_codes_and_returns_fresh_arrays() {
    for text in ["", "\0Aÿ", "€😀", "hello"] {
        let old = program(true, true, None, 1);
        let new = program(false, true, None, 1);
        validate(&old);
        validate(&new);
        let input = Value::String(text.into());
        let expected = items(vm(input.clone()).run(old).unwrap());
        assert_eq!(
            expected,
            text.encode_utf16()
                .map(|code| Value::I32(code as i32))
                .collect::<Vec<_>>()
        );
        let result = vm(input.clone()).run(new.clone()).unwrap();
        assert_eq!(items(result.clone()), expected);
        let other = vm(input).run(new).unwrap();
        if let (Value::Object(a), Value::Object(b)) = (result, other) {
            assert!(!std::sync::Arc::ptr_eq(&a, &b));
        }
    }
}
#[test]
fn repeated_conversion_is_shared_and_finalization_is_idempotent() {
    for reverse in [false, true] {
        let old = program(true, reverse, None, 20);
        let mut new = program(false, reverse, None, 20);
        let codes: Vec<_> = new.iter().map(|c| c.code.clone()).collect();
        bundle::finalize_with_runtime_helpers(&mut new);
        assert_eq!(
            new.iter().map(|c| c.code.clone()).collect::<Vec<_>>(),
            codes
        );
        validate(&new);
        let old_bytes: usize = old.iter().map(|c| c.code.len()).sum();
        let new_bytes: usize = new.iter().map(|c| c.code.len()).sum();
        println!(
            "binary-conversion reverse={reverse}: bytecode={old_bytes}/{new_bytes} wasm={}/{}",
            vybe_platform_wasm::write_wasm(&old).len(),
            vybe_platform_wasm::write_wasm(&new).len()
        );
        assert!(new_bytes * 3 < old_bytes);
        let input = if reverse {
            Value::String("\0Aÿ".into())
        } else {
            array(vec![Value::I32(0), Value::I32(65), Value::I32(255)])
        };
        let expected = vm(input.clone()).run(old).unwrap();
        let actual = vm(input).run(new).unwrap();
        if reverse {
            assert_eq!(items(actual), items(expected));
        } else {
            assert_eq!(actual.as_str(), expected.as_str());
        }
    }
}
#[test]
#[ignore = "manual interleaved debug execution benchmark"]
fn byte_conversion_execution_benchmark() {
    for size in [256, 8192, 131072] {
        let old = program(true, false, None, 1);
        let new = program(false, false, None, 1);
        let input = array((0..size).map(|i| Value::I32(i % 256)).collect());
        let expected = vm(input.clone())
            .run(old.clone())
            .unwrap()
            .as_str()
            .to_owned();
        let mut timings = [Vec::new(), Vec::new()];
        for iteration in 0..5 {
            for which in if iteration % 2 == 0 { [0, 1] } else { [1, 0] } {
                let mut vm = vm(input.clone());
                let chunks = if which == 0 { old.clone() } else { new.clone() };
                let start = std::time::Instant::now();
                let result = vm.run(chunks).unwrap();
                timings[which].push(start.elapsed().as_micros());
                assert_eq!(result.as_str(), expected);
            }
        }
        for samples in &mut timings {
            samples.sort();
        }
        println!(
            "binary-conversion size={size}: concatenation={}us join={}us speedup={:.2}x",
            timings[0][2],
            timings[1][2],
            timings[0][2] as f64 / timings[1][2] as f64
        );
    }
}

#[test]
fn python_base64_frontend_preserves_byte_arrays() {
    vybe_language_python::register();
    for (expression, input, expected) in [
        (
            "base64.b64encode(input)",
            vec![0, 65, 255],
            vec![65, 69, 72, 47],
        ),
        ("base64.b64encode(input)", vec![], vec![]),
        (
            "base64.b64decode(input)",
            vec![65, 69, 72, 47],
            vec![0, 65, 255],
        ),
    ] {
        let source = format!("import base64\nreturn {expression}\n");
        let module = vybe_language_python::parse(&source).unwrap();
        let profile =
            vybe_compiler::profile::parse_profile(vybe_language_python::profile_source()).unwrap();
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .unwrap();
        let binary = vybe_platform_wasm::write_wasm(&chunks);
        if let Err(error) =
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&binary)
        {
            std::fs::write(
                std::env::temp_dir().join("base64-frontend-invalid.wasm"),
                binary,
            )
            .unwrap();
            panic!("{expression}: {error}");
        }
        let actual = items(
            vm(uint8_array(input.into_iter().map(Value::I32).collect()))
                .run(chunks)
                .unwrap(),
        )
        .iter()
        .map(Value::as_i32)
        .collect::<Vec<_>>();
        assert_eq!(actual, expected, "{expression}");
    }
}

#[test]
fn typed_byte_array_input_keeps_bytes_and_slice_order() {
    let input = uint8_array(vec![
        Value::I32(0),
        Value::I32(65),
        Value::I32(255),
        Value::I32(66),
    ]);
    for (slice, expected) in [(None, "\0AÿB"), (Some((1, 3)), "Aÿ"), (Some((2, 2)), "")] {
        let old = program(true, false, slice, 1);
        let new = program(false, false, slice, 1);
        validate(&new);
        assert_eq!(vm(input.clone()).run(old).unwrap().as_str(), expected);
        assert_eq!(vm(input.clone()).run(new).unwrap().as_str(), expected);
        assert_eq!(
            items(input.clone())
                .iter()
                .map(Value::as_i32)
                .collect::<Vec<_>>(),
            [0, 65, 255, 66]
        );
    }
}

#[test]
#[ignore = "manual export for independent wasm-tools validation"]
fn export_binary_conversion_wasm_fixtures() {
    let dir = std::env::temp_dir().join("vybe-binary-conversion-wasm");
    std::fs::create_dir_all(&dir).unwrap();
    for (name, reverse, slice, calls) in [
        ("whole", false, None, 1),
        ("slice", false, Some((1, 3)), 1),
        ("inverse", true, None, 1),
        ("repeated-whole", false, None, 20),
        ("repeated-inverse", true, None, 20),
    ] {
        std::fs::write(
            dir.join(format!("{name}.wasm")),
            vybe_platform_wasm::write_wasm(&program(false, reverse, slice, calls)),
        )
        .unwrap();
    }
}
