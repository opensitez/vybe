use vybe_compiler::primitives::{bundle, globals, platforms, strings};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/text_optimization_scalar_length_selected.rs"]
mod candidate;
#[path = "fixtures/text_optimization_scalar_length_frozen.rs"]
mod frozen;
fn ops(c: &Chunk) -> Vec<(usize, Op)> {
    let mut at = 0;
    let mut out = vec![];
    while at < c.code.len() {
        let op = Op::from_code_at(&c.code, at);
        out.push((at, op));
        at += 4 + op.operand_format().size_in(&c.code, at + 4);
    }
    out
}
fn u16_at(c: &Chunk, at: usize) -> usize {
    u16::from_be_bytes(c.code[at + 4..at + 6].try_into().unwrap()) as usize
}
fn vm() -> VM {
    let mut runtime = VM::new();
    platforms::register_platforms_all(&mut runtime);
    runtime
}
fn run(runtime: &mut VM, chunks: Vec<Chunk>) -> Result<Value, String> {
    runtime.run(chunks).map_err(|mut e| {
        for frame in &mut e.call_stack {
            frame.offset = 0;
        }
        e.to_string()
    })
}
fn validate(chunks: &[Chunk]) -> usize {
    let bytes = vybe_platform_wasm::write_wasm(chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    bytes.len()
}
fn finalize(chunks: &mut Vec<Chunk>) {
    bundle::finalize_with_runtime_helpers(chunks);
    globals::declare_free_globals(chunks);
    globals::normalize_global_table(chunks);
}
fn program(selected: bool, base: u16) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<scalar-length>")];
    chunks[0].local_count = base;
    globals::emit_read(&mut chunks[0], "input", 0);
    if selected {
        strings::emit_scalar_length(&mut chunks, 0, 0);
    } else {
        frozen::emit(&mut chunks, 0, 0);
    }
    chunks[0].emit_op(Op::RETURN, 0);
    finalize(&mut chunks);
    chunks
}
fn parsed(body: &str) -> Vec<Chunk> {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_php::register);
    let module = vybe_language_php::parse(&format!("<?php {body}")).unwrap();
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    vybe_compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap()
}
#[test]
fn installed_emission_unicode_arbitrary_inputs_and_native() {
    let mut a = vec![Chunk::new("capture")];
    let mut b = a.clone();
    strings::emit_scalar_length(&mut a, 0, 0);
    candidate::emit(&mut b, 0, 0);
    assert_eq!(a[0].code, b[0].code);
    assert_eq!(a[0].imports, b[0].imports);
    assert_eq!(a[0].local_count, b[0].local_count);
    let old = program(false, 0);
    let new = program(true, 0);
    println!(
        "scalar_length bytes {}→{} ops {}→{} native {}→{} locals {}→{}",
        old[0].code.len(),
        new[0].code.len(),
        ops(&old[0]).len(),
        ops(&new[0]).len(),
        validate(&old),
        validate(&new),
        old[0].local_count,
        new[0].local_count
    );
    let mut inputs = vec![
        Value::Null,
        Value::Undefined,
        Value::I32(1),
        Value::Bool(true),
    ];
    for text in [
        "",
        "a",
        "a😀b",
        "e\u{301}",
        "\0",
        "�",
        "\u{85}\u{2028}\u{2029}",
    ] {
        for chunks in [&old, &new] {
            let mut runtime = vm();
            runtime.set_global("input", Value::String(text.into()));
            assert_eq!(
                run(&mut runtime, chunks.clone()),
                Ok(Value::I32(text.chars().count() as i32))
            );
        }
    }
    inputs.push(Value::Object(vybe_runtime::heap::alloc(
        vybe_runtime::value::Object::new_array(vec![Value::String("😀".into()), Value::Null]),
    )));
    let mut object = vybe_runtime::value::Object::new();
    object.properties.insert("length".into(), Value::F64(2.0));
    object
        .properties
        .insert("0".into(), Value::String("x".into()));
    inputs.push(Value::Object(vybe_runtime::heap::alloc(object)));
    for map in [false, true] {
        let mut object = vybe_runtime::value::Object::new();
        object.kind = if map {
            vybe_runtime::value::ObjectKind::Map(Default::default())
        } else {
            vybe_runtime::value::ObjectKind::Set(Default::default())
        };
        match &mut object.kind {
            vybe_runtime::value::ObjectKind::Map(values) => {
                values.insert(Value::String("key".into()), Value::I32(7));
            }
            vybe_runtime::value::ObjectKind::Set(values) => {
                values.insert(Value::String("member".into()));
            }
            _ => unreachable!(),
        }
        inputs.push(Value::Object(vybe_runtime::heap::alloc(object)));
    }
    let bytes = std::sync::Arc::new(std::sync::Mutex::new(vec![0, 127, 255]));
    let mut buffer = vybe_runtime::value::Object::new();
    buffer.kind =
        vybe_runtime::value::ObjectKind::ArrayBuffer(vybe_runtime::value::ArrayBufferState {
            bytes: bytes.clone(),
            max_byte_length: 3,
            resizable: false,
            detached: false,
            shared: false,
        });
    let mut typed = vybe_runtime::value::Object::new();
    typed.kind =
        vybe_runtime::value::ObjectKind::TypedArray(vybe_runtime::value::TypedArrayState {
            elem: vybe_runtime::value::TypedElemKind::U8,
            buffer: bytes.clone(),
            buffer_obj: vybe_runtime::heap::alloc(buffer),
            byte_offset: 0,
            length: 3,
        });
    inputs.push(Value::Object(vybe_runtime::heap::alloc(typed)));
    for input in inputs {
        let mut before = vm();
        let mut after = vm();
        before.set_global("input", input.clone());
        after.set_global("input", input);
        assert_eq!(run(&mut before, old.clone()), run(&mut after, new.clone()));
    }
    assert_eq!(
        *bytes.lock().unwrap(),
        vec![0, 127, 255],
        "Array.from never mutates caller-owned typed storage"
    );
}
#[test]
fn arraylike_length_fields_keep_existing_raw_host_contract() {
    for length in [
        Value::F64(-1.0),
        Value::F64(f64::NAN),
        Value::F64(2.9),
        Value::String("2".into()),
        Value::Null,
        Value::Undefined,
        Value::Object(vybe_runtime::heap::alloc(vybe_runtime::value::Object::new())),
    ] {
        let mut object = vybe_runtime::value::Object::new();
        object.properties.insert("length".into(), length.clone());
        object
            .properties
            .insert("0".into(), Value::String("😀".into()));
        object.properties.insert("1".into(), Value::Null);
        let expected = Value::I32(length.as_f64().max(0.0) as usize as i32);
        let input = Value::Object(vybe_runtime::heap::alloc(object));
        for selected in [false, true] {
            let chunks = program(selected, 0);
            validate(&chunks);
            let mut runtime = vm();
            runtime.set_global("input", input.clone());
            assert_eq!(
                run(&mut runtime, chunks),
                Ok(expected.clone()),
                "existing raw length field contract selected={selected}"
            );
        }
    }
}
#[test]
fn installed_actual_php_scalar_source_and_surplus_option_order() {
    for (expression, expected) in [
        ("'a😀b'", 3),
        ("''", 0),
        ("'é'", 2),
        ("null", 0),
        ("[1,2]", 2),
    ] {
        let old = parsed(&format!(
            "function scalar_subject($v){{return mb_strlen($v);}} return scalar_subject({expression});"
        ));
        validate(&old);
        for chunks in [old] {
            assert_eq!(run(&mut vm(), chunks), Ok(Value::I32(expected)));
        }
    }
    let old = parsed(
        "$order='';function mark($v){global $order;$order.='e';return $v;}function scalar_subject($v){return mb_strlen($v,mark('UTF-8'));}$n=scalar_subject('a😀b');return $order.'|'.$n;",
    );
    validate(&old);
    for chunks in [old] {
        assert_eq!(run(&mut vm(), chunks), Ok(Value::String("e|3".into())));
    }
}
#[test]
fn temporary_array_objects_remain_alive_until_same_frame_cleanup() {
    use std::sync::{Arc, Mutex};
    for selected in [false, true] {
        let mut runtime = vm();
        let weak = Arc::new(Mutex::new(
            None::<std::sync::Weak<Mutex<vybe_runtime::value::Object>>>,
        ));
        let capture = weak.clone();
        runtime.register_host_fn(
            "test:scalar",
            "supply",
            Box::new(move |_, _| {
                let target = vybe_runtime::heap::alloc(vybe_runtime::value::Object::new());
                *capture.lock().unwrap() = Some(Arc::downgrade(&target));
                Value::Object(vybe_runtime::heap::alloc(
                    vybe_runtime::value::Object::new_array(vec![Value::Object(target)]),
                ))
            }),
        );
        let capture = weak.clone();
        runtime.register_host_fn(
            "test:scalar",
            "observe",
            Box::new(move |_, _| {
                Value::Bool(
                    capture
                        .lock()
                        .unwrap()
                        .as_ref()
                        .unwrap()
                        .upgrade()
                        .is_some(),
                )
            }),
        );
        let mut chunks = vec![Chunk::new("temporary array liveness")];
        let supply = chunks[0].add_import("test:scalar", "supply");
        chunks[0].emit_call(supply, 0, 0);
        if selected {
            strings::emit_scalar_length(&mut chunks, 0, 0);
        } else {
            frozen::emit(&mut chunks, 0, 0);
        }
        chunks[0].emit_op(Op::DROP, 0);
        let observe = chunks[0].add_import("test:scalar", "observe");
        chunks[0].emit_call(observe, 0, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        finalize(&mut chunks);
        validate(&chunks);
        assert_eq!(
            run(&mut runtime, chunks),
            Ok(Value::Bool(true)),
            "array holds supplied object selected={selected}"
        );
        let mut control = vec![Chunk::new("liveness control without retaining array")];
        let supply = control[0].add_import("test:scalar", "supply");
        control[0].emit_call(supply, 0, 0);
        control[0].emit_op(Op::DROP, 0);
        let observe = control[0].add_import("test:scalar", "observe");
        control[0].emit_call(observe, 0, 0);
        control[0].emit_op(Op::RETURN, 0);
        finalize(&mut control);
        assert_eq!(
            run(&mut runtime, control),
            Ok(Value::Bool(false)),
            "observer must detect actual early release when private array is absent"
        );
    }
}
#[test]
fn largest_old_frame_remains_accepted() {
    let old = program(false, u16::MAX - 3);
    let new = program(true, u16::MAX - 3);
    assert_eq!(old[0].local_count, u16::MAX);
    assert_eq!(new[0].local_count, u16::MAX - 2);
    for chunks in [&old, &new] {
        let mut runtime = vm();
        runtime.set_global("input", Value::String("a😀b".into()));
        assert_eq!(run(&mut runtime, chunks.to_vec()), Ok(Value::I32(3)));
        let bytes = vybe_platform_wasm::write_wasm(chunks);
        let error = match wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes)
        {
            Ok(_) => panic!("native MAX50000 cap expected"),
            Err(e) => e,
        };
        assert!(error.message().contains("locals"));
    }
    println!(
        "scalar_length maximal old frame output passed both; native local-count implementation cap, not a native pass"
    );
}

#[test]
fn installed_producer_metadata_and_complete_module_equal_reviewed_selection() {
    for base in [0u16, 1, 129] {
        for own in [false, true] {
            for prototype in [false, true] {
                let mut production = vec![Chunk::new("installed scalar producer")];
                production[0].local_count = base;
                production[0].module_instance_fields_are_own_properties = own;
                production[0].module_members_on_prototype = prototype;
                globals::emit_read(&mut production[0], "input", 37);
                let mut selected = production.clone();
                strings::emit_scalar_length(&mut production, 0, 37);
                candidate::emit(&mut selected, 0, 37);
                for chunks in [&mut production, &mut selected] {
                    chunks[0].emit_op(Op::RETURN, 37);
                    finalize(chunks);
                }
                assert_eq!(production.len(), selected.len());
                for (actual, expected) in production.iter().zip(&selected) {
                    assert_eq!(actual.code, expected.code);
                    assert_eq!(actual.lines, expected.lines);
                    assert_eq!(actual.imports, expected.imports);
                    assert_eq!(actual.constants, expected.constants);
                    assert_eq!(actual.global_imports, expected.global_imports);
                    assert_eq!(actual.globals, expected.globals);
                    assert_eq!(actual.local_count, expected.local_count);
                    assert_eq!(actual.scratch_high_water, expected.scratch_high_water);
                    assert_eq!(actual.block_i32_results, expected.block_i32_results);
                    assert_eq!(actual.call_indirect_sigs, expected.call_indirect_sigs);
                    assert_eq!(actual.call_indirect_canon, expected.call_indirect_canon);
                    assert_eq!(actual.stack_switch_handlers, expected.stack_switch_handlers);
                    assert_eq!(actual.user_code_offset, expected.user_code_offset);
                    assert_eq!(
                        actual.module_instance_fields_are_own_properties,
                        expected.module_instance_fields_are_own_properties
                    );
                    assert_eq!(
                        actual.module_members_on_prototype,
                        expected.module_members_on_prototype
                    );
                }
                validate(&production);
                validate(&selected);
                assert_eq!(
                    vybe_platform_wasm::write_wasm(&production),
                    vybe_platform_wasm::write_wasm(&selected)
                );
                let options = vybe_platform_wasm::WasmWriteOptions {
                    include_vybe_metadata: true,
                    include_name_section: true,
                };
                assert_eq!(
                    vybe_platform_wasm::write_wasm_with_options(&production, options),
                    vybe_platform_wasm::write_wasm_with_options(&selected, options)
                );
            }
        }
    }
}
