//! Actual checked numeric packing leaf and frozen generic packing comparison.
use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::{bundle, globals, packing, platforms};
use vybe_runtime::value::{Function, Object, ObjectKind};
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/collection_optimization_packing_number_store.rs"]
mod baseline;
#[path = "fixtures/collection_optimization_bitconverter_caller.rs"]
mod frozen_caller;
fn object(o: Object) -> Value {
    Value::Object(vybe_runtime::heap::alloc(o))
}
fn array() -> Value {
    object(Object::new_array(vec![Value::I32(0); 8]))
}
fn values(input: &Value) -> Vec<Value> {
    let Value::Object(o) = input else {
        unreachable!()
    };
    let o = o.lock().unwrap();
    let ObjectKind::Array(v) = &o.kind else {
        unreachable!()
    };
    v.clone()
}
fn vm(input: Value, value: Value, lo: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    for (k, v) in [("input", input), ("value", value), ("lo", lo)] {
        vm.set_global(k, v);
    }
    vm
}
fn error(mut e: vybe_runtime::VMError) -> String {
    for frame in &mut e.call_stack {
        frame.offset = 0;
    }
    format!("{e:?}")
}
fn program(mode: u8, width: u8, little: bool, calls: usize) -> Vec<Chunk> {
    let mut c = Chunk::new("<packing-cache>");
    c.local_count = 3;
    for _ in 0..calls {
        for (name, slot) in [("input", 0), ("value", 1), ("lo", 2)] {
            globals::emit_read(&mut c, name, 0);
            c.emit_op_u16(Op::LOCAL_SET, slot, 0);
        }
        let mut chunks = vec![c];
        if mode == 0 {
            let order = if little {
                baseline::Endian::Little
            } else {
                baseline::Endian::Big
            };
            match width {
                2 => baseline::emit_store_u16_to_array_from_number_slot(
                    &mut chunks,
                    0,
                    0,
                    1,
                    order,
                    0,
                ),
                4 => baseline::emit_store_u32_to_array_from_number_slot(
                    &mut chunks,
                    0,
                    0,
                    1,
                    order,
                    0,
                ),
                _ => baseline::emit_store_u64_parts_to_array_from_number_slots(
                    &mut chunks,
                    0,
                    0,
                    1,
                    2,
                    order,
                    0,
                ),
            }
        } else if mode == 1 {
            let order = if little {
                packing::Endian::Little
            } else {
                packing::Endian::Big
            };
            match width {
                2 => packing::emit_store_u16_to_array_from_number_slot(
                    &mut chunks,
                    0,
                    0,
                    1,
                    order,
                    0,
                ),
                4 => packing::emit_store_u32_to_array_from_number_slot(
                    &mut chunks,
                    0,
                    0,
                    1,
                    order,
                    0,
                ),
                _ => packing::emit_store_u64_parts_to_array_from_number_slots(
                    &mut chunks,
                    0,
                    0,
                    1,
                    2,
                    order,
                    0,
                ),
            }
        } else {
            assert_eq!(width, 4);
            packing::emit_store_u32_to_array_from_f64_slot(
                &mut chunks,
                0,
                0,
                1,
                if little {
                    packing::Endian::Little
                } else {
                    packing::Endian::Big
                },
                0,
            );
        }
        c = chunks.remove(0);
    }
    globals::emit_read(&mut c, "input", 0);
    c.emit_op(Op::RETURN, 0);
    let mut callback = Chunk::new("<actual-valueof>");
    callback.emit_string_const("coerce", 0);
    callback.emit_i32_const(0, 0);
    let call = callback.add_import("ecma:string", "charAt");
    callback.emit_call(call, 2, 0);
    callback.emit_op(Op::RETURN, 0);
    let mut chunks = vec![c, callback];
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
// Keep failure diagnostics at the exact exported function/operator boundary.
fn native_diagnostic(chunks: &[Chunk], bytes: &[u8], error: &wasmparser::BinaryReaderError) {
    let mut names = std::collections::BTreeMap::new();
    let named_bytes = vybe_platform_wasm::write_wasm_with_options(
        chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_name_section: true,
            ..Default::default()
        },
    );
    for payload in wasmparser::Parser::new(0).parse_all(&named_bytes) {
        if let wasmparser::Payload::CustomSection(section) = payload.unwrap() {
            if let wasmparser::KnownCustom::Name(reader) = section.as_known() {
                for name in reader {
                    if let wasmparser::Name::Function(map) = name.unwrap() {
                        for entry in map {
                            let entry = entry.unwrap();
                            names.insert(entry.index, entry.name.to_string());
                        }
                    }
                }
            }
        }
    }
    let mut function = 0;
    for payload in wasmparser::Parser::new(0).parse_all(bytes) {
        if let wasmparser::Payload::ImportSection(section) = payload.unwrap() {
            for import in section.into_imports() {
                if matches!(
                    import.unwrap().ty,
                    wasmparser::TypeRef::Func(_) | wasmparser::TypeRef::FuncExact(_)
                ) {
                    function += 1;
                }
            }
        }
    }
    for payload in wasmparser::Parser::new(0).parse_all(bytes) {
        if let wasmparser::Payload::CodeSectionEntry(body) = payload.unwrap() {
            if body.range().contains(&error.offset()) {
                eprintln!(
                    "native failure function={function} name={:?} range={:?} error={error}",
                    names.get(&function),
                    body.range()
                );
                let mut operators = body.get_operators_reader().unwrap();
                while !operators.eof() {
                    let (operator, offset) = operators.read_with_offset().unwrap();
                    if offset + 100 >= error.offset() && offset <= error.offset() + 30 {
                        eprintln!("  wasm@{offset}: {operator:?}");
                        if let wasmparser::Operator::Call { function_index } = operator {
                            eprintln!("    target={:?}", names.get(&function_index));
                        }
                    }
                }
            }
            function += 1;
        }
    }
}
fn validate(chunks: &[Chunk]) {
    let bytes = vybe_platform_wasm::write_wasm(chunks);
    if let Err(error) = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
    {
        native_diagnostic(chunks, &bytes, &error);
        panic!("packing native validation: {error}");
    }
}
fn coercible(id: &str, value: f64) -> Value {
    let mut o = Object::new();
    o.properties.insert("id".into(), Value::String(id.into()));
    o.properties.insert("number".into(), Value::F64(value));
    o.properties.insert(
        "valueOf".into(),
        object(Object {
            kind: ObjectKind::Function(Function {
                name: Some("valueOf".into()),
                arity: 0,
                chunk_index: 1,
                upvalues: vec![],
            }),
            ..Object::new()
        }),
    );
    object(o)
}
#[test]
fn checked_numeric_leaf_and_generic_store_keep_bytes_endian_traps_and_no_partial_store() {
    for width in [2, 4, 8] {
        for little in [false, true] {
            let before = program(0, width, little, 1);
            for mode in if width == 4 { vec![1, 2] } else { vec![1] } {
                let after = program(mode, width, little, 1);
                validate(&before);
                validate(&after);
                for number in [
                    0.0,
                    1.5,
                    255.0,
                    65535.0,
                    2147483648.0,
                    4294967295.0,
                    -1.0,
                    f64::NAN,
                    f64::INFINITY,
                ] {
                    let outcomes = [&before, &after].map(|chunks| {
                        let input = array();
                        let result = vm(input.clone(), Value::F64(number), Value::F64(4660.0))
                            .run(chunks.clone())
                            .map(|_| ())
                            .map_err(error);
                        (result, values(&input))
                    });
                    assert_eq!(
                        outcomes[0], outcomes[1],
                        "width={width} mode={mode} little={little} number={number}"
                    );
                    if outcomes[0].0.is_ok() {
                        let bits = number as u32;
                        let mut bytes = if width == 8 {
                            [bits.to_be_bytes(), 4660u32.to_be_bytes()].concat()
                        } else {
                            bits.to_be_bytes()[4 - width as usize..].to_vec()
                        };
                        if little {
                            bytes.reverse();
                        }
                        let expected: Vec<Value> = bytes
                            .into_iter()
                            .map(|b| Value::I32(b as i32))
                            .chain(std::iter::repeat_n(Value::I32(0), 8 - width as usize))
                            .collect();
                        assert_eq!(outcomes[0].1, expected);
                    } else {
                        assert_eq!(outcomes[0].1, vec![Value::I32(0); 8]);
                    }
                }
            }
        }
    }
}
#[test]
fn mixed_object_coercion_order_other_operand_mutation_and_throw_stay_exact() {
    for width in [2, 4, 8] {
        for throwing in [false, true] {
            for mixed in [false, true] {
                let programs = [program(0, width, false, 1), program(1, width, false, 1)];
                let outcomes = programs.each_ref().map(|chunks| {
                    let input = array();
                    let hi = coercible("hi", 4660.0);
                    let lo = coercible("lo", 22136.0);
                    let observed_lo = lo.clone();
                    let value = if mixed { Value::F64(4660.0) } else { hi };
                    let mut runtime = vm(input.clone(), value, lo);
                    let calls = Arc::new(Mutex::new(Vec::new()));
                    let observed = calls.clone();
                    let observed_input = input.clone();
                    runtime.register_host_fn(
                        "ecma:string",
                        "charAt",
                        Box::new(move |ctx, args| {
                            assert_eq!(args[0], Value::String("coerce".into()));
                            let Value::Object(receiver) = ctx.current_js_this() else {
                                panic!("actual valueOf receiver")
                            };
                            let receiver = receiver.lock().unwrap();
                            let id = receiver.properties["id"].as_str().to_string();
                            let number = receiver.properties["number"].as_f64();
                            drop(receiver);
                            let mut trace = observed.lock().unwrap();
                            trace.push((id.clone(), values(&observed_input)));
                            if id == "hi" {
                                let Value::Object(lo) = &observed_lo else {
                                    unreachable!()
                                };
                                lo.lock()
                                    .unwrap()
                                    .properties
                                    .insert("number".into(), Value::F64(99.0));
                            }
                            if throwing && trace.len() == 2 {
                                ctx.throw_value(Value::String("coerce-error".into()));
                                Value::Undefined
                            } else {
                                Value::F64(number)
                            }
                        }),
                    );
                    let result = runtime.run(chunks.clone()).map(|_| ()).map_err(error);
                    let trace = calls.lock().unwrap().clone();
                    (result, values(&input), trace)
                });
                assert_eq!(
                    outcomes[0], outcomes[1],
                    "width={width} throwing={throwing} mixed={mixed}"
                );
                let expected = if mixed && width < 8 {
                    0
                } else if throwing {
                    2
                } else if mixed {
                    4
                } else {
                    width as usize
                };
                assert_eq!(outcomes[0].2.len(), expected);
                if throwing && expected > 0 {
                    assert!(outcomes[0].0.is_err());
                }
                for (_, state) in &outcomes[0].2 {
                    assert_eq!(state, &vec![Value::I32(0); 8]);
                }
            }
        }
    }
}
#[test]
fn numeric_leaf_reduces_emission_and_generic_body_is_exactly_preserved() {
    assert!(
        include_str!("../src/primitives/packing.rs").starts_with(include_str!(
            "fixtures/collection_optimization_packing_number_store_source.txt"
        ))
    );
    let old = program(0, 4, false, 16);
    let new = program(2, 4, false, 16);
    validate(&old);
    validate(&new);
    let bytes = |chunks: &[Chunk]| chunks.iter().map(|c| c.code.len()).sum::<usize>();
    assert!(bytes(&new) < bytes(&old));
    eprintln!(
        "packing-numeric bytecode={}/{} wasm={}/{}",
        bytes(&old),
        bytes(&new),
        vybe_platform_wasm::write_wasm(&old).len(),
        vybe_platform_wasm::write_wasm(&new).len()
    );
}
#[test]
#[ignore = "interleaved opt-level-0 checked numeric packing diagnostic"]
fn benchmark_checked_numeric_u32_store() {
    let programs = [program(0, 4, false, 256), program(2, 4, false, 256)];
    let mut times = [vec![], vec![]];
    for sample in 0..9 {
        for which in [sample % 2, 1 - sample % 2] {
            let input = array();
            let mut runtime = vm(input, Value::F64(4660.0), Value::Null);
            let chunks = programs[which].clone();
            let start = std::time::Instant::now();
            runtime.run(chunks).unwrap();
            times[which].push(start.elapsed().as_micros());
        }
    }
    for time in &mut times {
        time.sort_unstable();
    }
    eprintln!(
        "packing-numeric before={}us after={}us ratio={:.2}x",
        times[0][4],
        times[1][4],
        times[0][4] as f64 / times[1][4] as f64
    );
}

#[test]
fn parsed_csharp_bitconverter_preserves_all_dispatch_results_and_checked_numeric_leaf() {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_csharp::register);
    for (label, source, expected) in [
        (
            "boolean-true",
            "using System; bool value = true; byte[] answer = BitConverter.GetBytes(value);",
            vec![Value::I32(1)],
        ),
        (
            "boolean-false",
            "using System; bool value = false; byte[] answer = BitConverter.GetBytes(value);",
            vec![Value::I32(0)],
        ),
        (
            "string",
            "using System; string value = \"A\"; byte[] answer = BitConverter.GetBytes(value);",
            vec![Value::I32(65), Value::I32(0)],
        ),
        (
            "positive",
            "using System; int value = 4660; byte[] answer = BitConverter.GetBytes(value);",
            vec![Value::I32(52), Value::I32(18), Value::I32(0), Value::I32(0)],
        ),
        (
            "negative",
            "using System; int value = -2; byte[] answer = BitConverter.GetBytes(value);",
            vec![
                Value::I32(254),
                Value::I32(255),
                Value::I32(255),
                Value::I32(255),
            ],
        ),
        (
            "eight-byte",
            "using System; double value = 1.25; byte[] answer = BitConverter.GetBytes(value);",
            std::iter::once(Value::F64(1.25))
                .chain(std::iter::repeat_n(Value::I32(0), 7))
                .collect(),
        ),
    ] {
        let mut module = vybe_language_csharp::parse(source).unwrap();
        module
            .body
            .push(vybe_ast::Statement::new(vybe_ast::StmtKind::Return(Some(
                vybe_ast::Expression::ident("answer"),
            ))));
        let profile =
            vybe_compiler::profile::parse_profile(vybe_language_csharp::profile_source()).unwrap();
        let chunks = vybe_compiler::Compiler::with_profile(profile)
            .compile(&module)
            .unwrap();
        let wasm = vybe_platform_wasm::write_wasm(&chunks);
        if let Err(error) =
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&wasm)
        {
            native_diagnostic(&chunks, &wasm, &error);
            std::fs::write(
                format!("/private/tmp/vybe-packing-csharp-{label}.wasm"),
                wasm,
            )
            .unwrap();
            panic!("actual .NET Source label={label}: {error}");
        }
        let result = vm(Value::Null, Value::Null, Value::Null)
            .run(chunks)
            .unwrap();
        assert_eq!(values(&result), expected, "case={label}");
    }
}

// Exact pre-leaf caller had void outer branches despite producing an array.
// This is a negative native gate for that existing gap, separate from the
// passing generic packing baseline and the strict actual Source checks above.
#[test]
fn frozen_getbytes_caller_records_existing_void_branch_native_gap_and_vm_results() {
    let mut root = Chunk::new("<frozen-getbytes-caller>");
    globals::emit_read(&mut root, "value", 0);
    let mut chunks = vec![root];
    frozen_caller::emit_get_bytes(&mut chunks, 0, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    let bytes = vybe_platform_wasm::write_wasm(&chunks);
    let failure = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .err()
        .expect("frozen GetBytes caller must retain its known native block-type defect");
    assert!(
        failure
            .message()
            .contains("values remaining on stack at end of block"),
        "unexpected frozen caller validation failure: {failure}"
    );
    native_diagnostic(&chunks, &bytes, &failure);
    for (input, expected) in [
        (Value::Bool(true), vec![Value::I32(1)]),
        (Value::Bool(false), vec![Value::I32(0)]),
        (
            Value::String("A".into()),
            vec![Value::I32(65), Value::I32(0)],
        ),
        (
            Value::F64(4660.0),
            vec![Value::I32(52), Value::I32(18), Value::I32(0), Value::I32(0)],
        ),
        (
            Value::F64(-2.0),
            vec![
                Value::I32(254),
                Value::I32(255),
                Value::I32(255),
                Value::I32(255),
            ],
        ),
        (
            Value::F64(1.25),
            std::iter::once(Value::F64(1.25))
                .chain(std::iter::repeat_n(Value::I32(0), 7))
                .collect(),
        ),
    ] {
        let result = vm(Value::Null, input.clone(), Value::Null)
            .run(chunks.clone())
            .unwrap();
        assert_eq!(values(&result), expected, "frozen caller input={input:?}");
    }
}
