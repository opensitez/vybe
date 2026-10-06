use vybe_compiler::primitives::{bundle, globals, multivalue, platforms};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/multivalue_optimization_stack.rs"]
mod baseline;
fn program(old: bool, count: u16, calls: usize) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<multi-packet>")];
    for call in 0..calls {
        for index in 0..count {
            if index == 0 {
                globals::emit_read(&mut chunks[0], "first", 0);
            } else {
                chunks[0].emit_i32_const(index as i32, 0);
            }
        }
        if old {
            baseline::emit_from_stack(&mut chunks, 0, count, 0);
        } else {
            multivalue::emit_from_stack(&mut chunks, 0, count, 0);
        }
        if call + 1 != calls {
            chunks[0].emit_op(Op::DROP, 0);
        }
    }
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn runtime(first: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("first", first);
    vm
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}
#[test]
fn packet_order_identity_and_tag_are_preserved() {
    let input = Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
        Value::I32(7),
    ])));
    for count in [0, 1, 2, 32] {
        for old in [true, false] {
            let chunks = program(old, count, 1);
            validate(&chunks);
            let result = runtime(input.clone()).run(chunks).unwrap();
            let Value::Object(packet) = result else {
                panic!("expected packet")
            };
            let packet = packet.lock().unwrap();
            let ObjectKind::Array(values) = &packet.kind else {
                panic!("expected array")
            };
            assert_eq!(values.len(), count as usize);
            assert_eq!(
                packet.properties.get(multivalue::MULTI_VALUE_TAG),
                Some(&Value::Bool(true))
            );
            if count > 0 {
                let (Value::Object(actual), Value::Object(expected)) = (&values[0], &input) else {
                    unreachable!()
                };
                assert!(std::sync::Arc::ptr_eq(actual, expected));
            }
            for index in 1..count {
                assert_eq!(values[index as usize], Value::I32(index as i32));
            }
        }
    }
}
#[test]
fn scalar_packet_elements_keep_exact_values() {
    for input in [
        Value::Null,
        Value::Undefined,
        Value::Bool(false),
        Value::I32(-1),
        Value::F64(-0.0),
        Value::F64(f64::from_bits(0x7ff8000000000012)),
        Value::String("é😀\0".into()),
    ] {
        for old in [true, false] {
            let Value::Object(packet) = runtime(input.clone()).run(program(old, 1, 1)).unwrap()
            else {
                unreachable!()
            };
            let packet = packet.lock().unwrap();
            let ObjectKind::Array(values) = &packet.kind else {
                unreachable!()
            };
            if let (Value::F64(value), Value::F64(original)) = (&values[0], &input) {
                assert_eq!(value.to_bits(), original.to_bits());
            } else {
                assert_eq!(&values[0], &input);
            }
        }
    }
}
#[test]
fn repeated_packets_avoid_per_element_spills_and_host_pushes() {
    let old = program(true, 32, 20);
    let new = program(false, 32, 20);
    validate(&old);
    validate(&new);
    assert!(new[0].local_count < old[0].local_count);
    assert!(new[0].code.len() < old[0].code.len());
    assert!(
        !new[0]
            .imports
            .iter()
            .any(|i| i.module == "ecma:array" && i.name == "push")
    );
    println!(
        "packets20x32 locals {} -> {} bytecode {} -> {} wasm {} -> {}",
        old[0].local_count,
        new[0].local_count,
        old[0].code.len(),
        new[0].code.len(),
        vybe_platform_wasm::write_wasm(&old).len(),
        vybe_platform_wasm::write_wasm(&new).len()
    );
}
#[test]
#[ignore = "interleaved opt-level-0 packet assembly benchmark"]
fn benchmark_stack_packet_assembly() {
    use std::time::Instant;
    for count in [32, 1024, 8192] {
        let before = program(true, count, 1);
        let after = program(false, count, 1);
        validate(&before);
        validate(&after);
        let mut times = [vec![], vec![]];
        for sample in 0..9 {
            for which in if sample % 2 == 0 { [0, 1] } else { [1, 0] } {
                let mut vm = runtime(Value::I32(0));
                let chunks = if which == 0 {
                    before.clone()
                } else {
                    after.clone()
                };
                let start = Instant::now();
                let result = vm.run(chunks).unwrap();
                times[which].push(start.elapsed().as_micros());
                let Value::Object(packet) = result else {
                    unreachable!()
                };
                let packet = packet.lock().unwrap();
                let ObjectKind::Array(values) = &packet.kind else {
                    unreachable!()
                };
                assert_eq!(values.len(), count as usize);
            }
        }
        for time in &mut times {
            time.sort_unstable();
        }
        println!(
            "packet {count} elements median {} -> {} us {:.2}x",
            times[0][4],
            times[1][4],
            times[0][4] as f64 / times[1][4] as f64
        );
    }
}

#[test]
fn packet_scalar_conversion_and_first_value_branches_validate() {
    for first in [false, true] {
        for old in [true, false] {
            let mut chunks = vec![Chunk::new("<packet-branch>")];
            globals::emit_read(&mut chunks[0], "input", 0);
            match (old, first) {
                (true, true) => baseline::emit_first(&mut chunks, 0, 0),
                (false, true) => multivalue::emit_first(&mut chunks, 0, 0),
                (true, false) => baseline::emit_as_multi_value(&mut chunks, 0, 0),
                (false, false) => multivalue::emit_as_multi_value(&mut chunks, 0, 0),
            }
            chunks[0].emit_op(Op::RETURN, 0);
            bundle::finalize_with_runtime_helpers(&mut chunks);
            globals::declare_free_globals(&mut chunks);
            globals::normalize_global_table(&mut chunks);
            let checked = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&vybe_platform_wasm::write_wasm(&chunks));
            if old {
                assert!(
                    checked.is_err(),
                    "captured value branches have void results"
                );
            } else {
                checked.unwrap();
            }
            let mut tagged = Object::new_array(vec![Value::I32(19), Value::String("tail".into())]);
            tagged
                .properties
                .insert(multivalue::MULTI_VALUE_TAG.into(), Value::Bool(true));
            let packet = Value::Object(vybe_runtime::heap::alloc(tagged));
            for input in [Value::I32(7), Value::String("é😀".into()), packet.clone()] {
                let mut vm = runtime(Value::Null);
                vm.set_global("input", input.clone());
                let result = vm.run(chunks.clone()).unwrap();
                if first {
                    assert_eq!(
                        result,
                        if matches!(input, Value::Object(_)) {
                            Value::I32(19)
                        } else {
                            input
                        }
                    );
                } else {
                    let Value::Object(out) = result else {
                        panic!("expected packet")
                    };
                    if let Value::Object(original) = input {
                        assert!(std::sync::Arc::ptr_eq(&out, &original));
                    } else {
                        let out = out.lock().unwrap();
                        let ObjectKind::Array(items) = &out.kind else {
                            unreachable!()
                        };
                        assert_eq!(items, &vec![input]);
                        assert_eq!(
                            out.properties.get(multivalue::MULTI_VALUE_TAG),
                            Some(&Value::Bool(true))
                        );
                    }
                }
            }
        }
    }
}
