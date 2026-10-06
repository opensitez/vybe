use vybe_ast::ArrayTraversalOrder;
use vybe_compiler::primitives::{array_transforms, bundle, globals, platforms};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};

#[path = "fixtures/dynamic_array_transforms.rs"]
mod baseline;

#[derive(Clone, Copy, Debug)]
enum Transform {
    Pack(bool),
    Unpack,
    Merge,
    Reshape(bool),
}

fn array(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}

fn ints(values: &[i32]) -> Value {
    array(values.iter().copied().map(Value::I32).collect())
}

fn items(value: &Value) -> Option<Vec<Value>> {
    let Value::Object(object) = value else {
        return None;
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else {
        return None;
    };
    Some(values.clone())
}

fn assert_values(actual: &Value, expected: &Value) {
    match (items(actual), items(expected)) {
        (Some(a), Some(b)) => {
            assert_eq!(a.len(), b.len());
            for (a, b) in a.iter().zip(&b) {
                assert_values(a, b);
            }
        }
        (None, None) => assert_eq!(actual, expected),
        _ => panic!("array/scalar mismatch: {actual:?} / {expected:?}"),
    }
}

fn program(
    old: bool,
    transform: Transform,
    order: ArrayTraversalOrder,
    calls: usize,
    receiver: bool,
) -> Vec<Chunk> {
    program_with_facts(old, transform, order, calls, receiver, false, false)
}

fn program_with_facts(
    old: bool,
    transform: Transform,
    order: ArrayTraversalOrder,
    calls: usize,
    receiver: bool,
    own_fields: bool,
    prototype_members: bool,
) -> Vec<Chunk> {
    let mut chunk = Chunk::new("<array-transform>");
    chunk.module_instance_fields_are_own_properties = own_fields;
    chunk.module_members_on_prototype = prototype_members;
    if receiver {
        chunk.module_receiver_abi = vybe_runtime::chunk::ReceiverAbi::Parameter;
    }
    let mut chunks = vec![chunk];
    for i in 0..calls {
        let count = match transform {
            Transform::Pack(false) | Transform::Reshape(false) => 2,
            _ => 3,
        };
        for key in ["first", "second", "third"].into_iter().take(count) {
            globals::emit_read(&mut chunks[0], key, 0);
        }
        match (old, transform) {
            (true, Transform::Pack(v)) => baseline::emit_pack_mask(&mut chunks, 0, v, order, 0),
            (false, Transform::Pack(v)) => {
                array_transforms::emit_pack_mask(&mut chunks, 0, v, order, 0)
            }
            (true, Transform::Unpack) => baseline::emit_unpack_mask(&mut chunks, 0, order, 0),
            (false, Transform::Unpack) => {
                array_transforms::emit_unpack_mask(&mut chunks, 0, order, 0)
            }
            (true, Transform::Merge) => baseline::emit_merge_mask(&mut chunks, 0, order, 0),
            (false, Transform::Merge) => {
                array_transforms::emit_merge_mask(&mut chunks, 0, order, 0)
            }
            (true, Transform::Reshape(p)) => baseline::emit_reshape(&mut chunks, 0, p, order, 0),
            (false, Transform::Reshape(p)) => {
                array_transforms::emit_reshape(&mut chunks, 0, p, order, 0)
            }
        }
        if i + 1 != calls {
            chunks[0].emit_op(Op::DROP, 0);
        }
    }
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}

fn vm(inputs: &[Value; 3]) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    for (key, value) in ["first", "second", "third"].into_iter().zip(inputs) {
        vm.set_global(key, value.clone());
    }
    vm
}

#[test]
fn transforms_preserve_masks_broadcasting_padding_and_orders() {
    let mask = array(vec![
        Value::Bool(true),
        Value::Bool(false),
        Value::Bool(true),
        Value::Bool(false),
    ]);
    let ranked_mask = array(vec![
        array(vec![Value::Bool(true), Value::Bool(false)]),
        array(vec![Value::Bool(false), Value::Bool(true)]),
    ]);
    let cases = vec![
        (
            Transform::Pack(false),
            [ints(&[1, 2, 3, 4]), mask.clone(), Value::Null],
            ints(&[1, 3]),
        ),
        (
            Transform::Pack(true),
            [ints(&[1, 2, 3, 4]), mask.clone(), ints(&[9, 8, 7, 6])],
            ints(&[1, 3, 7, 6]),
        ),
        (
            Transform::Pack(false),
            [ints(&[]), ints(&[]), Value::Null],
            ints(&[]),
        ),
        (
            Transform::Pack(false),
            [
                ints(&[1, 2, 3, 4]),
                array(vec![
                    Value::I32(1),
                    Value::Null,
                    Value::String("yes".into()),
                    Value::String("".into()),
                ]),
                Value::Null,
            ],
            ints(&[1, 3]),
        ),
        (
            Transform::Unpack,
            [ints(&[10, 20]), mask.clone(), Value::I32(-1)],
            ints(&[10, -1, 20, -1]),
        ),
        (
            Transform::Unpack,
            [ints(&[10, 20]), mask, ints(&[1, 2, 3, 4])],
            ints(&[10, 2, 20, 4]),
        ),
        (
            Transform::Merge,
            [Value::I32(8), ints(&[1, 2, 3, 4]), ranked_mask],
            array(vec![ints(&[8, 2]), ints(&[3, 8])]),
        ),
        (
            Transform::Merge,
            [ints(&[1, 2]), ints(&[3, 4]), Value::Bool(true)],
            ints(&[1, 2]),
        ),
        (
            Transform::Merge,
            [ints(&[1, 2]), ints(&[3, 4]), Value::I32(0)],
            ints(&[3, 4]),
        ),
        (
            Transform::Reshape(true),
            [ints(&[1, 2]), ints(&[5]), ints(&[8, 9])],
            ints(&[1, 2, 8, 9, 8]),
        ),
    ];
    for order in [
        ArrayTraversalOrder::RowMajor,
        ArrayTraversalOrder::ColumnMajor,
    ] {
        for receiver in [false, true] {
            for (transform, input, expected) in &cases {
                for old in [true, false] {
                    let chunks = program(old, *transform, order, 1, receiver);
                    let result = vm(input).run(chunks).unwrap();
                    assert_values(&result, expected);
                }
            }
            for order_case in [order] {
                let input = [ints(&[1, 2, 3, 4, 5, 6]), ints(&[2, 3]), Value::Null];
                let expected = match order_case {
                    ArrayTraversalOrder::RowMajor => {
                        array(vec![ints(&[1, 2, 3]), ints(&[4, 5, 6])])
                    }
                    _ => array(vec![ints(&[1, 3, 5]), ints(&[2, 4, 6])]),
                };
                for old in [true, false] {
                    assert_values(
                        &vm(&input)
                            .run(program(
                                old,
                                Transform::Reshape(false),
                                order_case,
                                1,
                                receiver,
                            ))
                            .unwrap(),
                        &expected,
                    );
                }
                assert_values(&input[0], &ints(&[1, 2, 3, 4, 5, 6]));
            }
        }
    }
}

#[test]
fn transforms_reduce_emission_and_validate_complete_exports() {
    for transform in [
        Transform::Pack(false),
        Transform::Pack(true),
        Transform::Unpack,
        Transform::Merge,
        Transform::Reshape(false),
        Transform::Reshape(true),
    ] {
        for order in [
            ArrayTraversalOrder::RowMajor,
            ArrayTraversalOrder::ColumnMajor,
        ] {
            let programs = [
                program(true, transform, order, 20, false),
                program(false, transform, order, 20, false),
            ];
            let sizes = programs
                .each_ref()
                .map(|p| p.iter().map(|c| c.code.len()).sum::<usize>());
            let wasm = programs
                .each_ref()
                .map(|p| vybe_platform_wasm::write_wasm(p));
            println!(
                "{transform:?} {order:?}: bytecode={}/{} wasm={}/{}",
                sizes[0],
                sizes[1],
                wasm[0].len(),
                wasm[1].len()
            );
            assert!(sizes[1] < sizes[0]);
            for module in wasm {
                wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                    .validate_all(&module)
                    .unwrap();
            }
        }
    }
}

#[test]
#[ignore = "interleaved debug execution benchmark"]
fn transform_execution_benchmark() {
    for size in [256, 8192] {
        for transform in [Transform::Pack(false), Transform::Unpack, Transform::Merge] {
            let data = array((0..size).map(Value::I32).collect());
            let mask = array((0..size).map(|i| Value::Bool(i % 2 == 0)).collect());
            let input = match transform {
                Transform::Pack(_) | Transform::Unpack => [data, mask, Value::I32(-1)],
                _ => [data, Value::I32(-1), mask],
            };
            let programs = [
                program(true, transform, ArrayTraversalOrder::RowMajor, 1, false),
                program(false, transform, ArrayTraversalOrder::RowMajor, 1, false),
            ];
            let mut timings = [Vec::new(), Vec::new()];
            let mut results = [Value::Null, Value::Null];
            for run in 0..15 {
                for which in [run % 2, 1 - run % 2] {
                    let mut vm = vm(&input);
                    let chunks = programs[which].clone();
                    let start = std::time::Instant::now();
                    results[which] = vm.run(chunks).unwrap();
                    timings[which].push(start.elapsed());
                }
            }
            assert_values(&results[0], &results[1]);
            for t in &mut timings {
                t.sort();
            }
            println!(
                "{transform:?} size={size}: before={}us after={}us speedup={:.2}x",
                timings[0][7].as_micros(),
                timings[1][7].as_micros(),
                timings[0][7].as_secs_f64() / timings[1][7].as_secs_f64()
            );
        }
    }
}

#[test]
fn shared_transform_bodies_link_once_and_preserve_module_facts() {
    for transform in [
        Transform::Pack(false),
        Transform::Pack(true),
        Transform::Unpack,
        Transform::Merge,
        Transform::Reshape(false),
        Transform::Reshape(true),
    ] {
        for order in [
            ArrayTraversalOrder::RowMajor,
            ArrayTraversalOrder::ColumnMajor,
        ] {
            for (own_fields, prototype_members) in
                [(false, false), (true, false), (false, true), (true, true)]
            {
                let mut chunks = program_with_facts(
                    false,
                    transform,
                    order,
                    20,
                    true,
                    own_fields,
                    prototype_members,
                );
                let helpers: Vec<_> = chunks
                    .iter()
                    .filter(|c| c.name.starts_with("__stdlib_array_"))
                    .collect();
                assert_eq!(helpers.len(), 1, "{transform:?} {order:?}");
                assert_eq!(
                    helpers[0].module_instance_fields_are_own_properties,
                    own_fields
                );
                assert_eq!(helpers[0].module_members_on_prototype, prototype_members);
                let before = vybe_platform_wasm::write_wasm(&chunks);
                wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                    .validate_all(&before)
                    .unwrap();
                bundle::finalize_with_runtime_helpers(&mut chunks);
                assert_eq!(before, vybe_platform_wasm::write_wasm(&chunks));
            }
        }
    }
    for own_fields in [false, true] {
        let mut runtime = vm(&[Value::Null, Value::Null, Value::Null]);
        let mut ty = vybe_runtime::typedef::TypeDef::new("Shape");
        ty.add_field("0");
        let id = runtime.type_registry.register(ty);
        let mut shape = Object::new_typed_with_fields(id, 1);
        shape.fields[0] = Value::I32(2);
        shape.properties.insert("length".into(), Value::I32(1));
        runtime.set_global("first", ints(&[1, 2]));
        runtime.set_global("second", Value::Object(vybe_runtime::heap::alloc(shape)));
        let chunks = program_with_facts(
            false,
            Transform::Reshape(false),
            ArrayTraversalOrder::RowMajor,
            1,
            false,
            own_fields,
            false,
        );
        assert_values(
            &runtime.run(chunks).unwrap(),
            &if own_fields { ints(&[]) } else { ints(&[1, 2]) },
        );
    }
}
