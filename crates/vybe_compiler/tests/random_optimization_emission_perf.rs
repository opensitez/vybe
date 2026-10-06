use std::sync::{
    atomic::{AtomicUsize, Ordering},
    Arc,
};
use vybe_compiler::primitives::{bundle, globals, platforms, random};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, Value, VM};

#[path = "fixtures/random_optimization_baseline.rs"]
mod baseline;

#[derive(Clone, Copy)]
enum Operation {
    Draws(usize),
    Sample,
    Shuffle,
    SampleK,
}

fn program(old: bool, operation: Operation) -> Vec<Chunk> {
    let mut chunks = unlinked_program(old, operation);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}

fn unlinked_program(old: bool, operation: Operation) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<random-optimization>")];
    match operation {
        Operation::Draws(count) => {
            let result = chunks[0].alloc_scratch(1);
            let new = chunks[0].add_import("ecma:array", "new");
            let push = chunks[0].add_import("ecma:array", "push");
            chunks[0].emit_call(new, 0, 0);
            chunks[0].emit_op_u16(Op::LOCAL_SET, result, 0);
            for _ in 0..count {
                chunks[0].emit_op_u16(Op::LOCAL_GET, result, 0);
                if old {
                    baseline::emit_next_unit(&mut chunks, 0, 0);
                } else {
                    random::emit_next_unit(&mut chunks, 0, 0);
                }
                chunks[0].emit_call(push, 2, 0);
                chunks[0].emit_op(Op::DROP, 0);
            }
            chunks[0].emit_op_u16(Op::LOCAL_GET, result, 0);
        }
        operation => {
            globals::emit_read(&mut chunks[0], "input", 0);
            match (old, operation) {
                (true, Operation::Sample) => baseline::emit_sample(&mut chunks, 0, 1, 0),
                (false, Operation::Sample) => random::emit_sample(&mut chunks, 0, 1, 0),
                (true, Operation::Shuffle) => baseline::emit_shuffle(&mut chunks, 0, 1, 0),
                (false, Operation::Shuffle) => random::emit_shuffle(&mut chunks, 0, 1, 0),
                (old, Operation::SampleK) => {
                    globals::emit_read(&mut chunks[0], "count", 0);
                    if old {
                        baseline::emit_sample_k(&mut chunks, 0, 0);
                    } else {
                        random::emit_sample_k(&mut chunks, 0, 0);
                    }
                }
                _ => unreachable!(),
            }
        }
    }
    chunks[0].emit_op(Op::RETURN, 0);
    chunks
}

fn array(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}
fn items(value: &Value) -> Vec<Value> {
    let Value::Object(object) = value else {
        panic!("expected array");
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else {
        panic!("expected array");
    };
    values.clone()
}
fn runtime(state: Value, input: Value, count: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("__vybe_rng", state);
    vm.set_global("input", input);
    vm.set_global("count", count);
    vm
}
fn validate(chunks: &[Chunk]) {
    let bytes = vybe_platform_wasm::write_wasm(chunks);
    if let Err(error) = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
    {
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
        let mut first_defined = 0;
        for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
            if let wasmparser::Payload::ImportSection(section) = payload.unwrap() {
                for import in section.into_imports() {
                    if matches!(
                        import.unwrap().ty,
                        wasmparser::TypeRef::Func(_) | wasmparser::TypeRef::FuncExact(_)
                    ) {
                        first_defined += 1;
                    }
                }
            }
        }
        let mut function = first_defined;
        for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
            if let wasmparser::Payload::CodeSectionEntry(body) = payload.unwrap() {
                if body.range().contains(&error.offset()) {
                    eprintln!(
                        "native failure function={} name={:?} range={:?} error={error}",
                        function,
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
        panic!("native module rejected: {error}");
    }
}

fn step(mut state: i32) -> i32 {
    state ^= state.wrapping_shl(13);
    state ^= ((state as u32) >> 17) as i32;
    state ^= state.wrapping_shl(5);
    state
}

#[test]
fn shared_entropy_preserves_seeded_stream_and_final_state() {
    for state in [1i32, -1, i32::MIN, 0x12345678] {
        let mut expected_state = state;
        let expected: Vec<_> = (0..32)
            .map(|_| {
                expected_state = step(expected_state);
                Value::F64((expected_state & 0x7fffff) as f64 / 8388608.0)
            })
            .collect();
        for old in [true, false] {
            let chunks = program(old, Operation::Draws(32));
            validate(&chunks);
            let mut vm = runtime(Value::I32(state), Value::Null, Value::Null);
            assert_eq!(items(&vm.run(chunks).unwrap()), expected);
            assert_eq!(
                vm.global("__vybe_rng").expect("RNG state").clone(),
                Value::I32(expected_state)
            );
        }
    }
}

#[test]
fn lazy_entropy_initialization_consumes_one_host_draw() {
    for old in [true, false] {
        let chunks = program(old, Operation::Draws(3));
        let calls = Arc::new(AtomicUsize::new(0));
        let observed = calls.clone();
        let mut vm = runtime(Value::Null, Value::Null, Value::Null);
        // Instrument the existing nondeterministic capability identically for both emitters.
        vm.register_host_fn(
            "ecma:math",
            "random",
            Box::new(move |_, _| {
                observed.fetch_add(1, Ordering::Relaxed);
                Value::F64(0.5)
            }),
        );
        let actual = items(&vm.run(chunks).unwrap());
        let mut state = 536870913i32;
        let expected: Vec<_> = (0..3)
            .map(|_| {
                state = step(state);
                Value::F64((state & 0x7fffff) as f64 / 8388608.0)
            })
            .collect();
        assert_eq!(actual, expected);
        assert_eq!(
            vm.global("__vybe_rng").expect("RNG state").clone(),
            Value::I32(state)
        );
        assert_eq!(calls.load(Ordering::Relaxed), 1);
    }
}

#[test]
fn derived_operations_preserve_entropy_order_identity_and_population() {
    for len in [0usize, 1, 32] {
        for operation in [Operation::Sample, Operation::Shuffle, Operation::SampleK] {
            for count in [
                Value::I32(0),
                Value::I32(len.min(3) as i32),
                Value::F64(len.min(3) as f64),
            ] {
                let mut outputs = vec![];
                let mut states = vec![];
                let original: Vec<_> = (0..len).map(|i| Value::I32(i as i32)).collect();
                for old in [true, false] {
                    let chunks = program(old, operation);
                    validate(&chunks);
                    let input = array(original.clone());
                    let mut vm = runtime(Value::I32(42), input.clone(), count.clone());
                    let result = vm.run(chunks).unwrap();
                    if matches!(operation, Operation::Shuffle) {
                        let (Value::Object(a), Value::Object(b)) = (&input, &result) else {
                            panic!("arrays");
                        };
                        assert!(Arc::ptr_eq(a, b));
                        let mut permutation = items(&result);
                        permutation.sort_by_key(|v| match v {
                            Value::I32(i) => *i,
                            _ => panic!("integer"),
                        });
                        assert_eq!(permutation, original);
                    } else {
                        assert_eq!(items(&input), original);
                    }
                    outputs.push(if matches!(operation, Operation::Sample) {
                        vec![result]
                    } else {
                        items(&result)
                    });
                    states.push(vm.global("__vybe_rng").expect("RNG state").clone());
                }
                assert_eq!(outputs[0], outputs[1]);
                assert_eq!(states[0], states[1]);
                if len == 0 || matches!(operation, Operation::Shuffle) && len == 1 {
                    assert_eq!(states[1], Value::I32(42));
                }
            }
        }
    }
}

#[test]
fn repeated_entropy_draws_share_one_body_and_finalize_idempotently() {
    let old = program(true, Operation::Draws(20));
    let mut new = program(false, Operation::Draws(20));
    validate(&new);
    assert_eq!(
        new.iter()
            .filter(|c| c.name == "__stdlib_random_next_unit")
            .count(),
        1
    );
    let bytes = |chunks: &[Chunk]| chunks.iter().map(|c| c.code.len()).sum::<usize>();
    let before = bytes(&new);
    let wasm_before = vybe_platform_wasm::write_wasm(&new);
    bundle::finalize_with_runtime_helpers(&mut new);
    assert_eq!(bytes(&new), before);
    assert_eq!(vybe_platform_wasm::write_wasm(&new), wasm_before);
    validate(&new);
    assert!(bytes(&new) < bytes(&old));
    println!(
        "random20draws: bytecode={}/{} wasm={}/{}",
        bytes(&old),
        bytes(&new),
        vybe_platform_wasm::write_wasm(&old).len(),
        wasm_before.len()
    );
}

#[test]
#[ignore = "interleaved opt-level-0 shuffle benchmark"]
fn shuffle_execution_benchmark() {
    let programs = [
        program(true, Operation::Shuffle),
        program(false, Operation::Shuffle),
    ];
    let mut timings = [vec![], vec![]];
    let mut results = [vec![], vec![]];
    for round in 0..9 {
        for which in [round % 2, 1 - round % 2] {
            let input = array((0..8192).map(Value::I32).collect());
            let mut vm = runtime(Value::I32(42), input, Value::Null);
            let chunks = programs[which].clone();
            let start = std::time::Instant::now();
            let result = vm.run(chunks).unwrap();
            timings[which].push(start.elapsed());
            results[which] = items(&result);
        }
    }
    assert_eq!(results[0], results[1]);
    for samples in &mut timings {
        samples.sort();
    }
    println!(
        "shuffle8192: before={}us after={}us speedup={:.2}x",
        timings[0][4].as_micros(),
        timings[1][4].as_micros(),
        timings[0][4].as_secs_f64() / timings[1][4].as_secs_f64()
    );
}

fn source_alias_program(callable_kind: u8) -> Vec<Chunk> {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_python::register);
    // __py_random_r is the declared production Python random hook, lowered by
    // the real language adapter. No test dispatcher or replacement is used.
    let draw = match callable_kind {
        0 => "answers = [__py_random_r(), __py_random_r()]\n",
        1 => "def draw():\n    return __py_random_r()\nanswers = [draw(), draw()]\n",
        2 => "draw = lambda: __py_random_r()\nanswers = [draw(), draw()]\n",
        _ => unreachable!(),
    };
    let source = format!("def __stdlib_random_next_unit():\n    return 42.0\n__vybe_random_next_unit = __stdlib_random_next_unit\nalias_answer = __vybe_random_next_unit()\n{draw}");
    let mut module = vybe_language_python::parse(&source).unwrap();
    module
        .body
        .push(vybe_ast::Statement::new(vybe_ast::StmtKind::Return(Some(
            vybe_ast::Expression::ident("answers"),
        ))));
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_python::profile_source()).unwrap();
    vybe_compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap()
}

#[test]
fn source_entropy_alias_preserves_user_function_draws_and_final_state() {
    let mut baseline_vm = runtime(Value::I32(42), Value::Null, Value::Null);
    let expected = items(&baseline_vm.run(program(true, Operation::Draws(2))).unwrap());
    let expected_state = baseline_vm.global("__vybe_rng").unwrap().clone();
    // The same VM loads all three source forms after an unrelated program,
    // exercising the ordinary REF_FUNC relocation at nonzero chunk bases.
    let mut vm = runtime(Value::I32(42), Value::Null, Value::Null);
    vm.run(program(true, Operation::Draws(1))).unwrap();
    for callable_kind in 0..=2 {
        eprintln!("random source callable_kind={callable_kind} (0=root, 1=function, 2=lambda)");
        let mut chunks = source_alias_program(callable_kind);
        validate(&chunks);
        assert!(!chunks[0]
            .global_inits
            .iter()
            .any(|init| init.name == "__vybe_random_next_unit"));
        assert_eq!(
            chunks
                .iter()
                .filter(|chunk| chunk.name == "__stdlib_random_next_unit")
                .count(),
            2
        );
        let bytes: Vec<_> = chunks.iter().map(|chunk| chunk.code.clone()).collect();
        let count = chunks.len();
        bundle::finalize_with_runtime_helpers(&mut chunks);
        assert_eq!(chunks.len(), count);
        assert_eq!(
            chunks
                .iter()
                .map(|chunk| chunk.code.clone())
                .collect::<Vec<_>>(),
            bytes
        );
        vm.set_global("__vybe_rng", Value::I32(42));
        assert_eq!(items(&vm.run(chunks).unwrap()), expected);
        assert_eq!(vm.global("__vybe_rng"), Some(&expected_state));
        assert_eq!(vm.global("alias_answer"), Some(&Value::F64(42.0)));
    }
}

#[test]
fn zero_argument_private_entropy_reserves_only_one_valid_marker_slot() {
    let mut chunks = vec![Chunk::new("<zero-local-random>")];
    for draw in 0..10 {
        random::emit_next_unit(&mut chunks, 0, 1);
        if draw != 9 {
            chunks[0].emit_op(Op::DROP, 1);
        }
    }
    assert_eq!(chunks[0].local_count, 1);
    assert!(chunks[0]
        .local_names
        .iter()
        .all(|local| local.slot < chunks[0].local_count));
    chunks[0].emit_op(Op::RETURN, 1);
    bundle::finalize_with_runtime_helpers_excluding(&mut chunks, &["__stdlib_random_next_unit"]);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    validate(&chunks);
    assert_eq!(
        chunks
            .iter()
            .filter(|chunk| chunk.name == "__stdlib_random_next_unit")
            .count(),
        1
    );
    let mut vm = runtime(Value::I32(42), Value::Null, Value::Null);
    let value = vm.run(chunks).unwrap();
    let mut expected_state = 42;
    for _ in 0..10 {
        expected_state = step(expected_state);
    }
    assert_eq!(
        value,
        Value::F64((expected_state & 0x7fffff) as f64 / 8388608.0)
    );
    assert_eq!(vm.global("__vybe_rng"), Some(&Value::I32(expected_state)));
}

#[test]
#[ignore = "interleaved opt-level-0 direct entropy benchmark"]
fn direct_entropy_execution_benchmark() {
    let programs = [
        program(true, Operation::Draws(256)),
        program(false, Operation::Draws(256)),
    ];
    let mut timings = [vec![], vec![]];
    for round in 0..9 {
        for which in [round % 2, 1 - round % 2] {
            let mut vm = runtime(Value::I32(42), Value::Null, Value::Null);
            let chunks = programs[which].clone();
            let start = std::time::Instant::now();
            let value = vm.run(chunks).unwrap();
            timings[which].push(start.elapsed());
            assert_eq!(items(&value).len(), 256);
        }
    }
    for times in &mut timings {
        times.sort();
    }
    println!(
        "random256draws: before={}us after={}us speedup={:.2}x",
        timings[0][4].as_micros(),
        timings[1][4].as_micros(),
        timings[0][4].as_secs_f64() / timings[1][4].as_secs_f64()
    );
}

#[test]
fn native_saturating_f64_to_i32_preserves_boxed_and_host_result_boundaries() {
    for op in [Op::I32_TRUNC_SAT_F64_S, Op::I32_TRUNC_SAT_F64_U] {
        for through_raw_import in [false, true] {
            for as_condition in [false, true] {
                let mut chunks = vec![Chunk::new("<sat-f64-i32>")];
                globals::emit_read(&mut chunks[0], "input", 1);
                if through_raw_import {
                    let to_f64 = chunks[0].add_import("wasm:js-number", "toF64");
                    chunks[0].emit_call(to_f64, 1, 1);
                }
                chunks[0].emit_op(op, 1);
                if as_condition {
                    chunks[0].emit_if_value(1);
                    chunks[0].emit_i32_const(1, 1);
                    chunks[0].emit_else(1);
                    chunks[0].emit_i32_const(0, 1);
                    chunks[0].emit_end(1);
                } else {
                    let local = chunks[0].alloc_scratch(1);
                    chunks[0].emit_op_u16(Op::LOCAL_SET, local, 1);
                    chunks[0].emit_op_u16(Op::LOCAL_GET, local, 1);
                }
                chunks[0].emit_op(Op::RETURN, 1);
                globals::declare_free_globals(&mut chunks);
                globals::normalize_global_table(&mut chunks);
                validate(&chunks);
                for input in [
                    f64::NAN,
                    f64::NEG_INFINITY,
                    f64::INFINITY,
                    -0.0,
                    0.0,
                    -1.9,
                    1.9,
                    i32::MIN as f64 - 1024.0,
                    i32::MIN as f64,
                    i32::MAX as f64,
                    i32::MAX as f64 + 1024.0,
                    u32::MAX as f64,
                    u32::MAX as f64 + 1024.0,
                ] {
                    let converted = if op == Op::I32_TRUNC_SAT_F64_S {
                        input as i32
                    } else {
                        (input as u32) as i32
                    };
                    let expected = if as_condition {
                        i32::from(converted != 0)
                    } else {
                        converted
                    };
                    let mut vm = runtime(Value::I32(42), Value::F64(input), Value::Null);
                    assert_eq!(vm.run(chunks.clone()).unwrap(), Value::I32(expected));
                }
            }
        }
    }
}

#[test]
fn private_entropy_exact_u16_boundary_preserves_stream_and_state() {
    let mut old_vm = runtime(Value::I32(42), Value::Null, Value::Null);
    let expected = items(&old_vm.run(program(true, Operation::Draws(32))).unwrap());
    let expected_state = old_vm.global("__vybe_rng").unwrap().clone();
    for initial_chunks in [u16::MAX as usize, u16::MAX as usize + 1] {
        let mut chunks = unlinked_program(false, Operation::Draws(32));
        chunks.resize_with(initial_chunks, || {
            let mut chunk = Chunk::new("__stdlib_unused_boundary");
            chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 1);
            chunk.emit_op(Op::RETURN, 1);
            chunk
        });
        bundle::finalize_with_runtime_helpers(&mut chunks);
        let fits = initial_chunks == u16::MAX as usize;
        assert_eq!(chunks.len(), initial_chunks + usize::from(fits));
        assert_eq!(
            chunks
                .iter()
                .filter(|chunk| chunk.name == "__stdlib_random_next_unit")
                .count(),
            usize::from(fits)
        );
        assert_eq!(chunks[0].code.len(), chunks[0].lines.len());
        globals::declare_free_globals(&mut chunks);
        globals::normalize_global_table(&mut chunks);
        validate(&chunks);
        let before = chunks[0].code.clone();
        bundle::finalize_with_runtime_helpers(&mut chunks);
        assert_eq!(chunks[0].code, before);
        let mut vm = runtime(Value::I32(42), Value::Null, Value::Null);
        assert_eq!(items(&vm.run(chunks).unwrap()), expected);
        assert_eq!(vm.global("__vybe_rng"), Some(&expected_state));
    }
}
