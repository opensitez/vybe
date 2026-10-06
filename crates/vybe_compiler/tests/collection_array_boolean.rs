use vybe_compiler::primitives::{bundle, collections, globals, platforms};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};

#[path = "fixtures/collection_array_boolean_before.rs"]
mod before;

type Emitter = fn(&mut [Chunk], usize, u32);

fn array(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}

fn program(emit: Emitter) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<array-boolean>")];
    globals::emit_read(&mut chunks[0], "input", 0);
    emit(&mut chunks, 0, 0);
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
        .unwrap();
    chunks
}

fn run(chunks: Vec<Chunk>, input: Value) -> Value {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("input", input);
    vm.run(chunks).unwrap()
}

#[test]
fn ranks_preserve_scalar_empty_and_nested_array_results() {
    let old = program(before::emit_rank);
    let new = program(collections::emit_rank);
    assert!(new[0].code.len() < old[0].code.len());
    assert_eq!(new[0].local_count, old[0].local_count);
    eprintln!(
        "rank bytes {} -> {}, locals {} -> {}",
        old[0].code.len(),
        new[0].code.len(),
        old[0].local_count,
        new[0].local_count
    );
    for (value, expected) in [
        (Value::Null, 0),
        (Value::Undefined, 0),
        (Value::Bool(false), 0),
        (Value::I32(7), 0),
        (Value::String("hello".into()), 0),
        (array(vec![]), 1),
        (array(vec![Value::I32(9)]), 1),
        (array(vec![array(vec![])]), 2),
        (array(vec![array(vec![array(vec![])])]), 3),
    ] {
        for chunks in [&old, &new] {
            assert!(matches!(run(chunks.clone(), value.clone()), Value::I32(n) if n == expected));
        }
    }
}

#[test]
fn pop_and_shift_keep_results_and_mutation() {
    for (old_emit, new_emit, expected, remaining) in [
        (
            before::emit_pop as Emitter,
            collections::emit_pop as Emitter,
            3,
            vec![1, 2],
        ),
        (
            before::emit_shift as Emitter,
            collections::emit_shift as Emitter,
            1,
            vec![2, 3],
        ),
    ] {
        let old = program(old_emit);
        let new = program(new_emit);
        assert!(new[0].code.len() < old[0].code.len());
        eprintln!(
            "pop/shift bytes {} -> {}",
            old[0].code.len(),
            new[0].code.len()
        );
        for chunks in [old, new] {
            let input = array(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
            assert!(matches!(run(chunks, input.clone()), Value::I32(n) if n == expected));
            let Value::Object(object) = input else {
                unreachable!()
            };
            let object = object.lock().unwrap();
            let ObjectKind::Array(values) = &object.kind else {
                panic!("array representation changed")
            };
            assert_eq!(values.len(), remaining.len());
            for (value, expected) in values.iter().zip(&remaining) {
                assert!(matches!(value, Value::I32(n) if n == expected));
            }
        }
    }
}

#[test]
fn natural_iteration_and_clear_preserve_array_identity() {
    for (old_emit, new_emit) in [
        (
            before::emit_iter_natural as Emitter,
            collections::emit_iter_natural as Emitter,
        ),
        (
            before::emit_clear_keyed as Emitter,
            collections::emit_clear_keyed as Emitter,
        ),
    ] {
        let old = program(old_emit);
        let new = program(new_emit);
        assert!(new[0].code.len() < old[0].code.len());
        for chunks in [old, new] {
            let input = array(vec![Value::I32(42)]);
            let result = run(chunks, input.clone());
            if std::ptr::fn_addr_eq(new_emit, collections::emit_iter_natural as Emitter) {
                let (Value::Object(a), Value::Object(b)) = (&input, result) else {
                    panic!("expected array")
                };
                assert!(std::sync::Arc::ptr_eq(a, &b));
            } else {
                assert!(matches!(result, Value::Null));
            }
            let Value::Object(object) = input else {
                unreachable!()
            };
            let object = object.lock().unwrap();
            assert!(matches!(&object.kind, ObjectKind::Array(values) if values.len() == 1));
        }
    }
}

#[test]
fn promotion_keeps_empty_array_and_string_key_guards() {
    for old in [true, false] {
        let mut chunks = vec![Chunk::new("<array-promotion>")];
        let object_slot = chunks[0].alloc_scratch(1);
        let key_slot = chunks[0].alloc_scratch(1);
        globals::emit_read(&mut chunks[0], "input", 0);
        chunks[0].emit_op_u16(Op::LOCAL_SET, object_slot, 0);
        globals::emit_read(&mut chunks[0], "key", 0);
        chunks[0].emit_op_u16(Op::LOCAL_SET, key_slot, 0);
        if old {
            before::emit_promote_empty_array_for_string_key(
                &mut chunks,
                0,
                object_slot,
                key_slot,
                0,
            );
        } else {
            collections::emit_promote_empty_array_for_string_key(
                &mut chunks,
                0,
                object_slot,
                key_slot,
                0,
            );
        }
        chunks[0].emit_op_u16(Op::LOCAL_GET, object_slot, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        bundle::finalize_with_runtime_helpers(&mut chunks);
        globals::declare_free_globals(&mut chunks);
        globals::normalize_global_table(&mut chunks);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
            .unwrap();
        for (values, key, promote) in [
            (vec![], Value::String("key".into()), true),
            (vec![], Value::I32(0), false),
            (vec![Value::I32(1)], Value::String("key".into()), false),
        ] {
            let input = array(values);
            let mut vm = VM::new();
            platforms::register_platforms_all(&mut vm);
            vm.set_global("input", input.clone());
            vm.set_global("key", key);
            let Value::Object(result) = vm.run(chunks.clone()).unwrap() else {
                panic!("expected collection")
            };
            let result_guard = result.lock().unwrap();
            if promote {
                assert!(matches!(&result_guard.kind, ObjectKind::Map(_)));
            } else {
                assert!(matches!(&result_guard.kind, ObjectKind::Array(_)));
                let Value::Object(original) = &input else {
                    unreachable!()
                };
                assert!(std::sync::Arc::ptr_eq(original, &result));
            }
        }
    }
}

#[test]
fn paired_debug_rank_execution_measurement() {
    let mut runners = Vec::new();
    for emit in [
        before::emit_rank as Emitter,
        collections::emit_rank as Emitter,
    ] {
        let mut helper = Chunk::new("<rank-benchmark>");
        helper.arity = 1;
        helper.local_count = 1;
        helper.emit_op_u16(Op::LOCAL_GET, 0, 0);
        let mut body = vec![helper];
        emit(&mut body, 0, 0);
        body[0].emit_op(Op::RETURN, 0);
        let mut root = Chunk::new("<rank-benchmark-root>");
        root.emit_op_u16(Op::REF_FUNC, 1, 0);
        root.emit(0, 0);
        root.emit_op(Op::RETURN, 0);
        let mut chunks = vec![root, body.remove(0)];
        bundle::finalize_with_runtime_helpers(&mut chunks);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
            .unwrap();
        let mut vm = VM::new();
        platforms::register_platforms_all(&mut vm);
        let callable = vm.run(chunks).unwrap();
        runners.push((vm, callable));
    }
    let args = [array(vec![array(vec![array(vec![])])])];
    let mut n = 128;
    let measure = |runner: &mut (VM, Value), n: usize| {
        let start = std::time::Instant::now();
        for _ in 0..n {
            let result = runner.0.invoke(&runner.1, &args).unwrap();
            assert!(matches!(std::hint::black_box(result), Value::I32(3)));
        }
        start.elapsed().as_secs_f64()
    };
    while measure(&mut runners[1], n) < 0.03 {
        n *= 2;
    }
    let mut ratios = Vec::new();
    for round in 0..7 {
        let mut elapsed = [0.0; 2];
        for index in [round % 2, 1 - round % 2] {
            elapsed[index] = measure(&mut runners[index], n);
        }
        ratios.push(elapsed[0] / elapsed[1]);
        eprintln!(
            "rank debug pair {round}: n={n}, before={}ms after={}ms ratio={}",
            elapsed[0] * 1000.0,
            elapsed[1] * 1000.0,
            ratios.last().unwrap()
        );
    }
    ratios.sort_by(f64::total_cmp);
    eprintln!(
        "rank debug paired median speedup={}x range={}..{}",
        ratios[3], ratios[0], ratios[6]
    );
}
