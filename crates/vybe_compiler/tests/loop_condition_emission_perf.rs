use vybe_compiler::primitives::{bundle, collections, globals, platforms};
use vybe_runtime::value::ObjectKind;
use vybe_runtime::{Chunk, Op, VM, Value};

#[path = "fixtures/coercing_range_condition.rs"]
mod baseline;

fn program(old: bool, count: u8, inclusive: bool) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<range-condition>")];
    for name in ["first", "second", "third"]
        .into_iter()
        .take(count as usize)
    {
        globals::emit_read(&mut chunks[0], name, 0);
    }
    if old {
        baseline::emit_range(&mut chunks, 0, count, inclusive, 0);
    } else {
        collections::emit_range(&mut chunks, 0, count, inclusive, 0);
    }
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}

fn vm(values: &[Value]) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    for (key, value) in ["first", "second", "third"].into_iter().zip(values) {
        vm.set_global(key, value.clone());
    }
    vm
}

fn items(value: Value) -> Vec<Value> {
    let Value::Object(object) = value else {
        panic!("expected array")
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else {
        panic!("expected array")
    };
    values.clone()
}

#[test]
fn direct_conditions_preserve_dynamic_range_comparisons_and_results() {
    let cases = [
        (vec![Value::I32(4)], false, vec!["0", "1", "2", "3"]),
        (
            vec![Value::I32(2), Value::I32(4)],
            true,
            vec!["2", "3", "4"],
        ),
        (
            vec![Value::I32(4), Value::I32(-1), Value::I32(-2)],
            false,
            vec!["4", "2", "0"],
        ),
        (
            vec![Value::F64(0.5), Value::F64(2.5), Value::F64(0.5)],
            false,
            vec!["0.5", "1", "1.5", "2"],
        ),
        (
            vec![Value::String("a".into()), Value::String("c".into())],
            true,
            vec!["a", "b", "c"],
        ),
        (vec![Value::I32(-3)], false, vec![]),
    ];
    for (values, inclusive, expected) in cases {
        let old = program(true, values.len() as u8, inclusive);
        let new = program(false, values.len() as u8, inclusive);
        for chunks in [&old, &new] {
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&vybe_platform_wasm::write_wasm(chunks))
                .unwrap();
        }
        let a = items(vm(&values).run(old).unwrap());
        let b = items(vm(&values).run(new).unwrap());
        assert_eq!(a, b);
        assert_eq!(b.iter().map(Value::to_string).collect::<Vec<_>>(), expected);
    }
}

#[test]
fn direct_branch_consumer_reduces_emitted_bytecode() {
    let old = program(true, 3, false);
    let new = program(false, 3, false);
    let bytes = |chunks: &[Chunk]| chunks.iter().map(|c| c.code.len()).sum::<usize>();
    assert!(bytes(&new) < bytes(&old));
    println!(
        "range: bytecode={}/{} wasm={}/{}",
        bytes(&old),
        bytes(&new),
        vybe_platform_wasm::write_wasm(&old).len(),
        vybe_platform_wasm::write_wasm(&new).len()
    );
}

#[test]
#[ignore = "interleaved opt-level-0 execution benchmark"]
fn direct_range_condition_execution_benchmark() {
    let values = [Value::I32(8192)];
    let programs = [program(true, 1, false), program(false, 1, false)];
    let mut timings = [vec![], vec![]];
    let mut results = [Value::Null, Value::Null];
    for round in 0..9 {
        for which in [round % 2, 1 - round % 2] {
            let mut runtime = vm(&values);
            let chunks = programs[which].clone();
            let start = std::time::Instant::now();
            results[which] = runtime.run(chunks).unwrap();
            timings[which].push(start.elapsed());
        }
    }
    assert_eq!(items(results[0].clone()), items(results[1].clone()));
    for times in &mut timings {
        times.sort();
    }
    println!(
        "range 8192: before={}us after={}us speedup={:.2}x",
        timings[0][4].as_micros(),
        timings[1][4].as_micros(),
        timings[0][4].as_secs_f64() / timings[1][4].as_secs_f64()
    );
}

#[test]
fn stop_range_preserves_tags_and_non_numeric_outcomes() {
    for inclusive in [false, true] {
        for stop in [Value::I32(8), Value::F64(3.5), Value::F64(-0.0),
            Value::F64(f64::NAN), Value::Null, Value::Undefined,
            Value::Bool(true), Value::String("3".into()),
            Value::String("nonnumeric".into())] {
            let args = [stop];
            let old = vm(&args).run(program(true, 1, inclusive)).map(items).map_err(|e| e.message);
            let new = vm(&args).run(program(false, 1, inclusive)).map(items).map_err(|e| e.message);
            assert_eq!(old, new, "{args:?}, inclusive={inclusive}");
            if matches!(args[0], Value::I32(8)) {
                let values = new.unwrap();
                assert_eq!(values.len(), if inclusive { 9 } else { 8 });
                assert!(matches!(values[0], Value::I32(0)));
                assert!(matches!(values[1], Value::F64(1.0)));
            }
        }
    }
}

#[test]
fn implicit_unit_step_preserves_dynamic_start_coercion() {
    for inclusive in [false, true] {
        for start in [Value::Null, Value::Undefined, Value::Bool(false),
            Value::Bool(true), Value::F64(-0.0), Value::F64(f64::NAN),
            Value::F64(0.5), Value::String("".into()), Value::String("a".into())] {
            let stop = if matches!(start, Value::String(_)) {
                Value::String("d".into())
            } else { Value::I32(3) };
            let args = [start, stop];
            let old = vm(&args).run(program(true, 2, inclusive)).map(items).map_err(|e| e.message);
            let new = vm(&args).run(program(false, 2, inclusive)).map(items).map_err(|e| e.message);
            assert_eq!(old, new, "{args:?}, inclusive={inclusive}");
        }
    }
}
