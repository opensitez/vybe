use vybe_compiler::primitives::{bundle, globals, math, platforms};
use vybe_runtime::{Chunk, Op, VM, Value};

#[path = "fixtures/gamma_optimization_baseline.rs"]
mod baseline;

fn program(old: bool, logarithm: bool) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<gamma-optimization>")];
    globals::emit_read(&mut chunks[0], "input", 23);
    match (old, logarithm) {
        (true, false) => baseline::emit_gamma(&mut chunks[0], 23),
        (false, false) => math::emit_gamma(&mut chunks[0], 23),
        (true, true) => baseline::emit_lgamma(&mut chunks[0], 23),
        (false, true) => math::emit_lgamma(&mut chunks[0], 23),
    }
    chunks[0].emit_op(Op::RETURN, 23);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn runtime(input: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("input", input);
    vm
}
fn normalized(result: Result<Value, vybe_runtime::VMError>) -> Result<u64, String> {
    match result {
        Ok(Value::F64(number)) => Ok(number.to_bits()),
        Ok(other) => panic!("expected f64, got {other:?}"),
        Err(mut error) => {
            for frame in &mut error.call_stack {
                frame.offset = 0;
            }
            Err(error.to_string())
        }
    }
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}

#[test]
fn bounded_factorial_preserves_exact_f64_results_and_special_branches() {
    // Establish the boundary using the same IEEE arithmetic as the emitter.
    let factorial170 = (2..=170).fold(1.0f64, |product, i| product * i as f64);
    assert!(factorial170.is_finite());
    assert!((factorial170 * 171.0).is_infinite());
    for logarithm in [false, true] {
        let old = program(true, logarithm);
        let new = program(false, logarithm);
        validate(&new);
        assert!(
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&vybe_platform_wasm::write_wasm(&old))
                .is_err(),
            "captured gamma uses void branches carrying f64 results"
        );
        let mut inputs: Vec<_> = (1..=200).map(|n| Value::F64(n as f64)).collect();
        inputs.extend([
            Value::F64(1000.0),
            Value::F64(10000.0),
            Value::F64(-0.0),
            Value::F64(0.0),
            Value::F64(-1.0),
            Value::F64(-0.5),
            Value::F64(0.5),
            Value::F64(1.5),
            Value::F64(3.25),
            Value::F64(171.5),
            Value::F64(f64::NAN),
            Value::F64(f64::NEG_INFINITY),
            Value::I32(10),
            Value::Null,
            Value::Undefined,
            Value::Bool(false),
            Value::String("10".into()),
        ]);
        for input in inputs {
            assert_eq!(
                normalized(runtime(input.clone()).run(old.clone())),
                normalized(runtime(input.clone()).run(new.clone())),
                "input={input:?}, logarithm={logarithm}"
            );
        }
        // The captured loop does not terminate for +Infinity. Its positive
        // integral factorial result is +Infinity, now reached in bounded work.
        assert_eq!(
            runtime(Value::F64(f64::INFINITY)).run(new).unwrap(),
            Value::F64(f64::INFINITY)
        );
    }
}

#[test]
fn gamma_keeps_finite_factorial_values_at_the_overflow_boundary() {
    let chunks = program(false, false);
    for (input, expected) in [(1.0, 1.0), (2.0, 1.0), (3.0, 2.0), (6.0, 120.0)] {
        assert_eq!(
            runtime(Value::F64(input)).run(chunks.clone()).unwrap(),
            Value::F64(expected)
        );
    }
    let Value::F64(last_finite) = runtime(Value::F64(171.0)).run(chunks.clone()).unwrap() else {
        panic!("f64");
    };
    assert!(last_finite.is_finite());
    for input in [172.0, 173.0, 1e100] {
        assert_eq!(
            runtime(Value::F64(input)).run(chunks.clone()).unwrap(),
            Value::F64(f64::INFINITY)
        );
    }
}

#[test]
#[ignore = "interleaved opt-level-0 bounded-factorial benchmark"]
fn bounded_factorial_execution_benchmark() {
    let programs = [program(true, false), program(false, false)];
    for input in [10.0, 10000.0] {
        let mut times = [vec![], vec![]];
        let mut results = [Value::Null, Value::Null];
        for round in 0..9 {
            for which in [round % 2, 1 - round % 2] {
                let mut vm = runtime(Value::F64(input));
                let chunks = programs[which].clone();
                let start = std::time::Instant::now();
                results[which] = vm.run(chunks).unwrap();
                times[which].push(start.elapsed());
            }
        }
        assert_eq!(results[0], results[1]);
        for samples in &mut times {
            samples.sort();
        }
        println!(
            "gamma{input}9samples: before={}us after={}us speedup={:.2}x",
            times[0][4].as_micros(),
            times[1][4].as_micros(),
            times[0][4].as_secs_f64() / times[1][4].as_secs_f64()
        );
    }
}
