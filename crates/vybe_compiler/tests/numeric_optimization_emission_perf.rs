use vybe_compiler::primitives::{bundle, convert, globals, platforms};
use vybe_runtime::{Chunk, Op, VM, Value};

#[path = "fixtures/numeric_optimization_baseline.rs"]
mod baseline;

fn program(old: bool, calls: usize) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<numeric-format>")];
    for call in 0..calls {
        globals::emit_read(&mut chunks[0], "input", 0);
        if old {
            baseline::emit_float_to_string(&mut chunks, 0, 0);
        } else {
            convert::emit_float_to_string(&mut chunks, 0, 0);
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
fn runtime(input: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("input", input);
    vm
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}

#[test]
fn double_formatting_preserves_ieee_specials_signed_zero_and_finite_results() {
    let old = program(true, 1);
    let new = program(false, 1);
    validate(&new);
    // The captured formatter has void branches that leave string results.
    assert!(
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(&old))
            .is_err()
    );
    let known = [
        (f64::NAN, "NaN"),
        (f64::INFINITY, "Infinity"),
        (f64::NEG_INFINITY, "-Infinity"),
        (0.0, "0.0"),
        (-0.0, "-0.0"),
        (10.0, "10.0"),
        (-10.0, "-10.0"),
        (1.25, "1.25"),
        (-1.25, "-1.25"),
    ];
    for (input, expected) in known {
        let before = runtime(Value::F64(input)).run(old.clone()).unwrap();
        let after = runtime(Value::F64(input)).run(new.clone()).unwrap();
        assert_eq!(before, after);
        assert_eq!(after, Value::String(expected.into()));
    }
    for bits in [
        1u64,
        0x8000_0000_0000_0001,
        0x7fef_ffff_ffff_ffff,
        0xffef_ffff_ffff_ffff,
        0x7ff8_0000_0000_0001,
        0xfff8_0000_0000_0001,
    ] {
        let input = Value::F64(f64::from_bits(bits));
        assert_eq!(
            runtime(input.clone()).run(old.clone()).unwrap(),
            runtime(input).run(new.clone()).unwrap()
        );
    }
}

#[test]
fn double_formatting_emits_one_finite_body_and_validates_repeated_calls() {
    let old = program(true, 20);
    let new = program(false, 20);
    validate(&new);
    let bytecode = |chunks: &[Chunk]| chunks.iter().map(|c| c.code.len()).sum::<usize>();
    assert!(bytecode(&new) < bytecode(&old));
    assert!(new[0].local_count < old[0].local_count);
    assert!(
        vybe_platform_wasm::write_wasm(&new).len() < vybe_platform_wasm::write_wasm(&old).len()
    );
    println!(
        "double format 20 calls: bytecode={}/{} wasm={}/{} locals={}/{}",
        bytecode(&old),
        bytecode(&new),
        vybe_platform_wasm::write_wasm(&old).len(),
        vybe_platform_wasm::write_wasm(&new).len(),
        old[0].local_count,
        new[0].local_count
    );
}

#[test]
#[ignore = "interleaved opt-level-0 formatting benchmark"]
fn double_formatting_execution_benchmark() {
    let programs = [program(true, 256), program(false, 256)];
    for input in [10.0, -10.0, 1.25, -1.25, -0.0] {
        let mut timings = [vec![], vec![]];
        let mut results = [Value::Null, Value::Null];
        for round in 0..9 {
            for which in [round % 2, 1 - round % 2] {
                let mut vm = runtime(Value::F64(input));
                let chunks = programs[which].clone();
                let start = std::time::Instant::now();
                results[which] = vm.run(chunks).unwrap();
                timings[which].push(start.elapsed());
            }
        }
        assert_eq!(results[0], results[1]);
        for times in &mut timings {
            times.sort();
        }
        println!(
            "double format {input:?} 256 calls: before={}us after={}us speedup={:.2}x",
            timings[0][4].as_micros(),
            timings[1][4].as_micros(),
            timings[0][4].as_secs_f64() / timings[1][4].as_secs_f64()
        );
    }
}
