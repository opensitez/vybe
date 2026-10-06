use vybe_ast::datetime::{DateTimePolicy, MonthIndexing};
use vybe_compiler::primitives::{bundle, datetime, globals, platforms};
use vybe_runtime::{Chunk, Op, VM, Value};

#[path = "fixtures/datetime_optimization_baseline.rs"]
mod baseline;

#[derive(Clone, Copy, Debug)]
enum Operation {
    Days(MonthIndexing),
    AddMonths(DateTimePolicy),
}
fn program(old: bool, operation: Operation, calls: usize) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<datetime-optimization>")];
    for call in 0..calls {
        globals::emit_read(&mut chunks[0], "first", 29);
        globals::emit_read(&mut chunks[0], "second", 29);
        match (old, operation) {
            (true, Operation::Days(indexing)) => {
                baseline::emit_days_in_month(&mut chunks[0], indexing, 29)
            }
            (false, Operation::Days(indexing)) => {
                datetime::emit_days_in_month(&mut chunks[0], indexing, 29)
            }
            (true, Operation::AddMonths(policy)) => {
                baseline::emit_add_months(&mut chunks[0], policy, 29)
            }
            (false, Operation::AddMonths(policy)) => {
                datetime::emit_add_months(&mut chunks[0], policy, 29)
            }
        }
        if call + 1 != calls {
            chunks[0].emit_op(Op::DROP, 29);
        }
    }
    chunks[0].emit_op(Op::RETURN, 29);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn runtime(first: Value, second: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("first", first);
    vm.set_global("second", second);
    vm
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}
fn normalized(result: Result<Value, vybe_runtime::VMError>) -> Result<u64, String> {
    match result {
        Ok(Value::F64(number)) => Ok(number.to_bits()),
        Ok(other) => panic!("expectedf64, got{other:?}"),
        Err(mut error) => {
            for frame in &mut error.call_stack {
                frame.offset = 0;
            }
            Err(error.to_string())
        }
    }
}

#[test]
fn month_day_counts_preserve_calendar_and_caller_coercions() {
    for indexing in [MonthIndexing::ZeroBased, MonthIndexing::OneBased] {
        let old = program(true, Operation::Days(indexing), 1);
        let new = program(false, Operation::Days(indexing), 1);
        validate(&old);
        validate(&new);
        for year in [-400, -100, -4, 0, 1900, 2000, 2023, 2024] {
            for month in 1..=12 {
                let input_month = if indexing == MonthIndexing::ZeroBased {
                    month - 1
                } else {
                    month
                };
                let expected = match month {
                    2 if (year % 4 == 0 && year % 100 != 0) || year % 400 == 0 => 29.0,
                    2 => 28.0,
                    4 | 6 | 9 | 11 => 30.0,
                    _ => 31.0,
                };
                for chunks in [&old, &new] {
                    assert_eq!(
                        runtime(Value::I32(year), Value::I32(input_month))
                            .run(chunks.clone())
                            .unwrap(),
                        Value::F64(expected)
                    );
                }
            }
        }
        for month in [
            Value::Null,
            Value::Undefined,
            Value::Bool(false),
            Value::Bool(true),
            Value::F64(-0.0),
            Value::F64(f64::NAN),
            Value::F64(f64::INFINITY),
            Value::F64(-1.0),
            Value::F64(12.0),
            Value::F64(1.5),
            Value::String("1".into()),
            Value::String("2".into()),
        ] {
            assert_eq!(
                normalized(runtime(Value::I32(2024), month.clone()).run(old.clone())),
                normalized(runtime(Value::I32(2024), month.clone()).run(new.clone())),
                "{indexing:?} month={month:?}"
            );
        }
    }
}

#[test]
fn calendar_addition_preserves_clamp_overflow_negative_offsets_and_clock() {
    for policy in [DateTimePolicy::ECMA, DateTimePolicy::ISO] {
        let old = program(true, Operation::AddMonths(policy), 1);
        let new = program(false, Operation::AddMonths(policy), 1);
        validate(&old);
        validate(&new);
        let clamp = datetime::month_add_clamps(policy);
        for (input, delta, expected_clamp, expected_overflow) in [
            (1706704496000.0, 1.0, 1709210096000.0, 1709382896000.0),
            (1675168496000.0, 1.0, 1677587696000.0, 1677846896000.0),
            (1704026096000.0, -1.0, 1701347696000.0, 1701434096000.0),
            (-41104000.0, -1.0, -2719504000.0, -2633104000.0),
        ] {
            let expected = Value::F64(if clamp {
                expected_clamp
            } else {
                expected_overflow
            });
            for chunks in [&old, &new] {
                assert_eq!(
                    runtime(Value::F64(input), Value::F64(delta))
                        .run(chunks.clone())
                        .unwrap(),
                    expected
                );
            }
        }
        for input in [1706704496000.0, -41104000.0, 0.0, -0.0, f64::NAN] {
            for count in [
                Value::I32(-25),
                Value::I32(-12),
                Value::I32(0),
                Value::I32(13),
                Value::F64(-0.5),
                Value::F64(0.5),
                Value::F64(24.0),
                Value::F64(f64::NAN),
                Value::Null,
                Value::Bool(true),
                Value::String("1".into()),
            ] {
                assert_eq!(
                    normalized(runtime(Value::F64(input), count.clone()).run(old.clone())),
                    normalized(runtime(Value::F64(input), count.clone()).run(new.clone())),
                    "policy={policy:?} input={input:?} count={count:?}"
                );
            }
        }
    }
}

#[test]
fn proven_month_specialization_reduces_emission_and_keeps_other_policies() {
    for operation in [
        Operation::Days(MonthIndexing::ZeroBased),
        Operation::Days(MonthIndexing::OneBased),
        Operation::AddMonths(DateTimePolicy::ISO),
        Operation::AddMonths(DateTimePolicy::ECMA),
    ] {
        let old = program(true, operation, 20);
        let new = program(false, operation, 20);
        validate(&new);
        let bytes = |chunks: &[Chunk]| chunks.iter().map(|c| c.code.len()).sum::<usize>();
        if matches!(operation, Operation::Days(MonthIndexing::OneBased))
            || matches!(operation, Operation::AddMonths(policy) if !datetime::month_add_clamps(policy))
        {
            assert_eq!(bytes(&old), bytes(&new));
            assert_eq!(old[0].code, new[0].code);
        } else {
            assert!(bytes(&new) < bytes(&old));
            assert!(
                vybe_platform_wasm::write_wasm(&new).len()
                    < vybe_platform_wasm::write_wasm(&old).len()
            );
        }
        println!(
            "datetime{operation:?}20calls: bytecode={}/{} wasm={}/{}",
            bytes(&old),
            bytes(&new),
            vybe_platform_wasm::write_wasm(&old).len(),
            vybe_platform_wasm::write_wasm(&new).len()
        );
    }
}

#[test]
#[ignore = "interleaved opt-level-0 calendar benchmark"]
fn calendar_execution_benchmark() {
    for (operation, first, second) in [
        (
            Operation::Days(MonthIndexing::ZeroBased),
            Value::I32(2024),
            Value::I32(10),
        ),
        (
            Operation::AddMonths(DateTimePolicy::ISO),
            Value::F64(1706704496000.0),
            Value::I32(1),
        ),
    ] {
        let programs = [
            program(true, operation, 256),
            program(false, operation, 256),
        ];
        let mut times = [vec![], vec![]];
        let mut results = [Value::Null, Value::Null];
        for round in 0..9 {
            for which in [round % 2, 1 - round % 2] {
                let mut vm = runtime(first.clone(), second.clone());
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
            "datetime{operation:?}256calls: before={}us after={}us speedup={:.2}x",
            times[0][4].as_micros(),
            times[1][4].as_micros(),
            times[0][4].as_secs_f64() / times[1][4].as_secs_f64()
        );
    }
}
