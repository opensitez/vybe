use vybe_compiler::primitives::{bundle, globals, platforms};
use vybe_runtime::{Chunk, Op, VM, Value};
mod emitter {
    pub use vybe_language_php::emitter::*;
}
#[path = "fixtures/text_optimization_ascii_grouping_accepted.rs"]
mod accepted;
#[path = "fixtures/text_optimization_ascii_grouping_baseline.rs"]
mod baseline;
mod candidate {
    pub use vybe_compiler::primitives::strings::emit_group_digits_ascii_numeric as emit_group_digits;
}
#[path = "fixtures/text_optimization_ascii_grouping_php_producer.rs"]
mod producer;
fn program(old: bool, php: bool, calls: usize) -> Vec<Chunk> {
    program_with_retained(old, php, calls, false)
}
fn program_with_retained(old: bool, php: bool, calls: usize, retained: bool) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<ascii-numeric-grouping>")];
    let current = if retained {
        chunks[0].emit_op_u16(Op::REF_FUNC, 1, 0);
        chunks[0].emit(0, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        chunks.push(Chunk::new("<retained-ascii-numeric-grouping>"));
        1
    } else {
        0
    };
    for index in 0..calls {
        for name in if php {
            vec!["num", "decimals", "decpoint", "separator"]
        } else {
            vec!["formatted", "separator", "decpoint"]
        } {
            globals::emit_read(&mut chunks[current], name, 0);
        }
        if php {
            if old {
                producer::emit_number_format(
                    &mut chunks,
                    current,
                    4,
                    0,
                    baseline::emit_group_digits,
                );
            } else {
                vybe_language_php::emitter::string_adapter::emit_number_format(
                    &mut chunks,
                    current,
                    4,
                    0,
                );
            }
        } else if old {
            baseline::emit_group_digits(&mut chunks, current, 0);
        } else {
            candidate::emit_group_digits(&mut chunks, current, 0);
        }
        if index + 1 != calls {
            chunks[current].emit_op(Op::DROP, 0);
        }
    }
    chunks[current].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn vm() -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm
}
fn validate(chunks: &[Chunk]) -> usize {
    let bytes = vybe_platform_wasm::write_wasm(chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    bytes.len()
}
fn inputs(vm: &mut VM, formatted: &str, num: f64, decimals: f64, sep: Value, dec: Value) {
    for (name, value) in [
        ("formatted", Value::String(formatted.into())),
        ("num", Value::F64(num)),
        ("decimals", Value::F64(decimals)),
        ("separator", sep),
        ("decpoint", dec),
    ] {
        vm.set_global(name, value);
    }
}
fn run(vm: &mut VM, chunks: Vec<Chunk>) -> Result<Value, String> {
    vm.run(chunks).map_err(|mut error| {
        for frame in &mut error.call_stack {
            frame.offset = 0;
        }
        error.to_string()
    })
}
#[test]
fn ascii_group_boundaries_signs_fraction_and_strict_separator_errors_match() {
    let old = program(true, false, 1);
    let new = program(false, false, 1);
    let old_native = validate(&old);
    let new_native = validate(&new);
    let old_ops = ops(&old[0]);
    let new_ops = ops(&new[0]);
    let count = |decoded: &[(usize, Op)], op| {
        decoded
            .iter()
            .filter(|(_, observed)| *observed == op)
            .count()
    };
    assert_eq!(
        count(&old_ops, Op::IF),
        count(&new_ops, Op::IF) + 2,
        "prefix eliminates separator branch and initialization shares the decimal-point test"
    );
    assert!(
        count(&new_ops, Op::I32_EQZ) <= count(&old_ops, Op::I32_EQZ),
        "no extra zero-test"
    );
    let calls = |chunks: &[Chunk]| {
        ops(&chunks[0])
            .into_iter()
            .filter(|(_, op)| *op == Op::CALL)
            .map(|(position, _)| {
                let entry = &chunks[0].imports[local(&chunks[0], position) as usize];
                (entry.module.clone(), entry.name.clone())
            })
            .collect::<Vec<_>>()
    };
    let mut old_calls = calls(&old);
    let mut new_calls = calls(&new);
    old_calls.sort();
    new_calls.sort();
    assert_eq!(
        old_calls, new_calls,
        "same host contracts; pure length/slice reorder"
    );
    assert!(
        new_ops.len() < old_ops.len(),
        "fewer static opcodes than original generic"
    );
    println!(
        "ASCII grouping frozen-body emission bytes={}->{} wasm={old_native}->{new_native} opcodes={}->{} static_calls={}->{} IF={}->{} EQZ={}->{} locals={}->{}",
        old[0].code.len(),
        new[0].code.len(),
        old_ops.len(),
        new_ops.len(),
        count(&old_ops, Op::CALL),
        count(&new_ops, Op::CALL),
        count(&old_ops, Op::IF),
        count(&new_ops, Op::IF),
        count(&old_ops, Op::I32_EQZ),
        count(&new_ops, Op::I32_EQZ),
        old[0].local_count,
        new[0].local_count
    );

    let mut before = vm();
    let mut after = vm();
    let mut texts = [
        "",
        "0",
        "-0",
        "+0",
        "-",
        "+",
        "-.5",
        "+.5",
        ".",
        "123",
        "1234",
        "-123456.75",
        "+12345.",
        ".5",
        "NaN",
        "Infinity",
        "-Infinity",
        "1e+308",
        "1.2.3",
    ]
    .map(str::to_owned)
    .to_vec();
    // Explicit residues include an empty string and each first-group size;
    // expected results come from the public grouping behavior, not a cloned
    // implementation of the arithmetic initializer.
    for (text, expected) in [
        ("", ""),
        ("1", "1"),
        ("12", "12"),
        ("123", "123"),
        ("1234", "1,234"),
        ("12345", "12,345"),
        ("123456", "123,456"),
    ] {
        for (runtime, chunks) in [(&mut before, &old), (&mut after, &new)] {
            inputs(
                runtime,
                text,
                0.0,
                0.0,
                Value::String(",".into()),
                Value::String(".".into()),
            );
            assert_eq!(
                run(runtime, chunks.clone()),
                Ok(Value::String(expected.into()))
            );
        }
    }
    texts.extend((1..=310).map(|n| "1".repeat(n)));
    for text in texts {
        assert!(text.is_ascii());
        for (sep, dec) in [
            (Value::String(",".into()), Value::String(".".into())),
            (Value::String("😀".into()), Value::String("é".into())),
            (Value::String("".into()), Value::String("".into())),
            (Value::Null, Value::String(".".into())),
            (Value::String(",".into()), Value::I32(7)),
        ] {
            inputs(&mut before, &text, 0.0, 0.0, sep.clone(), dec.clone());
            inputs(&mut after, &text, 0.0, 0.0, sep, dec);
            assert_eq!(
                run(&mut before, old.clone()),
                run(&mut after, new.clone()),
                "formatted={text:?}"
            );
        }
    }
}
#[test]
fn actual_php_tofixed_producer_normal_special_and_failure_cases_match() {
    let old = program(true, true, 1);
    let new = program(false, true, 1);
    validate(&old);
    validate(&new);
    let mut before = vm();
    let mut after = vm();
    for num in [
        0.0,
        -0.0,
        1.0,
        -1234.5,
        123456789.125,
        1e95,
        1e308,
        f64::NAN,
        f64::INFINITY,
        f64::NEG_INFINITY,
    ] {
        for decimals in [-1.0, 0.0, 2.0, 100.0, 101.0, f64::NAN] {
            for sep in [
                Value::String(",".into()),
                Value::String("_".into()),
                Value::Null,
            ] {
                inputs(
                    &mut before,
                    "",
                    num,
                    decimals,
                    sep.clone(),
                    Value::String(".".into()),
                );
                inputs(
                    &mut after,
                    "",
                    num,
                    decimals,
                    sep,
                    Value::String(".".into()),
                );
                assert_eq!(
                    run(&mut before, old.clone()),
                    run(&mut after, new.clone()),
                    "num={num} decimals={decimals}"
                );
            }
        }
    }
}
#[test]
fn hostile_non_string_separators_are_not_coerced_or_invoked() {
    use std::sync::{
        Arc,
        atomic::{AtomicUsize, Ordering},
    };
    for old in [true, false] {
        for text in ["123", "1234", "123456.7"] {
            let chunks = program(old, false, 1);
            validate(&chunks);
            let mut runtime = vm();
            let count = Arc::new(AtomicUsize::new(0));
            let observed = count.clone();
            runtime.register_host_fn(
                "test:group",
                "coerce",
                Box::new(move |_, _| {
                    observed.fetch_add(1, Ordering::SeqCst);
                    Value::String("_".into())
                }),
            );
            let index = runtime
                .resolve_host_function_index("test:group", "coerce")
                .unwrap();
            let mut function = vybe_runtime::value::Object::new();
            function.kind = vybe_runtime::value::ObjectKind::HostFunction(index);
            let mut object = vybe_runtime::value::Object::new();
            object.properties.insert(
                "Symbol(toPrimitive)".into(),
                Value::Object(vybe_runtime::heap::alloc(function)),
            );
            inputs(
                &mut runtime,
                text,
                0.0,
                0.0,
                Value::Object(vybe_runtime::heap::alloc(object)),
                Value::String(".".into()),
            );
            let result = run(&mut runtime, chunks);
            if text.len() > 3 {
                assert!(result.unwrap_err().contains("concat"));
            } else {
                assert_eq!(result.unwrap(), Value::String("123".into()));
            }
            assert_eq!(count.load(Ordering::SeqCst), 0);
        }
    }
}
fn parsed_php(body: &str) -> Vec<Chunk> {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_php::register);
    let source = format!(
        "<?php function grouping_subject($n,$d,$point,$sep) {{ return number_format($n,$d,$point,$sep); }} {body}"
    );
    let module = vybe_language_php::parse(&source).unwrap();
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    vybe_compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap()
}
fn ops(chunk: &Chunk) -> Vec<(usize, Op)> {
    let mut offset = 0;
    let mut ops = vec![];
    while offset < chunk.code.len() {
        let op = Op::from_code_at(&chunk.code, offset);
        ops.push((offset, op));
        offset += 4 + op.operand_format().size_in(&chunk.code, offset + 4);
    }
    ops
}
fn local(chunk: &Chunk, offset: usize) -> u16 {
    u16::from_be_bytes([chunk.code[offset + 4], chunk.code[offset + 5]])
}
fn frozen_source(old: &[Chunk]) -> Vec<Chunk> {
    source_oracle(old, baseline::emit_group_digits)
}
fn source_oracle(old: &[Chunk], oracle: fn(&mut [Chunk], usize, u32)) -> Vec<Chunk> {
    let mut result = old.to_vec();
    let index = result
        .iter()
        .position(|c| c.name == "grouping_subject")
        .expect("actual parsed source function");
    let original = &old[index];
    let decoded = ops(original);
    let fixed = decoded
        .iter()
        .position(|(offset, op)| {
            *op == Op::CALL && {
                let i = local(original, *offset) as usize;
                old[0].imports[i].module == "ecma:number" && old[0].imports[i].name == "toFixed"
            }
        })
        .expect("actual direct numeric producer");
    let start = (fixed + 1..decoded.len() - 2)
        .find(|&i| {
            decoded[i].1 == Op::LOCAL_SET
                && decoded[i + 1].1 == Op::LOCAL_SET
                && decoded[i + 2].1 == Op::LOCAL_SET
                && local(original, decoded[i + 1].0) == local(original, decoded[i].0) + 1
                && local(original, decoded[i + 2].0) == local(original, decoded[i].0) + 2
        })
        .expect("group input triple");
    let offset = decoded[start].0;
    let base = local(original, offset);
    let line = original.lines[offset];
    assert!(
        !original.is_async && !original.is_generator,
        "plain source frame only"
    );
    assert!(original.block_i32_results.iter().all(|&key| key < offset));
    assert!(original.call_indirect_sigs.keys().all(|&key| key < offset));
    assert!(original.call_indirect_canon.keys().all(|&key| key < offset));
    assert!(
        original
            .stack_switch_handlers
            .keys()
            .all(|&key| key < offset)
    );
    assert!(
        original
            .user_code_offset
            .is_none_or(|key| (key as usize) < offset)
    );
    let mut depth = 0usize;
    for &(position, op) in decoded.iter().take(start) {
        match op {
            Op::BLOCK | Op::LOOP | Op::IF => depth += 1,
            Op::END => depth = depth.checked_sub(1).expect("balanced source prefix"),
            _ => {}
        }
        assert!(position < offset);
    }
    assert_eq!(
        depth, 0,
        "replacement cannot cut an enclosing control frame"
    );

    // Re-emission uses chunk-local import/constant indices. The parsed module
    // already has finalized module indices: relocate ONLY the appended domain.
    let relocate = |emitted: &mut Chunk| {
        assert_eq!(
            emitted.constants, original.constants,
            "frozen constant pool"
        );
        assert_eq!(
            emitted.global_imports, original.global_imports,
            "frozen global imports"
        );
        for (position, op) in ops(emitted)
            .into_iter()
            .filter(|(position, _)| *position >= offset)
        {
            match op {
                Op::CALL => {
                    let local_index = local(emitted, position) as usize;
                    let import = &emitted.imports[local_index];
                    let module_index = old[0]
                        .imports
                        .iter()
                        .position(|entry| entry == import)
                        .expect("grouping cannot introduce a new module import");
                    let module_index =
                        u16::try_from(module_index).expect("CALL index fits existing ABI");
                    emitted.code[position + 4..position + 6]
                        .copy_from_slice(&module_index.to_be_bytes());
                }
                Op::GLOBAL_GET => {
                    let constant_index = u32::from_be_bytes(
                        emitted.code[position + 4..position + 8].try_into().unwrap(),
                    ) as usize;
                    let Value::String(name) = &emitted.constants[constant_index] else {
                        panic!("grouping global operand must name a string constant");
                    };
                    let module_index = original
                        .globals
                        .iter()
                        .position(|entry| entry == name.as_ref())
                        .expect("grouping cannot introduce a new global");
                    emitted.code[position + 4..position + 8]
                        .copy_from_slice(&(module_index as u32).to_be_bytes());
                }
                Op::GLOBAL_SET => panic!("grouping cannot write module globals"),
                _ => {}
            }
        }
        emitted.imports = original.imports.clone();
    };
    let prepare = || {
        let mut emitted = original.clone();
        emitted.code.truncate(offset);
        emitted.lines.truncate(offset);
        emitted.local_count = base;
        emitted
    };
    let mut frozen = prepare();
    candidate::emit_group_digits(std::slice::from_mut(&mut frozen), 0, line);
    relocate(&mut frozen);
    let end = frozen.code.len();
    assert_eq!(
        frozen.code[offset..],
        original.code[offset..end],
        "selected compiler body reconstructs exact finalized source segment"
    );
    // PHP preserves the returned value across local cleanup, then emits an
    // unreachable fallthrough epilogue. Copy this finalized suffix verbatim.
    // Its supported opcodes contain no bytecode-offset operands or structured
    // branch depths. ALL offset-keyed metadata was checked before the cut, so
    // there are no suffix keys/handler/debug offsets that could become stale.
    let suffix_ops = decoded
        .iter()
        .copied()
        .filter(|(position, _)| *position >= end)
        .collect::<Vec<_>>();
    assert!(
        suffix_ops
            .first()
            .is_some_and(|(position, _)| *position == end),
        "replacement must end at an original instruction boundary"
    );
    assert!(
        suffix_ops.iter().all(|(_, op)| matches!(
            *op,
            Op::LOCAL_SET
                | Op::LOCAL_GET
                | Op::NULL
                | Op::STRUCT_GET
                | Op::CALL
                | Op::DROP
                | Op::RETURN
        )),
        "unsupported cleanup instruction"
    );
    assert_eq!(suffix_ops.last().map(|(_, op)| *op), Some(Op::RETURN));
    let mut changed = prepare();
    oracle(std::slice::from_mut(&mut changed), 0, line);
    // Extra specialization scratch may reuse a later cleanup temporary, but
    // cannot clobber any value live into that suffix. Every overlapping read
    // must follow a suffix write; prefix locals are allocated below `base`.
    assert!(
        decoded
            .iter()
            .take(start)
            .filter(|(_, op)| matches!(*op, Op::LOCAL_GET | Op::LOCAL_SET))
            .all(|(position, _)| local(original, *position) < base)
    );
    let old_group_local_end = frozen.local_count;
    let new_group_local_end = changed.local_count;
    let mut suffix_written = std::collections::HashSet::new();
    for &(position, op) in &suffix_ops {
        if op == Op::LOCAL_SET {
            suffix_written.insert(local(original, position));
        } else if op == Op::LOCAL_GET {
            let slot = local(original, position);
            assert!(
                slot < old_group_local_end
                    || slot >= new_group_local_end
                    || suffix_written.contains(&slot),
                "candidate scratch cannot overwrite a suffix-live local"
            );
        }
    }
    relocate(&mut changed);
    assert_eq!(changed.block_i32_results, original.block_i32_results);
    assert_eq!(changed.call_indirect_sigs, original.call_indirect_sigs);
    assert_eq!(changed.call_indirect_canon, original.call_indirect_canon);
    assert_eq!(
        changed.stack_switch_handlers,
        original.stack_switch_handlers
    );
    let new_end = changed.code.len();
    changed.code.extend_from_slice(&original.code[end..]);
    changed.lines.extend_from_slice(&original.lines[end..]);
    assert_eq!(
        &changed.code[new_end..],
        &original.code[end..],
        "exact cleanup bytes"
    );
    assert_eq!(
        &changed.lines[new_end..],
        &original.lines[end..],
        "exact cleanup debug lines"
    );
    let new_suffix_ops = ops(&changed)
        .into_iter()
        .filter(|(position, _)| *position >= new_end)
        .map(|(position, op)| (position - new_end, op))
        .collect::<Vec<_>>();
    assert_eq!(
        new_suffix_ops,
        suffix_ops
            .iter()
            .map(|(position, op)| (*position - end, *op))
            .collect::<Vec<_>>(),
        "cleanup instruction widths and order are unchanged"
    );
    changed.local_count = changed
        .local_count
        .max(original.local_count)
        .max(changed.scratch_high_water);
    assert_eq!(&changed.code[..offset], &original.code[..offset]);
    assert_eq!(&changed.lines[..offset], &original.lines[..offset]);
    assert_eq!(changed.lines.len(), changed.code.len());
    assert_eq!(changed.arity, original.arity);
    assert_eq!(changed.param_count, original.param_count);
    assert_eq!(changed.result_arity, original.result_arity);
    assert_eq!(changed.takes_receiver, original.takes_receiver);
    assert_eq!(changed.is_method, original.is_method);
    assert_eq!(changed.module_receiver_abi, original.module_receiver_abi);
    assert_eq!(changed.capture_count, original.capture_count);
    assert_eq!(changed.capture_base, original.capture_base);
    assert_eq!(changed.func_sig, original.func_sig);
    assert_eq!(changed.declared_func_type, original.declared_func_type);
    assert_eq!(changed.globals, original.globals);
    result[index] = changed;
    result
}
#[test]
fn actual_parsed_php_source_preserves_arguments_aliases_and_errors() {
    for body in [
        "return grouping_subject(1234567.25,2,'.',',');",
        "return grouping_subject(-0.0,0,'.',',');",
        "return grouping_subject(1234.5,-1,'.',',');",
        "return grouping_subject(1234.5,101,'.',',');",
        "return grouping_subject(1234.5,2,'.',null);",
        "$alias='grouping_subject'; return $alias(1234.5,2,'.',',');",
        "$order=''; function tap($name,$value) { global $order; $order.=$name; return $value; } $result=grouping_subject(tap('n',1234567.25),tap('d',2),tap('p','.'),tap('s',',')); return $order.'|'.$result;",
    ] {
        let new = parsed_php(body);
        let old = frozen_source(&new);
        validate(&old);
        validate(&new);
        let mut before = vm();
        let mut after = vm();
        let old_result = run(&mut before, old);
        let new_result = run(&mut after, new);
        assert_eq!(old_result, new_result, "{body}");
        if body.starts_with("$order=") {
            let expected = Ok(Value::String("ndps|1,234,567.25".into()));
            assert_eq!(
                old_result, expected,
                "positive baseline argument order and output"
            );
            assert_eq!(
                new_result, expected,
                "positive selected argument order and output"
            );
        }
    }
}
#[test]
#[ignore = "opt-level-0 grouping experiment, preparation outside timing"]
fn benchmark_ascii_runs_against_generic_unit_copy() {
    use std::time::Instant;
    for php in [false, true] {
        for (digits, calls) in [(8, 1), (24, 1), (96, 1), (309, 1), (24, 16)] {
            let old = program(true, php, calls);
            let new = program(false, php, calls);
            let old_size = validate(&old);
            let new_size = validate(&new);
            let mut before = vm();
            let mut after = vm();
            let text = "1".repeat(digits);
            let num = 10f64.powi(digits as i32 - 1);
            for runtime in [&mut before, &mut after] {
                inputs(
                    runtime,
                    &text,
                    num,
                    0.0,
                    Value::String(",".into()),
                    Value::String(".".into()),
                );
            }
            let expected = run(&mut before, old.clone()).unwrap();
            let mut times = [vec![], vec![]];
            for sample in 0..7 {
                for variant in if sample % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let (runtime, chunks) = if variant == 0 {
                        (&mut before, &old)
                    } else {
                        (&mut after, &new)
                    };
                    let prepared = chunks.clone();
                    let start = Instant::now();
                    let observed = run(runtime, prepared).unwrap();
                    let elapsed = start.elapsed().as_nanos();
                    assert_eq!(observed, expected);
                    times[variant].push(elapsed);
                }
            }
            for time in &mut times {
                time.sort();
            }
            println!(
                "ASCII grouping php={php} digits={digits} calls={calls} code={}->{} wasm={old_size}->{new_size} old={}ns new={}ns gain={:.2}x",
                old.iter().map(|c| c.code.len()).sum::<usize>(),
                new.iter().map(|c| c.code.len()).sum::<usize>(),
                times[0][3],
                times[1][3],
                times[0][3] as f64 / times[1][3] as f64
            );
        }
    }
}

fn known_native_local_cap(chunks: &[Chunk]) -> String {
    let bytes = vybe_platform_wasm::write_wasm(chunks);
    let error = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .err()
        .expect("known wasmparser MAX50000 local implementation cap");
    assert_eq!(error.message(), "too many locals: locals exceed maximum");
    error.message().to_owned()
}
#[test]
fn candidate_reuses_dead_sign_slot_at_largest_generic_local_budget() {
    for base in [0u16, 1, u16::MAX - 12] {
        let emit = |old: bool| {
            let mut chunks = vec![Chunk::new("<grouping-local-budget>")];
            chunks[0].local_count = base;
            for value in ["1234567", ",", "."] {
                chunks[0].emit_string_const(value, 0);
            }
            if old {
                baseline::emit_group_digits(&mut chunks, 0, 0);
            } else {
                candidate::emit_group_digits(&mut chunks, 0, 0);
            }
            assert_eq!(chunks[0].local_count, base + 12);
            assert_eq!(chunks[0].scratch_high_water, base + 12);
            chunks[0].emit_op(Op::RETURN, 0);
            bundle::finalize_with_runtime_helpers(&mut chunks);
            globals::declare_free_globals(&mut chunks);
            globals::normalize_global_table(&mut chunks);
            chunks
        };
        let old = emit(true);
        let new = emit(false);
        let expected = Ok(Value::String("1,234,567".into()));
        assert_eq!(
            run(&mut vm(), old.clone()),
            expected,
            "baseline full-frame output"
        );
        assert_eq!(
            run(&mut vm(), new.clone()),
            expected,
            "candidate full-frame output"
        );
        if base == u16::MAX - 12 {
            assert_eq!(known_native_local_cap(&old), known_native_local_cap(&new));
            println!(
                "Grouping generic largest frame: VM old/new expected output passed; native both hit wasmparser MAX50000 local implementation cap (not native pass)"
            );
        } else {
            validate(&old);
            validate(&new);
        }
        assert_eq!(old[0].local_count, new[0].local_count);
    }
    let php_emit = |base: u16, old: bool| {
        let mut chunks = vec![Chunk::new("<php-grouping-local-budget>")];
        chunks[0].local_count = base;
        chunks[0].emit_f64_const(1234567.0, 0);
        chunks[0].emit_f64_const(0.0, 0);
        chunks[0].emit_string_const(".", 0);
        chunks[0].emit_string_const(",", 0);
        if old {
            producer::emit_number_format(&mut chunks, 0, 4, 0, baseline::emit_group_digits);
        } else {
            vybe_language_php::emitter::string_adapter::emit_number_format(&mut chunks, 0, 4, 0);
        }
        chunks
    };
    let old = php_emit(0, true);
    let new = php_emit(0, false);
    assert_eq!(old[0].local_count, new[0].local_count);
    assert_eq!(old[0].scratch_high_water, new[0].scratch_high_water);
    let budget = old[0].local_count;
    println!(
        "Grouping local budgets generic=12 candidate=12 PHP={budget} candidate={} delta=0",
        new[0].local_count
    );
    let mut native_caps = vec![];
    for old in [true, false] {
        let mut chunks = php_emit(u16::MAX - budget, old);
        assert_eq!(
            chunks[0].local_count,
            u16::MAX,
            "largest old PHP caller remains accepted"
        );
        chunks[0].emit_op(Op::RETURN, 0);
        bundle::finalize_with_runtime_helpers(&mut chunks);
        globals::declare_free_globals(&mut chunks);
        globals::normalize_global_table(&mut chunks);
        assert_eq!(
            run(&mut vm(), chunks.clone()),
            Ok(Value::String("1,234,567".into())),
            "actual PHP largest frame preserves positive output old={old}"
        );
        native_caps.push(known_native_local_cap(&chunks));
    }
    assert_eq!(native_caps[0], native_caps[1]);
    println!(
        "Grouping PHP largest frame: VM old/new expected output passed; native both hit wasmparser MAX50000 local implementation cap (not native pass)"
    );
}
#[test]
#[ignore = "interleaved opt-level-0 retained-callee grouping and actual PHP Source diagnostic"]
fn benchmark_loaded_grouping_and_actual_php_source() {
    use std::time::Instant;
    for php in [false, true] {
        for (digits, calls) in [(8, 1), (24, 1), (96, 1), (309, 1), (24, 16)] {
            let old = program_with_retained(true, php, calls, true);
            let new = program_with_retained(false, php, calls, true);
            validate(&old);
            validate(&new);
            let mut before = vm();
            let mut after = vm();
            let text = "1".repeat(digits);
            let num = 10f64.powi(digits as i32 - 1);
            for runtime in [&mut before, &mut after] {
                inputs(
                    runtime,
                    &text,
                    num,
                    0.0,
                    Value::String(",".into()),
                    Value::String(".".into()),
                );
            }
            let old_ref = before.run(old).unwrap();
            let new_ref = after.run(new).unwrap();
            let expected = before.invoke(&old_ref, &[]).unwrap();
            let mut times = [vec![], vec![]];
            for sample in 0..7 {
                for variant in if sample % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let (runtime, callee) = if variant == 0 {
                        (&mut before, &old_ref)
                    } else {
                        (&mut after, &new_ref)
                    };
                    let start = Instant::now();
                    let observed = runtime.invoke(callee, &[]);
                    let elapsed = start.elapsed().as_nanos();
                    assert_eq!(observed.unwrap(), expected);
                    times[variant].push(elapsed);
                }
            }
            for times in &mut times {
                times.sort();
            }
            println!(
                "ASCII grouping loaded php={php} digits={digits} calls={calls} old={}ns new={}ns gain={:.2}x",
                times[0][3],
                times[1][3],
                times[0][3] as f64 / times[1][3] as f64
            );
        }
    }
    let new = parsed_php("return null;");
    let old = frozen_source(&new);
    validate(&old);
    validate(&new);
    let alias = old[0]
        .globals
        .iter()
        .filter(|name| name.starts_with("__vybe_func$") && name.ends_with("grouping_subject"))
        .collect::<Vec<_>>();
    assert_eq!(alias.len(), 1, "one actual source callable alias");
    let mut before = vm();
    let mut after = vm();
    before.run(old.clone()).unwrap();
    after.run(new.clone()).unwrap();
    let old_ref = before
        .global(alias[0])
        .expect("source declared callable")
        .clone();
    let new_ref = after
        .global(alias[0])
        .expect("source declared callable")
        .clone();
    for (digits, calls) in [(8, 1), (24, 1), (96, 1), (309, 1), (24, 16)] {
        let args = [
            Value::F64(10f64.powi(digits as i32 - 1)),
            Value::F64(0.0),
            Value::String(".".into()),
            Value::String(",".into()),
        ];
        let expected = before.invoke(&old_ref, &args).unwrap();
        let mut times = [vec![], vec![]];
        for sample in 0..7 {
            for variant in if sample % 2 == 0 { [0, 1] } else { [1, 0] } {
                let (runtime, callee) = if variant == 0 {
                    (&mut before, &old_ref)
                } else {
                    (&mut after, &new_ref)
                };
                let mut observed = None;
                let start = Instant::now();
                for _ in 0..calls {
                    observed = Some(runtime.invoke(callee, &args));
                }
                let elapsed = start.elapsed().as_nanos();
                assert_eq!(observed.unwrap().unwrap(), expected);
                times[variant].push(elapsed);
            }
        }
        for times in &mut times {
            times.sort();
        }
        println!(
            "ASCII grouping actual Source loaded digits={digits} calls={calls} old={}ns new={}ns gain={:.2}x",
            times[0][3],
            times[1][3],
            times[0][3] as f64 / times[1][3] as f64
        );
    }
}

fn cold_batch(runtime: &mut VM, chunks: &[Chunk], iterations: usize, expected: &Value) -> u128 {
    let mut total = 0;
    for _ in 0..iterations {
        let prepared = chunks.to_vec();
        let start = std::time::Instant::now();
        let observed = run(runtime, prepared);
        total += start.elapsed().as_nanos();
        assert_eq!(observed.unwrap(), *expected);
    }
    total
}
fn loaded_batch(
    runtime: &mut VM,
    callee: &Value,
    args: &[Value],
    calls: usize,
    iterations: usize,
    expected: &Value,
) -> u128 {
    let mut observed = Vec::with_capacity(iterations * calls);
    let start = std::time::Instant::now();
    for _ in 0..iterations {
        for _ in 0..calls {
            observed.push(runtime.invoke(callee, args));
        }
    }
    let elapsed = start.elapsed().as_nanos();
    for result in observed {
        assert_eq!(result.unwrap(), *expected);
    }
    elapsed
}
fn batched_report(label: &str, calls: usize, mut measure: impl FnMut(usize, usize) -> u128) {
    // Calibration is outside reported measurements. Use the fastest of three
    // probes for each variant so a scheduler pause does not choose tiny batches.
    let mut fastest = u128::MAX;
    for variant in 0..2 {
        for _ in 0..3 {
            fastest = fastest.min(measure(variant, 1));
        }
    }
    let iterations = (30_000_000u128.div_ceil(fastest.max(1)) as usize).clamp(1, 10_000);
    let mut raw = [vec![], vec![]];
    for sample in 0..7 {
        for variant in if sample % 2 == 0 { [0, 1] } else { [1, 0] } {
            raw[variant].push(measure(variant, iterations));
        }
    }
    let mut paired = raw[0]
        .iter()
        .zip(&raw[1])
        .map(|(old, new)| *old as f64 / *new as f64)
        .collect::<Vec<_>>();
    paired.sort_by(f64::total_cmp);
    println!(
        "ASCII grouping batch {label} iterations={iterations} calls_per_iteration={calls} raw_old_ns={:?} raw_new_ns={:?}",
        raw[0], raw[1]
    );
    for list in &mut raw {
        list.sort();
    }
    println!(
        "ASCII grouping batch {label} median_paired={:.2}x ratio_medians={:.2}x median_old_ns={} median_new_ns={}",
        paired[3],
        raw[0][3] as f64 / raw[1][3] as f64,
        raw[0][3],
        raw[1][3]
    );
}
#[test]
#[ignore = "batched opt-level-0 cold and actual source retained calls under scheduler load"]
fn benchmark_grouping_batched_scheduler_diagnostic() {
    for digits in [8, 24] {
        let old = program(true, true, 1);
        let new = program(false, true, 1);
        validate(&old);
        validate(&new);
        let mut before = vm();
        let mut after = vm();
        for runtime in [&mut before, &mut after] {
            inputs(
                runtime,
                "unused",
                10f64.powi(digits - 1),
                0.0,
                Value::String(",".into()),
                Value::String(".".into()),
            );
        }
        let expected = run(&mut before, old.clone()).unwrap();
        batched_report(
            &format!("PHP cold digits={digits}"),
            1,
            |variant, iterations| {
                if variant == 0 {
                    cold_batch(&mut before, &old, iterations, &expected)
                } else {
                    cold_batch(&mut after, &new, iterations, &expected)
                }
            },
        );
    }
    let new = parsed_php("return null;");
    let old = frozen_source(&new);
    validate(&old);
    validate(&new);
    let aliases = old[0]
        .globals
        .iter()
        .filter(|name| name.starts_with("__vybe_func$") && name.ends_with("grouping_subject"))
        .collect::<Vec<_>>();
    assert_eq!(aliases.len(), 1);
    let mut before = vm();
    let mut after = vm();
    before.run(old.clone()).unwrap();
    after.run(new).unwrap();
    let old_ref = before.global(aliases[0]).unwrap().clone();
    let new_ref = after.global(aliases[0]).unwrap().clone();
    for (digits, calls) in [(8, 1), (24, 1), (309, 1), (24, 16)] {
        let args = [
            Value::F64(10f64.powi(digits - 1)),
            Value::F64(0.0),
            Value::String(".".into()),
            Value::String(",".into()),
        ];
        let expected = before.invoke(&old_ref, &args).unwrap();
        batched_report(
            &format!("actual Source loaded digits={digits}"),
            calls,
            |variant, iterations| {
                if variant == 0 {
                    loaded_batch(&mut before, &old_ref, &args, calls, iterations, &expected)
                } else {
                    loaded_batch(&mut after, &new_ref, &args, calls, iterations, &expected)
                }
            },
        );
    }
}

#[test]
fn generic_unicode_emitter_reduces_instructions_and_retains_frozen_behavior() {
    for text in ["", "é😀123456", "😀123456.é", "-é12345"] {
        let emit = |old: bool| {
            let mut chunks = vec![Chunk::new("<generic-grouping-unchanged>")];
            for value in [text, "😀", "é"] {
                chunks[0].emit_string_const(value, 0);
            }
            if old {
                baseline::emit_group_digits(&mut chunks, 0, 0);
            } else {
                vybe_compiler::primitives::strings::emit_group_digits(&mut chunks, 0, 0);
            }
            chunks[0].emit_op(Op::RETURN, 0);
            chunks
        };
        let mut old = emit(true);
        let mut new = emit(false);
        assert!(
            new[0].code.len() < old[0].code.len(),
            "smaller generic bytecode"
        );
        assert!(
            ops(&new[0]).len() < ops(&old[0]).len(),
            "fewer generic instructions"
        );
        assert_eq!(old[0].constants, new[0].constants);
        assert_eq!(old[0].imports, new[0].imports);
        assert_eq!(old[0].global_imports, new[0].global_imports);
        assert_eq!(old[0].local_count, new[0].local_count);
        assert_eq!(old[0].scratch_high_water, new[0].scratch_high_water);
        for chunks in [&mut old, &mut new] {
            bundle::finalize_with_runtime_helpers(chunks);
            globals::declare_free_globals(chunks);
            globals::normalize_global_table(chunks);
            validate(chunks);
        }
        assert_eq!(
            run(&mut vm(), old),
            run(&mut vm(), new),
            "legacy UTF16 {text:?}"
        );
    }
}

#[test]
fn prefix_refinement_compares_frozen_accepted_source_and_budget() {
    for body in [
        "return grouping_subject(1234.5,2,'.',',');",
        "return grouping_subject(-1234.5,2,'.',',');",
        "return grouping_subject(0,0,'.',null);",
        "return grouping_subject(1234,0,'.',null);",
        "return grouping_subject(1234.5,101,'.',',');",
    ] {
        let new = parsed_php(body);
        let old = source_oracle(&new, accepted::emit_group_digits);
        validate(&old);
        validate(&new);
        let a = old.iter().find(|c| c.name == "grouping_subject").unwrap();
        let b = new.iter().find(|c| c.name == "grouping_subject").unwrap();
        assert!(b.code.len() < a.code.len());
        assert!(ops(b).len() < ops(a).len());
        assert_eq!(a.local_count, b.local_count);
        assert_eq!(a.scratch_high_water, b.scratch_high_water);
        assert_eq!(
            run(&mut vm(), old),
            run(&mut vm(), new),
            "accepted vs prefix {body}"
        );
    }
}
