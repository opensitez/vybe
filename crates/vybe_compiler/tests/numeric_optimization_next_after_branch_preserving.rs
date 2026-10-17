//! Frozen behavior and actual production proof for next_after scalar-copy elimination.
use vybe_ast::FloatLane;
use vybe_compiler::primitives::{bundle, callable, globals, math, platforms};
use vybe_runtime::chunk::{ConstExpr, GlobalInit};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/numeric_optimization_next_after_baseline.rs"]
mod baseline;
#[path = "fixtures/numeric_optimization_next_after_candidate.rs"]
mod candidate;
#[path = "fixtures/nextafter_ordered_before.rs"]
mod before_ordered;
#[path="fixtures/nextafter_copysign_before.rs"]
mod ordered_stage;

fn instructions(code: &[u8]) -> Vec<Vec<u8>> {
    let mut rows = vec![];
    let mut offset = 0;
    while offset < code.len() {
        let op = Op::decode(
            u16::from_be_bytes([code[offset], code[offset + 1]]),
            u16::from_be_bytes([code[offset + 2], code[offset + 3]]),
        )
        .unwrap();
        let size = 4 + op.operand_format().size_in(code, offset + 4);
        rows.push(code[offset..offset + size].to_vec());
        offset += size;
    }
    rows
}
fn opcode(row: &[u8]) -> Op {
    Op::decode(
        u16::from_be_bytes([row[0], row[1]]),
        u16::from_be_bytes([row[2], row[3]]),
    )
    .unwrap()
}
#[test]
fn only_two_scalar_copies_and_local_relocation_change_the_kernel() {
    for lane in [FloatLane::F32, FloatLane::F64] {
        for base in [0u16, 9] {
            let mut old = Chunk::new("old");
            old.local_count = base;
            let mut new = old.clone();
            baseline::emit_next_after(&mut old, lane, 37);
            candidate::emit_next_after(&mut new, lane, 37);
            let rows = instructions(&old.code);
            let mut normalized = vec![];
            let (mut index, mut removed) = (0, 0);
            while index < rows.len() {
                let get = &rows[index];
                if opcode(get) == Op::LOCAL_GET
                    && u16::from_be_bytes([get[4], get[5]]) == base
                    && index + 1 < rows.len()
                {
                    let set = &rows[index + 1];
                    if opcode(set) == Op::LOCAL_SET
                        && matches!(u16::from_be_bytes([set[4], set[5]]), slot if slot == base + 2 || slot == base + 3)
                    {
                        index += 2;
                        removed += 1;
                        continue;
                    }
                }
                let mut row = get.clone();
                if opcode(&row) == Op::LOCAL_GET || opcode(&row) == Op::LOCAL_SET {
                    let slot = u16::from_be_bytes([row[4], row[5]]);
                    if slot == base + 2 || slot == base + 3 {
                        row[4..6].copy_from_slice(&base.to_be_bytes());
                    }
                }
                normalized.extend(row);
                index += 1;
            }
            assert_eq!(removed, 2, "one old copy in each direction branch");
            assert_eq!(
                normalized, new.code,
                "all constants/conditions/branches/reinterprets must remain exact"
            );
            assert_eq!(old.code.len() - new.code.len(), 24);
            assert_eq!(old.local_count, base + 4);
            assert_eq!(new.local_count, base + 2);
            assert_eq!(old.block_i32_results, new.block_i32_results);
            for up in [true, false] {
                let mut old = Chunk::new("toward-old");
                old.local_count = base;
                let mut new = old.clone();
                baseline::emit_next_toward(&mut old, up, lane, 37);
                candidate::emit_next_toward(&mut new, up, lane, 37);
                let mut production = Chunk::new("toward-production");
                production.local_count = base;
                ordered_stage::emit_next_toward(&mut production, up, lane, 37);
                assert_eq!(
                    production.code, old.code,
                    "historical next_toward stage stays frozen"
                );
                assert_eq!(production.lines, old.lines);
                assert_eq!(production.local_count, old.local_count);
                assert_eq!(old.code, new.code, "public next_toward emission unchanged");
                assert_eq!(old.lines, new.lines);
                assert_eq!(old.local_count, new.local_count);
            }
            let mut actual_ulp = Chunk::new("ulp");
            actual_ulp.local_count = base;
            ordered_stage::emit_ulp(&mut actual_ulp, lane, 37);
            let mut proposed_ulp = Chunk::new("ulp-proposed");
            proposed_ulp.local_count = base;
            let x = proposed_ulp.alloc_scratch(1);
            proposed_ulp.emit_op(Op::F64_ABS, 37);
            proposed_ulp.emit_op_u16(Op::LOCAL_SET, x, 37);
            proposed_ulp.emit_op_u16(Op::LOCAL_GET, x, 37);
            candidate::emit_next_toward(&mut proposed_ulp, true, lane, 37);
            proposed_ulp.emit_op_u16(Op::LOCAL_GET, x, 37);
            proposed_ulp.emit_op(Op::F64_SUB, 37);
            assert_eq!(
                actual_ulp.code, proposed_ulp.code,
                "public ULP's next_toward use unchanged"
            );
            assert_eq!(actual_ulp.local_count, proposed_ulp.local_count);
        }
    }
}
fn runtime(x: Value, y: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("x", x);
    vm.set_global("y", y);
    vm
}
fn program(old: bool, lane: FloatLane, calls: usize) -> Vec<Chunk> {
    let mut c = vec![
        Chunk::new("<next-after-experiment>"),
        Chunk::new("<next-after-callee>"),
    ];
    c[1].arity = 2;
    c[1].local_count = 2;
    for slot in 0..2 {
        c[1].emit_op_u16(Op::LOCAL_GET, slot, 37);
    }
    if old {
        baseline::emit_next_after(&mut c[1], lane, 37)
    } else {
        candidate::emit_next_after(&mut c[1], lane, 37)
    }
    c[1].emit_op(Op::RETURN, 37);
    c[0].global_inits.push(GlobalInit {
        name: "next_after_callee".into(),
        init: ConstExpr::RefFunc(1),
    });
    for n in 0..calls {
        for name in ["next_after_callee", "x", "y"] {
            globals::emit_read(&mut c[0], name, 37);
        }
        callable::emit_direct_invoke_chunk(&mut c[0], 2, 37);
        if n + 1 != calls {
            c[0].emit_op(Op::DROP, 37);
        }
    }
    c[0].emit_op(Op::RETURN, 37);
    bundle::finalize_with_runtime_helpers(&mut c);
    globals::declare_free_globals(&mut c);
    globals::normalize_global_table(&mut c);
    c
}
fn validate_result(c: &[Chunk], label: &str) -> Result<(), String> {
    let bytes = vybe_platform_wasm::write_wasm(c);
    if let Err(e) = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
    {
        eprintln!("next_after {label} native failure {e}");
        let path = format!(
            "/Users/youness/www/html/vybe/tools/primitive_experiments/next_after/{label}-native-failure.wasm"
        );
        std::fs::write(path, &bytes).unwrap();
        let named = vybe_platform_wasm::write_wasm_with_options(
            c,
            vybe_platform_wasm::WasmWriteOptions {
                include_name_section: true,
                ..Default::default()
            },
        );
        let mut names = std::collections::BTreeMap::new();
        for payload in wasmparser::Parser::new(0).parse_all(&named) {
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
        for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
            match payload.unwrap() {
                wasmparser::Payload::ImportSection(section) => {
                    for import in section.into_imports() {
                        if matches!(
                            import.unwrap().ty,
                            wasmparser::TypeRef::Func(_) | wasmparser::TypeRef::FuncExact(_)
                        ) {
                            function += 1;
                        }
                    }
                }
                wasmparser::Payload::CodeSectionEntry(body) => {
                    if body.range().contains(&e.offset()) {
                        eprintln!(
                            "next_after {label} function={function} name={:?} range={:?}",
                            names.get(&function),
                            body.range()
                        );
                        let mut reader = body.get_operators_reader().unwrap();
                        while !reader.eof() {
                            let (op, offset) = reader.read_with_offset().unwrap();
                            if offset + 140 >= e.offset() && offset <= e.offset() + 40 {
                                eprintln!("  wasm@{offset}: {op:?}");
                                if let wasmparser::Operator::Call { function_index } = op {
                                    eprintln!("    target={:?}", names.get(&function_index));
                                }
                            }
                        }
                    }
                    function += 1;
                }
                _ => {}
            }
        }
        return Err(e.to_string());
    }
    Ok(())
}
fn validate(c: &[Chunk]) {
    validate_result(c, "single").unwrap();
}
fn validate_pair(before: &[Chunk], after: &[Chunk]) {
    let before = validate_result(before, "frozen");
    let after = validate_result(after, "candidate");
    before.unwrap();
    after.unwrap();
}
fn exact(v: &Value) -> String {
    match v {
        Value::F64(x) => format!("F64:{:016x}", x.to_bits()),
        Value::F32(x) => format!("F32:{:08x}", x.to_bits()),
        _ => format!("{v:?}"),
    }
}
fn error(e: vybe_runtime::VMError) -> (String, bool, Vec<(String, Option<u32>)>) {
    (
        e.message.clone(),
        e.is_trap(),
        e.call_stack
            .iter()
            .map(|f| (f.chunk_name.clone(), f.line))
            .collect(),
    )
}
#[test]
fn actual_production_matches_copy_elimination_and_preserves_frozen_baseline() {
    let mut native_results = vec![];
    for lane in [FloatLane::F32, FloatLane::F64] {
        let mut actual = Chunk::new("actual");
        let mut proposed = actual.clone();
        math::emit_next_after(&mut actual, lane, 37);
        candidate::emit_next_after(&mut proposed, lane, 37);
        let mut prior = Chunk::new("before-ordered-pruning");
        before_ordered::emit_next_after(&mut prior, lane, 37);
        assert_eq!(prior.code, proposed.code, "the frozen copy-elimination stage remains exact");
        assert_eq!(prior.lines, proposed.lines);
        assert!(actual.code.len() < proposed.code.len(), "ordered upward pruning is verified in nextafter_ordered_operands");
        assert_eq!(actual.local_count, proposed.local_count);
        let old = program(true, lane, 1);
        let new = program(false, lane, 1);
        native_results.push(validate_result(&old, &format!("frozen_{lane:?}")));
        native_results.push(validate_result(&new, &format!("candidate_{lane:?}")));
        assert_eq!(old[1].local_count, 6);
        assert_eq!(new[1].local_count, 4);
        assert!(new[1].code.len() < old[1].code.len());
        eprintln!(
            "next_after {lane:?} callee bytes={} -> {} locals={} -> {}",
            old[1].code.len(),
            new[1].code.len(),
            old[1].local_count,
            new[1].local_count
        );
    }
    // Diagnose every lane/version first; preserve strict native acceptance.
    for result in native_results {
        result.unwrap();
    }
}
fn probes() -> Vec<f64> {
    vec![
        0.0,
        -0.0,
        1.0,
        -1.0,
        1.25,
        -1.25,
        2.0,
        -2.0,
        f64::MAX,
        -f64::MAX,
        f64::MIN_POSITIVE,
        -f64::MIN_POSITIVE,
        f64::from_bits(1),
        -f64::from_bits(1),
        f32::from_bits(1) as f64,
        -(f32::from_bits(1) as f64),
        f32::MIN_POSITIVE as f64,
        f32::MAX as f64,
        (f32::MAX as f64) * 2.0,
        f64::INFINITY,
        f64::NEG_INFINITY,
        f64::from_bits(0x7ff8000000001234),
        f64::from_bits(0x7ff0000000000001),
        f64::from_bits(0xfff8000000001234),
    ]
}
#[test]
fn exact_edge_cross_product_and_deterministic_bit_patterns_preserve_all_quirks() {
    let mut pairs: Vec<_> = probes()
        .iter()
        .flat_map(|&x| probes().into_iter().map(move |y| (x, y)))
        .collect();
    let mut bits = 0x4c957f2d8a139b60u64;
    for _ in 0..4096 {
        bits = bits.wrapping_mul(6364136223846793005).wrapping_add(1);
        let x = f64::from_bits(bits);
        bits = bits.wrapping_mul(6364136223846793005).wrapping_add(1);
        let y = f64::from_bits(bits);
        pairs.push((x, y));
    }
    for lane in [FloatLane::F32, FloatLane::F64] {
        let mut loaded = vec![];
        for old in [true, false] {
            let c = program(old, lane, 1);
            validate(&c);
            let mut vm = runtime(Value::F64(0.0), Value::F64(0.0));
            vm.run(c).unwrap();
            let callback = vm.global("next_after_callee").unwrap().clone();
            loaded.push((vm, callback));
        }
        for &(x, y) in &pairs {
            let mut answers = vec![];
            for (vm, callback) in &mut loaded {
                answers.push(exact(
                    &vm.invoke_callback(callback, &[Value::F64(x), Value::F64(y)]),
                ));
            }
            assert_eq!(
                answers[0],
                answers[1],
                "lane={lane:?} x={:016x} y={:016x}",
                x.to_bits(),
                y.to_bits()
            );
        }
    }
}
#[test]
fn strict_input_type_result_abi_and_failure_order_match_baseline() {
    for lane in [FloatLane::F32, FloatLane::F64] {
        let programs = [program(true, lane, 1), program(false, lane, 1)];
        let inputs = vec![
            Value::F64(1.25),
            Value::I32(1),
            Value::I64(1),
            Value::F32(1.0),
            Value::Null,
            Value::Undefined,
            Value::Bool(true),
            Value::String("1".into()),
            Value::bigint_i64(1),
            Value::Object(vybe_runtime::heap::alloc(vybe_runtime::value::Object::new())),
        ];
        for x in &inputs {
            for y in &inputs {
                let mut outcomes = vec![];
                for which in 0..2 {
                    outcomes.push(
                        runtime(x.clone(), y.clone())
                            .run(programs[which].clone())
                            .map(|v| exact(&v))
                            .map_err(error),
                    );
                }
                assert_eq!(outcomes[0], outcomes[1], "lane={lane:?} x={x:?} y={y:?}");
            }
        }
    }
}
#[test]
fn full_capacity_and_existing_offset_metadata_are_preserved() {
    for lane in [FloatLane::F32, FloatLane::F64] {
        let mut before = Chunk::new("largest-before");
        before.local_count = u16::MAX - 4;
        baseline::emit_next_after(&mut before, lane, 37);
        assert_eq!(before.local_count, u16::MAX);
        let mut after = Chunk::new("largest-after");
        after.local_count = u16::MAX - 4;
        candidate::emit_next_after(&mut after, lane, 37);
        assert_eq!(after.local_count, u16::MAX - 2);
        let mut max = Chunk::new("largest-candidate");
        max.local_count = u16::MAX - 2;
        candidate::emit_next_after(&mut max, lane, 37);
        assert_eq!(max.local_count, u16::MAX);
        let mut rejected = Chunk::new("refused");
        rejected.local_count = u16::MAX - 1;
        let code = rejected.code.clone();
        let slots = rejected.local_count;
        assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| candidate::emit_next_after(
                &mut rejected,
                lane,
                37
            )))
            .is_err()
        );
        assert_eq!(rejected.code, code);
        assert_eq!(rejected.local_count, slots);
        let mut chunk = Chunk::new("metadata");
        let marker = chunk.code.len();
        chunk.emit_if_i32(19);
        chunk.emit_i32_const(0, 19);
        chunk.emit_else(19);
        chunk.emit_i32_const(1, 19);
        chunk.emit_end(19);
        chunk.emit_op(Op::DROP, 19);
        let typed = chunk.block_i32_results.clone();
        let lines = chunk.lines.clone();
        let prefix = chunk.code.clone();
        candidate::emit_next_after(&mut chunk, lane, 37);
        assert!(chunk.block_i32_results.contains(&marker));
        assert_eq!(chunk.block_i32_results, typed);
        assert_eq!(&chunk.code[..prefix.len()], prefix);
        assert_eq!(&chunk.lines[..lines.len()], &lines[..]);
        assert!(
            chunk.call_indirect_sigs.is_empty()
                && chunk.call_indirect_canon.is_empty()
                && chunk.stack_switch_handlers.is_empty()
        );
    }
}
#[test]
#[ignore = "interleaved opt-level-0 next_after cold and loaded callee comparison"]
fn benchmark_next_after_branch_preserving_copy_elimination() {
    for lane in [FloatLane::F32, FloatLane::F64] {
        for (x, y) in [
            (1.25, 2.0),
            (-1.25, -2.0),
            (0.0, 1.0),
            (f64::INFINITY, 0.0),
            (f64::NAN, 1.0),
        ] {
            let expected = exact(
                &runtime(Value::F64(x), Value::F64(y))
                    .run(program(true, lane, 1))
                    .unwrap(),
            );
            let modules = [program(true, lane, 4096), program(false, lane, 4096)];
            for c in &modules {
                validate(c);
            }
            for phase in ["cold/load", "loaded/callee"] {
                let mut loaded = vec![];
                if phase == "loaded/callee" {
                    for old in [true, false] {
                        let mut vm = runtime(Value::F64(x), Value::F64(y));
                        vm.run(program(old, lane, 1)).unwrap();
                        let callback = vm.global("next_after_callee").unwrap().clone();
                        loaded.push((vm, callback));
                    }
                }
                let args = [Value::F64(x), Value::F64(y)];
                let mut samples = [vec![], vec![]];
                for round in 0..7 {
                    for which in [round % 2, 1 - round % 2] {
                        let answer;
                        if phase == "cold/load" {
                            let mut vm = runtime(Value::F64(x), Value::F64(y));
                            let c = modules[which].clone();
                            let start = std::time::Instant::now();
                            answer = vm.run(c).unwrap();
                            samples[which].push(start.elapsed());
                        } else {
                            let (vm, callback) = &mut loaded[which];
                            let mut last = Value::Null;
                            let start = std::time::Instant::now();
                            for _ in 0..4096 {
                                last = std::hint::black_box(vm.invoke_callback(callback, &args));
                            }
                            samples[which].push(start.elapsed());
                            answer = last;
                        }
                        assert_eq!(exact(&answer), expected);
                    }
                }
                let raw_us = samples
                    .iter()
                    .map(|sample| {
                        sample
                            .iter()
                            .map(|elapsed| elapsed.as_micros())
                            .collect::<Vec<_>>()
                    })
                    .collect::<Vec<_>>();
                let mut paired_ratios = samples[0]
                    .iter()
                    .zip(&samples[1])
                    .map(|(before, after)| before.as_secs_f64() / after.as_secs_f64())
                    .collect::<Vec<_>>();
                paired_ratios.sort_by(f64::total_cmp);
                for s in &mut samples {
                    s.sort();
                }
                eprintln!(
                    "next_after_copy_elimination {lane:?} {phase} x={x:?} y={y:?} frozen={}us copy_eliminated={}us speedup={:.2}x paired_median={:.2}x raw_frozen_us={:?} raw_candidate_us={:?}",
                    samples[0][3].as_micros(),
                    samples[1][3].as_micros(),
                    samples[0][3].as_secs_f64() / samples[1][3].as_secs_f64(),
                    paired_ratios[3],
                    raw_us[0],
                    raw_us[1]
                );
            }
        }
    }
}
#[test]
fn actual_c_nextafter_and_nextafterf_source_keep_frontend_abi_and_native_results() {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_c::register);
    for (name, lane) in [
        ("nextafter", FloatLane::F64),
        ("nextafterf", FloatLane::F32),
    ] {
        for (a, b) in [("1.25", "2.0"), ("-1.25", "-2.0"), ("0.0", "-1.0")] {
            let source = format!(
                "double calculate(double a, double b) {{ return {name}(a,b); }}\ndouble answer = calculate({a},{b});\n"
            );
            let mut module = vybe_language_c::parse(&source).unwrap();
            module
                .body
                .push(vybe_ast::Statement::new(vybe_ast::StmtKind::Return(Some(
                    vybe_ast::Expression::ident("answer"),
                ))));
            let profile =
                vybe_compiler::profile::parse_profile(vybe_language_c::profile_source()).unwrap();
            let production = vybe_compiler::Compiler::with_profile(profile)
                .compile(&module)
                .unwrap();
            let index = production
                .iter()
                .position(|c| c.name == "calculate")
                .expect("actual parsed Source function");
            let original = &production[index];
            let receiver_params = u8::from(
                original.module_receiver_abi == vybe_runtime::chunk::ReceiverAbi::Parameter,
            );
            assert_eq!(original.arity, 2 + receiver_params);
            // Ordinary Source function declarations use create_function_chunk:
            // arity/takes_receiver are assigned, while param_count stays at
            // its default0 and func_sig is absent (classes.rs1898-1910).
            // The native writer derives this function's params from arity,
            // not that optional declared-Wasm metadata (types.rs1079).
            assert_eq!(original.param_count, 0);
            assert!(original.func_sig.is_none());
            assert_eq!(original.takes_receiver, receiver_params == 1);
            let slot = |name: &str| {
                original
                    .local_names
                    .iter()
                    .find(|n| {
                        n.origin == vybe_runtime::chunk::LocalOrigin::Source && n.name == name
                    })
                    .expect("actual declared parameter slot")
                    .slot
            };
            let (a_slot, b_slot) = (slot("a"), slot("b"));
            assert_ne!(a_slot, b_slot);
            assert_eq!(a_slot, u16::from(receiver_params));
            assert_eq!(b_slot, u16::from(receiver_params) + 1);
            assert!(original.block_i32_results.is_empty());
            assert!(
                original.call_indirect_sigs.is_empty()
                    && original.call_indirect_canon.is_empty()
                    && original.stack_switch_handlers.is_empty()
            );
            assert!(original.user_code_offset.is_none_or(|offset| offset == 0));
            let mut experiment = production.clone();
            let replacement = &mut experiment[index];
            replacement.code.clear();
            replacement.lines.clear();
            replacement.emit_op_u16(Op::LOCAL_GET, a_slot, 37);
            replacement.emit_op_u16(Op::LOCAL_GET, b_slot, 37);
            // Rewrite only the isolated function whose parsed body is exactly
            // this intrinsic call. Preserve actual Source slots/names/imports,
            // module policy and function index; allocate against its full scope.
            candidate::emit_next_after(replacement, lane, 37);
            replacement.emit_op(Op::RETURN, 37);
            assert_eq!(replacement.name, original.name);
            assert_eq!(replacement.arity, original.arity);
            assert_eq!(replacement.param_count, original.param_count);
            assert_eq!(replacement.result_arity, original.result_arity);
            assert_eq!(replacement.func_sig, original.func_sig);
            assert_eq!(replacement.is_method, original.is_method);
            assert_eq!(
                replacement.module_receiver_abi,
                original.module_receiver_abi
            );
            let names = |chunk: &Chunk| {
                chunk
                    .local_names
                    .iter()
                    .map(|name| (name.slot, name.name.clone(), name.origin))
                    .collect::<Vec<_>>()
            };
            assert_eq!(names(replacement), names(original));
            // Actual compiler is now optimized. Retain a genuinely frozen
            // old body as an independent Source baseline, instead of comparing
            // optimized production only against the same candidate twice.
            let mut frozen_source = production.clone();
            let frozen = &mut frozen_source[index];
            frozen.code.clear();
            frozen.lines.clear();
            frozen.emit_op_u16(Op::LOCAL_GET, a_slot, 37);
            frozen.emit_op_u16(Op::LOCAL_GET, b_slot, 37);
            baseline::emit_next_after(frozen, lane, 37);
            frozen.emit_op(Op::RETURN, 37);
            assert_eq!(frozen.arity, original.arity);
            assert_eq!(frozen.param_count, original.param_count);
            assert_eq!(frozen.func_sig, original.func_sig);
            assert_eq!(frozen.module_receiver_abi, original.module_receiver_abi);
            assert_eq!(names(frozen), names(original));
            validate_pair(&frozen_source, &production);
            validate_pair(&frozen_source, &experiment);
            let expected = exact(
                &runtime(Value::Null, Value::Null)
                    .run(frozen_source)
                    .unwrap(),
            );
            assert_eq!(
                exact(&runtime(Value::Null, Value::Null).run(production).unwrap()),
                expected
            );
            assert_eq!(
                exact(&runtime(Value::Null, Value::Null).run(experiment).unwrap()),
                expected
            );
        }
    }
}
