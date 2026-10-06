//! Frozen versus integrated producer-proven IsNumeric factory.
use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::{bundle, callable, convert, globals, platforms};
use vybe_runtime::chunk::{ConstExpr, GlobalInit};
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/numeric_optimization_isnumeric_baseline.rs"]
mod baseline;
#[path = "fixtures/numeric_optimization_isnumeric_candidate.rs"]
mod candidate;
fn factory(old: bool, root: &mut Chunk) -> Chunk {
    if old {
        baseline::build_is_numeric(root)
    } else {
        candidate::build_is_numeric(root)
    }
}
fn runtime(input: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("input_value", input.clone());
    vm.set_global("input", input);
    vm
}
fn program(old: bool, calls: usize) -> Vec<Chunk> {
    let mut root = Chunk::new("<isnumeric-experiment>");
    let helper = factory(old, &mut root);
    root.global_inits.push(GlobalInit {
        name: "__vybe_isnumeric".into(),
        init: ConstExpr::RefFunc(1),
    });
    for i in 0..calls {
        globals::emit_read(&mut root, "__vybe_isnumeric", 37);
        globals::emit_read(&mut root, "input", 37);
        callable::emit_direct_invoke_chunk(&mut root, 1, 37);
        if i + 1 != calls {
            root.emit_op(Op::DROP, 37);
        }
    }
    root.emit_op(Op::RETURN, 37);
    let mut chunks = vec![root, helper];
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn validate(chunks: &[Chunk], label: &str) {
    let bytes = vybe_platform_wasm::write_wasm(chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap_or_else(|e| panic!("{label}: {e}"));
}
fn cases() -> Vec<(Value, i32)> {
    vec![
        (Value::I32(i32::MIN), 1),
        (Value::I32(i32::MAX), 1),
        (Value::I64(i64::MIN), 1),
        (Value::I64(i64::MAX), 1),
        (Value::F32(-0.0), 1),
        (Value::F32(f32::INFINITY), 1),
        (Value::F32(f32::from_bits(0x7f800123)), 1),
        (Value::F64(0.0), 1),
        (Value::F64(-0.0), 1),
        (Value::F64(f64::INFINITY), 1),
        (Value::F64(f64::NEG_INFINITY), 1),
        (Value::F64(f64::from_bits(0x7ff0000000001234)), 1),
        (Value::F64(f64::from_bits(0xfff8000000001234)), 1),
        (Value::Bool(true), 0),
        (Value::Bool(false), 0),
        (Value::Null, 0),
        (Value::TypedNull(0), 0),
        (Value::Undefined, 0),
        (Value::WeakRef(std::sync::Weak::new()), 0),
        (Value::bigint_i64(12), 0),
        (Value::Symbol("number".into()), 0),
        (Value::V128([0; 16]), 0),
        (Value::String("".into()), 0),
        (Value::String("   ".into()), 0),
        (Value::String("12".into()), 1),
        (Value::String("12x".into()), 1),
        (Value::String(" -0.5e2tail".into()), 1),
        (Value::String("NaN".into()), 0),
        (Value::String("∞".into()), 0),
        (Value::String("１２".into()), 0),
        (Value::Object(vybe_runtime::heap::alloc(Object::new())), 0),
    ]
}
#[test]
fn exact_tags_nan_prefix_and_i32_result_match_frozen_factory() {
    let programs = [program(true, 1), program(false, 1)];
    for (i, p) in programs.iter().enumerate() {
        validate(p, &format!("standalone {i}"));
    }
    for (input, expected) in cases() {
        for (i, p) in programs.iter().enumerate() {
            assert_eq!(
                runtime(input.clone()).run(p.clone()).unwrap(),
                Value::I32(expected),
                "variant{i}, {input:?}"
            );
        }
    }
}
fn semantic_code(
    c: &Chunk,
    imports: &[vybe_runtime::chunk::Import],
) -> Vec<(
    Op,
    Vec<u8>,
    Option<vybe_runtime::chunk::Import>,
    Option<String>,
)> {
    let mut out = vec![];
    let mut pc = 0;
    while pc < c.code.len() {
        let op = Op::decode(c.read_u16(pc), c.read_u16(pc + 2)).unwrap();
        let end = pc + 4 + op.operand_format().size_in(&c.code, pc + 4);
        let mut operands = c.code[pc + 4..end].to_vec();
        let import = if op == Op::CALL {
            let import = imports[c.read_u16(pc + 4) as usize].clone();
            operands[..2].fill(0);
            Some(import)
        } else {
            None
        };
        let global = if op == Op::GLOBAL_GET || op == Op::GLOBAL_SET {
            let index = c.read_u32(pc + 4) as usize;
            let name = if !c.globals.is_empty() {
                c.globals[index].clone()
            } else {
                let Value::String(name) = &c.constants[index] else {
                    panic!("named frozen global")
                };
                name.to_string()
            };
            operands.fill(0);
            Some(name)
        } else {
            None
        };
        out.push((op, operands, import, global));
        pc = end;
    }
    out
}
#[test]
fn actual_production_is_exact_candidate_and_reduces_frozen_calls() {
    let mut root = Chunk::new("imports");
    let actual = convert::build_is_numeric(&mut root);
    let old = factory(true, &mut Chunk::new("imports"));
    let new = factory(false, &mut Chunk::new("imports"));
    assert_eq!(actual.code, new.code);
    assert_eq!(actual.lines, new.lines);
    assert_eq!(actual.imports, new.imports);
    assert_eq!(actual.constants, new.constants);
    assert_eq!(old.arity, new.arity);
    assert_eq!(old.local_count, new.local_count);
    let count = |c: &Chunk| {
        semantic_code(c, &c.imports)
            .into_iter()
            .filter(|x| x.0 == Op::CALL)
            .count()
    };
    assert_eq!(count(&old), 11);
    assert_eq!(count(&new), 3);
    assert!(new.code.len() < old.code.len());
    // emit_dup preserves the parsed F64 in its separately allocated scratch.
    assert_eq!(old.local_count, 3);
    assert_eq!(new.local_count, 3);
    for (name, c) in [("old", old), ("new", new)] {
        eprintln!(
            "isnumeric {name} bytes={} locals={} static_calls={}",
            c.code.len(),
            c.local_count,
            count(&c)
        );
    }
}
#[test]
fn rejected_objects_do_not_invoke_coercion_methods_or_mutate_identity() {
    for old in [true, false] {
        let mut vm = runtime(Value::Null);
        let observed = Arc::new(Mutex::new(0usize));
        let count = observed.clone();
        vm.register_host_fn(
            "experiment:isnumeric",
            "coerce",
            Box::new(move |_, _| {
                *count.lock().unwrap() += 1;
                Value::String("12".into())
            }),
        );
        let mut root = Chunk::new("<callback-owner>");
        root.global_inits.push(GlobalInit {
            name: "isnumeric_coercion_method".into(),
            init: ConstExpr::RefFunc(1),
        });
        globals::emit_read(&mut root, "isnumeric_coercion_method", 9);
        root.emit_op(Op::RETURN, 9);
        let mut callback = Chunk::new("<must-not-coerce>");
        callback.arity = 1;
        callback.local_count = 1;
        let call = callback.add_import("experiment:isnumeric", "coerce");
        callback.emit_call(call, 0, 9);
        callback.emit_op(Op::RETURN, 9);
        let mut owner = vec![root, callback];
        bundle::finalize_with_runtime_helpers(&mut owner);
        globals::declare_free_globals(&mut owner);
        globals::normalize_global_table(&mut owner);
        validate(&owner, "genuine callback owner");
        let method = vm.run(owner).unwrap();
        let mut object = Object::new();
        object.properties.insert("valueOf".into(), method.clone());
        object.properties.insert("toString".into(), method.clone());
        let object = vybe_runtime::heap::alloc(object);
        let before = object.lock().unwrap().properties.clone();
        let mut proxy = Object::new();
        proxy
            .properties
            .insert("__vybe_proxy_target".into(), method.clone());
        proxy.properties.insert("valueOf".into(), method.clone());
        proxy.properties.insert("toString".into(), method.clone());
        let proxy = vybe_runtime::heap::alloc(proxy);
        let proxy_before = proxy.lock().unwrap().properties.clone();
        let Value::Object(function) = &method else {
            panic!("genuine user callable")
        };
        let function_before = function.lock().unwrap().properties.clone();
        for input in [
            Value::Object(object.clone()),
            Value::WeakRef(Arc::downgrade(&object)),
            method.clone(),
            Value::Object(proxy.clone()),
            Value::WeakRef(Arc::downgrade(&proxy)),
        ] {
            vm.set_global("input", input.clone());
            assert_eq!(vm.run(program(old, 1)).unwrap(), Value::I32(0));
            assert_eq!(vm.global("input"), Some(&input));
            assert_eq!(*observed.lock().unwrap(), 0);
            assert_eq!(object.lock().unwrap().properties, before);
            assert_eq!(proxy.lock().unwrap().properties, proxy_before);
            assert_eq!(function.lock().unwrap().properties, function_before);
        }
        // A genuine writable callable alias must still be invoked. This is
        // independent of the primitive's implementation and has no native
        // substitute: it executes the actual function installed above.
        vm.set_global("__vybe_isnumeric", method.clone());
        assert_eq!(vm.run(program(old, 1)).unwrap(), Value::String("12".into()));
        assert_eq!(*observed.lock().unwrap(), 1);
        assert_eq!(vm.global("__vybe_isnumeric"), Some(&method));
    }
}
#[test]
fn public_assignment_errors_and_source_diagnostics_remain_identical() {
    let mut errors = vec![];
    for old in [true, false] {
        let mut vm = runtime(Value::String("12x".into()));
        vm.set_global("__vybe_isnumeric", Value::I32(42));
        let mut e = vm.run(program(old, 1)).unwrap_err();
        for frame in &mut e.call_stack {
            frame.offset = 0;
        }
        errors.push(e.to_string());
        assert_eq!(vm.global("__vybe_isnumeric"), Some(&Value::I32(42)));
    }
    assert_eq!(errors[0], errors[1]);
}
fn source_pair(kind: usize) -> [Vec<Chunk>; 2] {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_vb::register);
    let source=match kind {
        0=>"Dim answer = IsNumeric(input_value)\n".to_string(),
        1=>"Function Classify(v As Object) As Object\nReturn IsNumeric(v)\nEnd Function\nDim answer = Classify(input_value)\n".to_string(),
        _=>"Dim classify = Function(v As Object) IsNumeric(v)\nDim answer = classify(input_value)\n".to_string(),
    };
    let mut module = vybe_language_vb::parse(&source).unwrap();
    module
        .body
        .push(vybe_ast::Statement::new(vybe_ast::StmtKind::Return(Some(
            vybe_ast::Expression::ident("answer"),
        ))));
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_vb::profile_source()).unwrap();
    let old = vybe_compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap();
    let index = old
        .iter()
        .position(|c| c.name == "__stdlib_isnumeric")
        .expect("actual VB IsNumeric route");
    let isolated = factory(false, &mut Chunk::new("imports"));
    assert_eq!(
        semantic_code(&old[index], &old[0].imports),
        semantic_code(&isolated, &isolated.imports)
    );
    assert_eq!(old[index].arity, isolated.arity);
    assert_eq!(old[index].local_count, isolated.local_count);
    let frozen = factory(true, &mut Chunk::new("imports"));
    let mut before = old.clone();
    before[index].code = frozen.code;
    before[index].lines = frozen.lines;
    before[index].imports = frozen.imports;
    before[index].constants = frozen.constants;
    before[index].global_imports.extend(frozen.global_imports);
    // The frozen body adds named string constants absent from production.
    // Restore existing module globalidx operands to their original named
    // constant-pool form, then rebuild ONE genuine module global index space.
    // Instruction widths and every offset-bearing metadata key stay fixed.
    before[index].globals = Arc::new(vec![]);
    for chunk in &mut before {
        if chunk.globals.is_empty() {
            continue;
        }
        let mut pc = 0;
        while pc < chunk.code.len() {
            let op = Op::decode(chunk.read_u16(pc), chunk.read_u16(pc + 2)).unwrap();
            let size = 4 + op.operand_format().size_in(&chunk.code, pc + 4);
            if op == Op::GLOBAL_GET || op == Op::GLOBAL_SET {
                let name = chunk.globals[chunk.read_u32(pc + 4) as usize].clone();
                let ci = chunk.intern_string_constant(&name);
                chunk.code[pc + 4..pc + 8].copy_from_slice(&ci.to_be_bytes());
            }
            pc += size;
        }
        chunk.globals = Arc::new(vec![]);
    }
    globals::declare_free_globals(&mut before);
    globals::normalize_global_table(&mut before);
    assert_eq!(
        semantic_code(&before[index], &before[index].imports),
        semantic_code(
            &factory(true, &mut Chunk::new("imports")),
            &factory(true, &mut Chunk::new("imports")).imports
        )
    );
    [before, old]
}
#[test]
fn actual_vb_root_function_lambda_validate_and_match_frozen_inputs() {
    for kind in 0..3 {
        let pair = source_pair(kind);
        for (i, p) in pair.iter().enumerate() {
            validate(p, &format!("VB kind{kind} variant{i}"));
        }
        for (input, expected) in cases() {
            for (i, p) in pair.iter().enumerate() {
                assert_eq!(
                    runtime(input.clone()).run(p.clone()).unwrap(),
                    Value::I32(expected),
                    "VB kind{kind} variant{i} input{input:?}"
                );
            }
        }
    }
}
#[test]
#[ignore = "calibrated >=30ms paired debug cold+loaded practical IsNumeric workloads"]
fn debug_practical_isnumeric_calibrated_cold_and_loaded() {
    use std::time::{Duration, Instant};
    const MIN: Duration = Duration::from_millis(30);
    for input in [
        Value::F64(12.5),
        Value::String("12.5tail".into()),
        Value::String("invalid".into()),
        Value::Bool(false),
    ] {
        for calls in [16usize, 64, 256] {
            let programs = [program(true, calls), program(false, calls)];
            for p in &programs {
                validate(p, "calibrated preparation");
            }
            let expected = runtime(input.clone()).run(program(true, 1)).unwrap();
            let mut retained = vec![];
            for old in [true, false] {
                let mut vm = runtime(input.clone());
                assert_eq!(vm.run(program(old, 1)).unwrap(), expected);
                let callback = vm.global("__vybe_isnumeric").unwrap().clone();
                retained.push((vm, callback));
            }
            let args = [input.clone()];
            for phase in ["cold", "loaded"] {
                let measure =
                    |which: usize, batches: usize, retained: &mut Vec<(VM, Value)>| -> Duration {
                        if phase == "cold" {
                            // Aggregate real load/run intervals. Registering host
                            // platforms and cloning inputs happens OUTSIDE each
                            // interval; avoid retaining hundreds of whole VMs.
                            let mut elapsed = Duration::ZERO;
                            for _ in 0..batches {
                                let mut vm = runtime(input.clone());
                                let chunks = programs[which].clone();
                                let start = Instant::now();
                                let answer = vm.run(chunks).unwrap();
                                elapsed += start.elapsed();
                                assert_eq!(answer, expected);
                            }
                            elapsed
                        } else {
                            let (vm, callback) = &mut retained[which];
                            let mut answer = Value::Null;
                            let start = Instant::now();
                            for _ in 0..batches {
                                for _ in 0..calls {
                                    answer =
                                        std::hint::black_box(vm.invoke_callback(callback, &args));
                                }
                            }
                            let elapsed = start.elapsed();
                            assert_eq!(answer, expected);
                            elapsed
                        }
                    };
                let mut batches = 1usize;
                while (0..2)
                    .map(|i| measure(i, batches, &mut retained))
                    .min()
                    .unwrap()
                    < MIN
                {
                    batches = batches.checked_mul(2).unwrap();
                    assert!(batches <= 1 << 20, "calibration bound");
                }
                let (times, ratios) = loop {
                    let mut times = [vec![], vec![]];
                    for round in 0..7 {
                        for which in [round % 2, 1 - round % 2] {
                            times[which].push(measure(which, batches, &mut retained).as_nanos());
                        }
                    }
                    if times.iter().flatten().any(|t| *t < MIN.as_nanos()) {
                        batches = batches.checked_mul(2).unwrap();
                        assert!(batches <= 1 << 20);
                        continue;
                    }
                    let ratios: Vec<f64> = (0..7)
                        .map(|r| times[0][r] as f64 / times[1][r] as f64)
                        .collect();
                    break (times, ratios);
                };
                eprintln!("isnumeric input={input:?} calls/module={calls} phase={phase} batches={batches} raw_ns={times:?} raw_paired_ratios={ratios:?}");
                let mut times = times;
                let mut ratios = ratios;
                for t in &mut times {
                    t.sort();
                }
                ratios.sort_by(f64::total_cmp);
                eprintln!("isnumeric frozen={}ns candidate={}ns paired_median={:.3}x measured_calls={} shortest_sample={}ns",times[0][3],times[1][3],ratios[3],calls*batches,times[0][0].min(times[1][0]));
            }
        }
    }
}
