//! Frozen versus integrated producer-proven radix prefix concat.
use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::{bundle, callable, convert, globals, platforms};
use vybe_runtime::chunk::{ConstExpr, GlobalInit};
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, Value, VM};
#[path = "fixtures/numeric_optimization_pyradix_baseline.rs"]
mod baseline;
#[path = "fixtures/numeric_optimization_pyradix_candidate.rs"]
mod candidate;
include!("fixtures/numeric_optimization_pyradix_semantics.rs");
fn spec(radix: i32) -> (&'static str, &'static str, &'static str) {
    match radix {
        16 => ("__stdlib_pyhex", "__vybe_pyhex", "0x"),
        8 => ("__stdlib_pyoct", "__vybe_pyoct", "0o"),
        2 => ("__stdlib_pybin", "__vybe_pybin", "0b"),
        _ => unreachable!(),
    }
}
fn factory(old: bool, root: &mut Chunk, radix: i32) -> Chunk {
    let (name, _, prefix) = spec(radix);
    if old {
        baseline::build_pyradix(root, name, prefix, radix)
    } else {
        candidate::build_pyradix(root, name, prefix, radix)
    }
}
fn program(old: bool, radix: i32, calls: usize) -> Vec<Chunk> {
    let mut root = Chunk::new("<radix-experiment>");
    let helper = factory(old, &mut root, radix);
    root.global_inits.push(GlobalInit {
        name: spec(radix).1.into(),
        init: ConstExpr::RefFunc(1),
    });
    for i in 0..calls {
        globals::emit_read(&mut root, spec(radix).1, 37);
        globals::emit_read(&mut root, "input_value", 37);
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
fn runtime(input: Value) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("input_value", input);
    vm
}
fn validate(chunks: &[Chunk], label: &str) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap_or_else(|e| panic!("{label}: {e}"));
}
fn inputs() -> Vec<Value> {
    vec![
        Value::I32(0),
        Value::I32(-1),
        Value::I32(i32::MIN),
        Value::I32(i32::MAX),
        Value::I64(i64::MIN),
        Value::I64(i64::MAX),
        Value::F64(0.),
        Value::F64(-0.),
        Value::F64(1.25),
        Value::F64(-1.25),
        Value::F64(f64::MAX),
        Value::F64(f64::INFINITY),
        Value::F64(f64::from_bits(0x7ff0000000001234)),
        Value::F32(12.),
        Value::Null,
        Value::TypedNull(0),
        Value::Undefined,
        Value::Bool(true),
        Value::String("12".into()),
        Value::bigint_i64(12),
        Value::Symbol("12".into()),
        Value::WeakRef(std::sync::Weak::new()),
        Value::V128([0; 16]),
        Value::Object(vybe_runtime::heap::alloc(Object::new())),
    ]
}
#[test]
fn exact_frozen_factories_static_dispatch_and_local_budgets() {
    for radix in [2, 8, 16] {
        let (name, _, prefix) = spec(radix);
        let actual = convert::build_pyradix(&mut Chunk::new("imports"), name, prefix, radix);
        let old = factory(true, &mut Chunk::new("imports"), radix);
        let new = factory(false, &mut Chunk::new("imports"), radix);
        assert_eq!(actual.code, new.code);
        assert_eq!(actual.lines, new.lines);
        assert_eq!(actual.imports, new.imports);
        assert_eq!(actual.constants, new.constants);
        assert_eq!(old.local_count, 3);
        assert_eq!(new.local_count, 1);
        assert_eq!(old.arity, new.arity);
        let count = |c: &Chunk| {
            semantic_code(c, &c.imports)
                .into_iter()
                .filter(|r| r.0 == Op::CALL)
                .count()
        };
        assert_eq!(count(&new), 2);
        assert!(count(&old) > count(&new));
        assert!(new.code.len() < old.code.len());
        eprintln!("radix{radix} old_bytes={} new_bytes={} old_static_calls={} new_static_calls=2 old_locals=3 new_locals=1",old.code.len(),new.code.len(),count(&old));
    }
}
#[test]
fn exact_original_tags_fraction_nonnumber_and_nan_behavior_native_valid() {
    for radix in [2, 8, 16] {
        let pair = [program(true, radix, 1), program(false, radix, 1)];
        for (i, p) in pair.iter().enumerate() {
            validate(p, &format!("radix{radix} variant{i}"));
        }
        for input in inputs() {
            let old = runtime(input.clone()).run(pair[0].clone()).unwrap();
            let new = runtime(input.clone()).run(pair[1].clone()).unwrap();
            assert!(matches!(old, Value::String(_)));
            assert_eq!(old, new, "radix{radix} input{input:?}");
        }
        assert_eq!(
            runtime(Value::I32(-1)).run(pair[1].clone()).unwrap(),
            Value::String(format!("{}-1", spec(radix).2).into()),
            "preserve existing prefix placement"
        );
    }
}
#[test]
fn objects_do_not_gain_coercion_callbacks_and_alias_errors_stay_exact() {
    for radix in [2, 8, 16] {
        let mut errors = vec![];
        for old in [true, false] {
            let mut vm = runtime(Value::Null);
            let count = Arc::new(Mutex::new(0usize));
            let observed = count.clone();
            vm.register_host_fn(
                "experiment:radix",
                "coerce",
                Box::new(move |_, _| {
                    *observed.lock().unwrap() += 1;
                    Value::F64(12.)
                }),
            );
            let mut root = Chunk::new("<method-owner>");
            root.global_inits.push(GlobalInit {
                name: "radix_coercer".into(),
                init: ConstExpr::RefFunc(1),
            });
            globals::emit_read(&mut root, "radix_coercer", 37);
            root.emit_op(Op::RETURN, 37);
            let mut method = Chunk::new("<coercer>");
            method.arity = 1;
            method.local_count = 1;
            let call = method.add_import("experiment:radix", "coerce");
            method.emit_call(call, 0, 37);
            method.emit_op(Op::RETURN, 37);
            let mut owner = vec![root, method];
            bundle::finalize_with_runtime_helpers(&mut owner);
            globals::declare_free_globals(&mut owner);
            globals::normalize_global_table(&mut owner);
            validate(&owner, "actual method owner");
            let method = vm.run(owner).unwrap();
            let mut obj = Object::new();
            obj.properties.insert("valueOf".into(), method.clone());
            obj.properties.insert("toString".into(), method.clone());
            let obj = vybe_runtime::heap::alloc(obj);
            let before = obj.lock().unwrap().properties.clone();
            vm.set_global("input_value", Value::Object(obj.clone()));
            assert_eq!(
                vm.run(program(old, radix, 1)).unwrap(),
                Value::String(format!("{}0", spec(radix).2).into())
            );
            assert_eq!(*count.lock().unwrap(), 0);
            assert_eq!(obj.lock().unwrap().properties, before);
            vm.set_global(spec(radix).1, method.clone());
            assert_eq!(vm.run(program(old, radix, 1)).unwrap(), Value::F64(12.));
            assert_eq!(*count.lock().unwrap(), 1);
            assert_eq!(vm.global(spec(radix).1), Some(&method));
            vm.set_global(spec(radix).1, Value::I32(42));
            let mut e = vm.run(program(old, radix, 1)).unwrap_err();
            for f in &mut e.call_stack {
                f.offset = 0;
            }
            errors.push(e.to_string());
            assert_eq!(vm.global(spec(radix).1), Some(&Value::I32(42)));
        }
        assert_eq!(errors[0], errors[1]);
    }
}
#[test]
fn real_python_hex_oct_bin_root_function_lambda_frozen_candidate_native() {
    static ONCE: std::sync::Once = std::sync::Once::new();
    ONCE.call_once(vybe_language_python::register);
    for (radix, spelling) in [(2, "bin"), (8, "oct"), (16, "hex")] {
        for kind in 0..3 {
            let source = match kind {
                0 => format!("answer={spelling}(input_value)\n"),
                1 => format!(
                    "def calculate(v):\n    return {spelling}(v)\nanswer=calculate(input_value)\n"
                ),
                _ => format!("calculate=lambda v: {spelling}(v)\nanswer=calculate(input_value)\n"),
            };
            let mut module = vybe_language_python::parse(&source).unwrap();
            module
                .body
                .push(vybe_ast::Statement::new(vybe_ast::StmtKind::Return(Some(
                    vybe_ast::Expression::ident("answer"),
                ))));
            let profile =
                vybe_compiler::profile::parse_profile(vybe_language_python::profile_source())
                    .unwrap();
            let old = vybe_compiler::Compiler::with_profile(profile)
                .compile(&module)
                .unwrap();
            let index = old
                .iter()
                .position(|c| c.name == spec(radix).0)
                .expect("actual Python radix helper route");
            let frozen = factory(false, &mut Chunk::new("imports"), radix);
            assert_eq!(
                semantic_code(&old[index], &old[0].imports),
                semantic_code(&frozen, &frozen.imports)
            );
            assert_eq!(old[index].local_count, 1);
            let mut new = old.clone();
            let mut c = factory(true, &mut Chunk::new("imports"), radix);
            // Factory emits named constant-pool globals; real Source already
            // carries module globalidx. Relocate the SAME existing prefix
            // against the genuine compiled global table before substitution.
            let mut pc = 0;
            while pc < c.code.len() {
                let op = Op::decode(c.read_u16(pc), c.read_u16(pc + 2)).unwrap();
                let size = 4 + op.operand_format().size_in(&c.code, pc + 4);
                if op == Op::GLOBAL_GET {
                    let Value::String(name) = &c.constants[c.read_u32(pc + 4) as usize] else {
                        panic!("named prefix global")
                    };
                    let index = old[index]
                        .globals
                        .iter()
                        .position(|n| n == name.as_ref())
                        .expect("prefix already in actual Source globals");
                    c.code[pc + 4..pc + 8].copy_from_slice(&(index as u32).to_be_bytes());
                }
                pc += size;
            }
            new[index].code = c.code;
            new[index].lines = c.lines;
            new[index].imports = c.imports;
            new[index].constants = c.constants;
            new[index].local_count = c.local_count;
            let unrelocated = factory(true, &mut Chunk::new("imports"), radix);
            assert_eq!(
                semantic_code(&new[index], &new[index].imports),
                semantic_code(&unrelocated, &unrelocated.imports)
            );
            for (i, p) in [&old, &new].iter().enumerate() {
                validate(p, &format!("Python radix{radix} kind{kind} variant{i}"));
            }
            for input in [Value::I32(31), Value::F64(-1.25), Value::F64(-0.)] {
                let a = runtime(input.clone()).run(new.clone()).unwrap_or_else(|e| {
                    panic!("frozen Python{spelling} kind{kind} input{input:?}: {e}")
                });
                let b = runtime(input.clone()).run(old.clone()).unwrap_or_else(|e| {
                    panic!("candidate Python{spelling} kind{kind} input{input:?}: {e}")
                });
                assert_eq!(a, b, "actual Python{spelling} kind{kind} input{input:?}");
            }
        }
    }
}

#[test]
#[ignore = "bounded interleaved practical radix debug benchmark; cold uncalibrated explicitly"]
fn debug_radix_practical_cold_and_loaded() {
    use std::time::{Duration, Instant};
    for radix in [2, 8, 16] {
        for calls in [16usize, 64, 256] {
            let programs = [program(true, radix, calls), program(false, radix, calls)];
            for p in &programs {
                validate(p, "benchmark preparation");
            }
            for input in [Value::I32(31), Value::I32(-257)] {
                let expected = runtime(input.clone()).run(program(true, radix, 1)).unwrap();
                let mut cold = [vec![], vec![]];
                let mut retained = vec![];
                for old in [true, false] {
                    let mut vm = runtime(input.clone());
                    assert_eq!(vm.run(program(old, radix, 1)).unwrap(), expected);
                    let callback = vm.global(spec(radix).1).unwrap().clone();
                    retained.push((vm, callback));
                }
                // Exactly14 cold run intervals per case: no repeated whole-VM
                // calibration. Setup and module clones remain outside measured work.
                for round in 0..7 {
                    for i in [round % 2, 1 - round % 2] {
                        let mut vm = runtime(input.clone());
                        let chunks = programs[i].clone();
                        let start = Instant::now();
                        let answer = vm.run(chunks).unwrap();
                        cold[i].push(start.elapsed().as_nanos());
                        assert_eq!(answer, expected);
                    }
                }
                let args = [input.clone()];
                let measure =
                    |i: usize, batches: usize, retained: &mut Vec<(VM, Value)>| -> Duration {
                        let (vm, callback) = &mut retained[i];
                        let mut answer = Value::Null;
                        let start = Instant::now();
                        for _ in 0..batches {
                            for _ in 0..calls {
                                answer = std::hint::black_box(vm.invoke_callback(callback, &args));
                            }
                        }
                        let elapsed = start.elapsed();
                        assert_eq!(answer, expected);
                        elapsed
                    };
                let mut batches = 1usize;
                while (0..2)
                    .map(|i| measure(i, batches, &mut retained))
                    .min()
                    .unwrap()
                    < Duration::from_millis(20)
                {
                    batches *= 2;
                    assert!(batches <= 65536, "bounded loaded calibration");
                }
                let mut loaded = [vec![], vec![]];
                for round in 0..7 {
                    for i in [round % 2, 1 - round % 2] {
                        loaded[i].push(measure(i, batches, &mut retained).as_nanos());
                    }
                }
                for (phase, mut times) in [("cold", cold), ("loaded", loaded)] {
                    let mut ratios: Vec<f64> = (0..7)
                        .map(|r| times[0][r] as f64 / times[1][r] as f64)
                        .collect();
                    eprintln!("radix{radix} input{input:?} phase={phase} calls/module={calls} loaded_batches={batches} raw_ns={times:?} raw_paired_ratios={ratios:?}");
                    for t in &mut times {
                        t.sort();
                    }
                    ratios.sort_by(f64::total_cmp);
                    eprintln!("radix medians frozen={}ns candidate={}ns paired_ratio={:.3}x shortest={}ns cold_note=short_uncalibrated_no_uniform_gain_claim",times[0][3],times[1][3],ratios[3],times[0][0].min(times[1][0]));
                }
            }
        }
    }
}
