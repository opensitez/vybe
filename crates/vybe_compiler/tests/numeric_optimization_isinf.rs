use std::sync::{Arc, Mutex};
// Isolated rejected isinf experiment: production retains the frozen factory.
// Checked Boolean removes two host calls, but Infinity timings regressed.
use vybe_compiler::primitives::{bundle, callable, globals, platforms};
use vybe_runtime::chunk::{ConstExpr, GlobalInit};
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, VM, ValType, Value};
#[path = "fixtures/numeric_optimization_isinf_baseline.rs"]
mod baseline;
#[path = "fixtures/numeric_optimization_isinf_candidate.rs"]
mod candidate;
fn factory(old: bool, root: &mut Chunk) -> Chunk {
    if old {
        baseline::build_isinf(root)
    } else {
        candidate::build_isinf(root)
    }
}
fn program(old: bool, calls: usize) -> Vec<Chunk> {
    let mut root = Chunk::new("<numeric-isinf>");
    let helper = factory(old, &mut root);
    root.global_inits.push(GlobalInit {
        name: "__vybe_isinf".into(),
        init: ConstExpr::RefFunc(1),
    });
    for n in 0..calls {
        globals::emit_read(&mut root, "__vybe_isinf", 37);
        globals::emit_read(&mut root, "input", 37);
        callable::emit_direct_invoke_chunk(&mut root, 1, 37);
        if n + 1 != calls {
            root.emit_op(Op::DROP, 37)
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
    vm.set_global("input", input);
    vm
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}
fn assert_same_factory_semantics(
    actual: &Chunk,
    module_imports: &[vybe_runtime::chunk::Import],
    frozen: &Chunk,
) {
    assert_eq!(actual.name, frozen.name);
    assert_eq!(actual.arity, frozen.arity);
    assert_eq!(actual.local_count, frozen.local_count);
    assert_eq!(actual.module_receiver_abi, frozen.module_receiver_abi);
    let decode = |chunk: &Chunk, imports: &[vybe_runtime::chunk::Import]| {
        let mut instructions = vec![];
        let mut offset = 0;
        while offset < chunk.code.len() {
            let op = Op::decode(chunk.read_u16(offset), chunk.read_u16(offset + 2)).unwrap();
            let end = offset + 4 + op.operand_format().size_in(&chunk.code, offset + 4);
            assert!(end <= chunk.code.len());
            let mut operands = chunk.code[offset + 4..end].to_vec();
            let import = if op == Op::CALL {
                let import = imports.get(chunk.read_u16(offset + 4) as usize)
                    .expect("CALL resolves in its compiler-finalized import index space");
                let signature = vybe_runtime::declared_host_signature(&import.module, &import.name)
                    .map(|(sig, _)| (sig.params, sig.results));
                // Normalize only the relocated function index. Keep argc and
                // every other operand byte exact; compare import identity and
                // its declared parameter/result ABI separately.
                operands[0] = 0;
                operands[1] = 0;
                Some((import.clone(), signature))
            } else {
                None
            };
            instructions.push((op, operands, import));
            offset = end;
        }
        instructions
    };
    assert_eq!(decode(actual, module_imports), decode(frozen, &frozen.imports));
}
#[test]
fn declared_boolean_producers_preserve_strict_inputs_and_exact_i32_result_abi() {
    let _vm = runtime(Value::Null);
    for name in ["isFinite", "isNaN"] {
        let (sig, _) = vybe_runtime::declared_host_signature("ecma:number", name)
            .expect("actual registered signature");
        assert_eq!(sig.results, vec![ValType::Bool]);
        assert_eq!(sig.params.len(), 1)
    }
    let cases = vec![
        (Value::I32(0), 0),
        (Value::I32(i32::MIN), 0),
        (Value::F64(-0.0), 0),
        (Value::F64(f64::MAX), 0),
        (Value::F64(f64::MIN_POSITIVE), 0),
        (Value::F64(f64::INFINITY), 1),
        (Value::F64(f64::NEG_INFINITY), 1),
        (Value::F64(f64::from_bits(0x7ff8000000001234)), 0),
        (Value::F64(f64::from_bits(0xfff0000000000011)), 0),
        // Preserve the helper's existing strict-nonnumber result, even though
        // its old Python-labelled doccomment suggested a broader numeric rule.
        (Value::Null, 1),
        (Value::Undefined, 1),
        (Value::Bool(false), 1),
        (Value::Bool(true), 1),
        (Value::String("inf".into()), 1),
        (Value::String("12".into()), 1),
        (Value::I64(12), 1),
        (Value::bigint_i64(12), 1),
        (Value::Object(vybe_runtime::heap::alloc(Object::new())), 1),
    ];
    let programs = [program(true, 1), program(false, 1)];
    for chunks in &programs {
        validate(chunks)
    }
    for (input, expected) in cases {
        for old in [true, false] {
            assert_eq!(
                runtime(input.clone())
                    .run(programs[usize::from(!old)].clone())
                    .unwrap(),
                Value::I32(expected),
                "old={old} input={input:?}"
            )
        }
    }
}
#[test]
fn predicate_order_argument_identity_and_throw_diagnostics_are_unchanged() {
    for throw_at in [None, Some("isFinite"), Some("isNaN")] {
        let mut outcomes = vec![];
        for old in [true, false] {
            let input = Value::Object(vybe_runtime::heap::alloc(Object::new()));
            let mut vm = runtime(input.clone());
            let log = Arc::new(Mutex::new(vec![]));
            for name in ["isFinite", "isNaN"] {
                let log = log.clone();
                let input = input.clone();
                vm.register_host_fn(
                    "ecma:number",
                    name,
                    Box::new(move |ctx, args| {
                        let (Value::Object(expected), Value::Object(actual)) = (&input, &args[0])
                        else {
                            panic!("object identity")
                        };
                        assert!(Arc::ptr_eq(expected, actual));
                        log.lock().unwrap().push(name);
                        if throw_at == Some(name) {
                            ctx.throw_value(Value::String(format!("strict {name} failure").into()))
                        }
                        Value::Bool(false)
                    }),
                );
            }
            let mut result = vm.run(program(old, 1));
            if let Err(error) = &mut result {
                for frame in &mut error.call_stack {
                    frame.offset = 0
                }
            }
            let result = result.map_err(|e| e.to_string());
            let observed = log.lock().unwrap().clone();
            assert_eq!(
                observed,
                if throw_at == Some("isFinite") {
                    vec!["isFinite"]
                } else {
                    vec!["isFinite", "isNaN"]
                }
            );
            outcomes.push(result);
        }
        assert_eq!(outcomes[0], outcomes[1]);
        if throw_at.is_none() {
            assert_eq!(outcomes[0], Ok(Value::I32(1)))
        } else {
            assert!(outcomes[0].is_err())
        }
    }
}
#[test]
fn helper_removes_only_redundant_truthiness_calls_and_validates_native() {
    let mut root = Chunk::new("imports");
    let old = factory(true, &mut root);
    let new = factory(false, &mut root);
    assert_eq!(old.arity, new.arity);
    assert_eq!(old.local_count, new.local_count);
    assert!(
        old.imports
            .iter()
            .any(|i| i.module == "ecma:boolean" && i.name == "toBoolean")
    );
    assert!(
        !new.imports
            .iter()
            .any(|i| i.module == "ecma:boolean" && i.name == "toBoolean")
    );
    let count = |c: &Chunk| {
        let mut calls = 0;
        let mut at = 0;
        while at < c.code.len() {
            let op = Op::decode(c.read_u16(at), c.read_u16(at + 2)).unwrap();
            if op == Op::CALL {
                calls += 1
            }
            at += 4 + op.operand_format().size_in(&c.code, at + 4);
        }
        calls
    };
    assert_eq!(count(&old), 6);
    assert_eq!(count(&new), 4);
    assert!(new.code.len() < old.code.len());
    for old in [true, false] {
        validate(&program(old, 256))
    }
}
#[test]
fn real_python_existing_public_helper_root_function_lambda_and_alias_contracts() {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_python::register);
    for kind in 0..3 {
        let call = "__vybe_isinf(float('inf'))";
        let body = match kind {
            0 => format!("answer={call}\n"),
            1 => format!("def calculate():\n    return {call}\nanswer=calculate()\n"),
            _ => format!("calculate=lambda: {call}\nanswer=calculate()\n"),
        };
        let mut module = vybe_language_python::parse(&body).unwrap();
        module
            .body
            .push(vybe_ast::Statement::new(vybe_ast::StmtKind::Return(Some(
                vybe_ast::Expression::ident("answer"),
            ))));
        let profile =
            vybe_compiler::profile::parse_profile(vybe_language_python::profile_source()).unwrap();
        let production = vybe_compiler::Compiler::with_profile(profile)
            .compile(&module)
            .unwrap();
        let index = production
            .iter()
            .position(|c| c.name == "__stdlib_isinf")
            .expect("real public helper dependency");
        let mut frozen = baseline::build_isinf(&mut Chunk::new("imports"));
        let isolated = candidate::build_isinf(&mut Chunk::new("imports"));
        assert_eq!(frozen.module_receiver_abi, isolated.module_receiver_abi);
        assert_eq!(frozen.module_receiver_abi, vybe_runtime::chunk::ReceiverAbi::Ambient);
        // Standalone factories default to Ambient; actual Source compilation
        // attaches its frontend receiver policy to every linked helper. Apply
        // that same metadata to the frozen bytecode before comparing it.
        frozen.module_receiver_abi = production[index].module_receiver_abi;
        assert_same_factory_semantics(&production[index], &production[0].imports, &frozen);
        assert_eq!(isolated.name, production[index].name);
        assert_eq!(isolated.arity, production[index].arity);
        assert_eq!(isolated.local_count, production[index].local_count);
        let mut experiment = production.clone();
        // Keep the actual linked factory's metadata and index. Only this
        // isolated fixture's bytecode, line table and import operands change.
        experiment[index].code = isolated.code;
        experiment[index].lines = isolated.lines;
        experiment[index].imports = isolated.imports;
        assert_eq!(experiment[index].name, production[index].name);
        assert_eq!(experiment[index].arity, production[index].arity);
        assert_eq!(experiment[index].local_count, production[index].local_count);
        assert_eq!(experiment[index].module_receiver_abi, production[index].module_receiver_abi);
        for chunks in [production, experiment] {
            validate(&chunks);
            assert_eq!(runtime(Value::Null).run(chunks).unwrap(), Value::I32(1));
        }
    }
    // This pre-existing public binding remains writable in both versions.
    let mut errors = vec![];
    for old in [true, false] {
        let mut vm = runtime(Value::F64(f64::INFINITY));
        vm.set_global("__vybe_isinf", Value::I32(42));
        let mut error = vm.run(program(old, 1)).unwrap_err();
        for frame in &mut error.call_stack {
            frame.offset = 0;
        }
        errors.push(error.to_string());
        assert_eq!(vm.global("__vybe_isinf"), Some(&Value::I32(42)));
    }
    assert_eq!(errors[0], errors[1]);
}
#[test]
#[ignore = "interleaved opt-level-0 frozen exact factory/public caller comparison"]
fn debug_repeated_known_boolean_isinf_comparison() {
    for input in [
        Value::F64(1.25),
        Value::F64(f64::INFINITY),
        Value::F64(f64::NAN),
    ] {
        let programs = [program(true, 4096), program(false, 4096)];
        let mut samples = [vec![], vec![]];
        let mut answers = [Value::Null, Value::Null];
        for round in 0..7 {
            for which in [round % 2, 1 - round % 2] {
                let mut vm = runtime(input.clone());
                let chunks = programs[which].clone();
                let start = std::time::Instant::now();
                answers[which] = vm.run(chunks).unwrap();
                samples[which].push(start.elapsed())
            }
        }
        assert_eq!(answers[0], answers[1]);
        for s in &mut samples {
            s.sort();
        }
        eprintln!(
            "isinf4096 cold/load input={input:?} frozen={}us checked_bool={}us speedup={:.2}x",
            samples[0][3].as_micros(),
            samples[1][3].as_micros(),
            samples[0][3].as_secs_f64() / samples[1][3].as_secs_f64()
        );

        // Isolate the actual loaded callee from module loading, global-table
        // setup, and the emitted public caller. Both callbacks are obtained
        // from the real published binding after executing the actual module.
        let expected = Value::I32(i32::from(matches!(&input, Value::F64(x) if x.is_infinite())));
        let args = [input.clone()];
        let mut loaded = Vec::new();
        for old in [true, false] {
            let chunks = program(old, 1);
            validate(&chunks);
            let mut vm = runtime(input.clone());
            assert_eq!(vm.run(chunks).unwrap(), expected);
            let callback = vm.global("__vybe_isinf").expect("published actual helper").clone();
            assert_eq!(vm.invoke_callback(&callback, &args), expected);
            loaded.push((vm, callback));
        }
        let mut loaded_samples = [vec![], vec![]];
        for round in 0..7 {
            for which in [round % 2, 1 - round % 2] {
                let (vm, callback) = &mut loaded[which];
                let mut answer = Value::Null;
                let start = std::time::Instant::now();
                for _ in 0..4096 {
                    answer = std::hint::black_box(vm.invoke_callback(callback, &args));
                }
                let elapsed = start.elapsed();
                assert_eq!(answer, expected);
                loaded_samples[which].push(elapsed);
            }
        }
        for samples in &mut loaded_samples {
            samples.sort();
        }
        eprintln!(
            "isinf4096 loaded/callee input={input:?} frozen={}us checked_bool={}us speedup={:.2}x",
            loaded_samples[0][3].as_micros(),
            loaded_samples[1][3].as_micros(),
            loaded_samples[0][3].as_secs_f64() / loaded_samples[1][3].as_secs_f64()
        );
    }
}
