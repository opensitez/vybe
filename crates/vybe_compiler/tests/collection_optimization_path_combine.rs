//! Artifact-only actual baseline/candidate behavior and native/debug gates.
use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::{bundle, globals, paths, platforms};
use vybe_runtime::value::{Function, Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/collection_optimization_path_combine_before.rs"]
mod baseline;
#[path = "fixtures/collection_optimization_path_combine_compact.rs"]
mod candidate;
fn object(o: Object) -> Value {
    Value::Object(vybe_runtime::heap::alloc(o))
}
fn program(mode: u8, n: u8, warm: bool, locals: u16) -> Vec<Chunk> {
    let mut body = Chunk::new("<path-combine-body>");
    body.local_count = locals;
    if warm {
        body.arity = n;
        body.local_count = body.local_count.max(n as u16);
    }
    for i in 0..n {
        if warm {
            body.emit_op_u16(Op::LOCAL_GET, i as u16, 9);
        } else {
            globals::emit_read(&mut body, &format!("arg{i}"), 9);
        }
    }
    match mode {
        0 => baseline::emit_combine(&mut body, n, 9),
        1 => paths::emit_combine(&mut body, n, 9),
        3 => candidate::emit_combine(&mut body, n, 9),
        _ => paths::emit_combine(&mut body, n, 9),
    }
    body.emit_op(Op::RETURN, 9);
    let mut chunks = if warm {
        let mut root = Chunk::new("<path-combine-host-entry>");
        root.emit_op_u16(Op::REF_FUNC, 1, 9);
        root.emit(0, 9);
        root.emit_op(Op::RETURN, 9);
        vec![root, body]
    } else {
        let mut coercion = Chunk::new("<forbidden-object-coercion>");
        let call = coercion.add_import("test:path", "coerce");
        coercion.emit_call(call, 0, 9);
        coercion.emit_op(Op::RETURN, 9);
        vec![body, coercion]
    };
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
fn validate(chunks: &[Chunk]) {
    let bytes = vybe_platform_wasm::write_wasm(chunks);
    if let Err(e) = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
    {
        let out = "/Users/youness/www/html/vybe/tools/primitive_experiments/path_combine";
        std::fs::write(format!("{out}/native-failure.wasm"), bytes).unwrap();
        panic!("actual path native validation: {e}");
    }
}
fn error(mut e: vybe_runtime::VMError) -> String {
    for f in &mut e.call_stack {
        f.offset = 0;
    }
    format!("{e:?}")
}
fn runtime(args: &[Value]) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    // Every cold module includes the coercion probe chunk. Linking resolves its
    // import even when the chunk is never invoked; ordinary cases must register
    // it too. Any actual conversion remains an immediate test failure.
    vm.register_host_fn(
        "test:path",
        "coerce",
        Box::new(|_ctx, _args| panic!("Path.Combine invoked forbidden object coercion")),
    );
    for (i, v) in args.iter().enumerate() {
        vm.set_global(&format!("arg{i}"), v.clone());
    }
    vm
}
fn sample(n: usize, kind: usize) -> Vec<String> {
    (0..n)
        .map(|i| match kind {
            0 => format!("segment{i}"),
            1 => {
                if i % 3 == 0 {
                    "".into()
                } else if i % 3 == 1 {
                    "tail/".into()
                } else {
                    "next".into()
                }
            }
            2 => {
                if i + 1 == n {
                    "/final/root".into()
                } else {
                    format!("discarded{i}")
                }
            }
            3 => ["é/", "😀", "é:drive", "/😀root/", "后续", "😀:relative"][i % 6].into(),
            _ => [
                "C:\\dir\\",
                "child",
                "\\server\\share",
                "leaf/",
                "",
                "E:relative",
            ][i % 6]
                .into(),
        })
        .collect()
}
fn reference(args: &[String]) -> String {
    let mut out = String::new();
    for s in args {
        if s.is_empty() {
            continue;
        }
        let colon = s
            .find(':')
            .is_some_and(|byte| s[..byte].encode_utf16().count() == 1);
        if s.starts_with('/') || s.starts_with('\\') || colon || out.is_empty() {
            out = s.clone();
        } else {
            if !out.ends_with('/') {
                out.push('/');
            }
            out.push_str(s);
        }
    }
    out
}
#[test]
fn actual_production_matches_checked_compact_and_preserves_common_emission() {
    for n in [0, 1, 2, 4, 8, 15, 16, 32, 255] {
        let old = program(0, n, false, 0);
        let expected = program(3, n, false, 0);
        let actual = program(2, n, false, 0);
        assert_eq!(expected.len(), actual.len());
        for (a, b) in expected.iter().zip(&actual) {
            assert_eq!(a.code, b.code);
            assert_eq!(a.lines, b.lines);
            assert_eq!(a.local_count, b.local_count);
            assert_eq!(a.constants, b.constants);
            assert_eq!(format!("{:?}", a.imports), format!("{:?}", b.imports));
            assert_eq!(a.block_i32_results, b.block_i32_results);
        }
        if n < 16 {
            assert_eq!(old[0].code, actual[0].code);
            assert_eq!(old[0].local_count, actual[0].local_count);
        } else {
            assert_ne!(old[0].code, actual[0].code);
            assert!(actual[0].code.len() < old[0].code.len());
        }
    }
}
#[test]
fn native_and_vm_rooted_reset_empty_drive_backslash_and_utf16_are_exact() {
    for n in [0, 1, 2, 4, 8, 15, 16, 17, 32, 255] {
        for kind in 0..5 {
            let text = sample(n, kind);
            let expected = reference(&text);
            let args: Vec<_> = text
                .iter()
                .map(|s| Value::String(s.as_str().into()))
                .collect();
            for mode in [0, 1] {
                let chunks = program(mode, n as u8, false, 0);
                validate(&chunks);
                let mut vm = runtime(&args);
                assert_eq!(
                    vm.run(chunks).unwrap(),
                    Value::String(expected.as_str().into()),
                    "mode={mode} n={n} kind={kind}"
                );
                for (i, v) in args.iter().enumerate() {
                    assert_eq!(vm.global(&format!("arg{i}")), Some(v));
                }
            }
        }
    }
}
fn coercible() -> Value {
    let mut o = Object::new();
    o.properties.insert(
        "valueOf".into(),
        object(Object {
            descriptor: None,
            kind: ObjectKind::Function(Function {
                name: Some("forbidden".into()),
                arity: 0,
                chunk_index: 1,
                upvalues: vec![],
            }),
            ..Object::new()
        }),
    );
    o.properties
        .insert("toString".into(), o.properties["valueOf"].clone());
    object(o)
}
#[test]
fn strict_type_failure_order_does_not_invoke_object_coercion_even_before_root_reset() {
    for n in [16, 32, 255] {
        for bad in [0, n / 2, n - 2] {
            for value in [Value::Null, Value::Undefined, Value::I32(42), coercible()] {
                let mut args: Vec<_> = sample(n, 2)
                    .into_iter()
                    .map(|s| Value::String(s.into()))
                    .collect();
                args[bad] = value;
                let outcomes = [0, 1].map(|mode| {
                    let calls = Arc::new(Mutex::new(0));
                    let observed = calls.clone();
                    let mut vm = runtime(&args);
                    vm.register_host_fn(
                        "test:path",
                        "coerce",
                        Box::new(move |_ctx, _args| {
                            *observed.lock().unwrap() += 1;
                            Value::String("coerced".into())
                        }),
                    );
                    let result = vm.run(program(mode, n as u8, false, 0)).map_err(error);
                    assert!(result.is_err());
                    assert_eq!(
                        *calls.lock().unwrap(),
                        0,
                        "strict String proof forbids coercion"
                    );
                    result
                });
                assert_eq!(outcomes[0], outcomes[1], "n={n} bad={bad}");
            }
        }
    }
}
#[test]
fn exact_legacy_local_budget_largest_fit_and_preflight_refusal() {
    for n in [16u8, 32, 255] {
        let old = program(0, n, false, 0);
        let new = program(1, n, false, 0);
        assert_eq!(new[0].local_count, candidate::LONG_LOCAL_BUDGET);
        assert!(new[0].local_count < old[0].local_count);
        assert!(new[0].code.len() < old[0].code.len());
        let needed = old[0].local_count;
        let old_fit = program(0, n, false, u16::MAX - needed);
        let new_fit = program(1, n, false, u16::MAX - needed);
        assert_eq!(old_fit[0].local_count, u16::MAX);
        assert!(new_fit[0].local_count < u16::MAX);
        let largest_new = program(1, n, false, u16::MAX - candidate::LONG_LOCAL_BUDGET);
        assert_eq!(largest_new[0].local_count, u16::MAX);
        let args: Vec<_> = sample(n as usize, 2)
            .into_iter()
            .map(|s| Value::String(s.into()))
            .collect();
        let expected = reference(&sample(n as usize, 2));
        assert_eq!(
            runtime(&args).run(largest_new).unwrap(),
            Value::String(expected.into())
        );
        // Native MAX50000 cap remains unchanged; near65535 is an actual VM gate.
        let mut refused = Chunk::new("<one-past-path-budget>");
        refused.local_count = u16::MAX - candidate::LONG_LOCAL_BUDGET + 1;
        let before = format!("{refused:?}");
        assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| paths::emit_combine(
                &mut refused,
                n,
                9
            )))
            .is_err()
        );
        assert_eq!(format!("{refused:?}"), before);
        eprintln!(
            "path n={n} locals={} bytecode {}→{} native {}→{}",
            needed,
            old[0].code.len(),
            new[0].code.len(),
            vybe_platform_wasm::write_wasm(&old).len(),
            vybe_platform_wasm::write_wasm(&new).len()
        );
    }
}

#[test]
#[ignore = "bounded compact opt0 benchmark; one call per mode in seven raw pairs"]
fn compact_path_bounded_complete_call_benchmark() {
    use std::time::Instant;
    for (n, rooted) in [(16u8, false), (64, false), (255, false), (255, true)] {
        let programs = [program(0, n, true, 0), program(1, n, true, 0)];
        for p in &programs {
            validate(p);
        }
        // Immutable source construction, expected result and compiled programs
        // are outside all timing. Module loading has its own cold diagnostic.
        let mut text = vec!["x".repeat(4096); n as usize];
        if rooted {
            *text.last_mut().unwrap() = "/final-root".into();
        }
        let expected = Value::String(reference(&text).into());
        let args: Vec<_> = text.into_iter().map(|s| Value::String(s.into())).collect();
        let mut samples = [Vec::new(), Vec::new()];
        let mut cold = [Vec::new(), Vec::new()];
        for batch in 0..7 {
            let mut pair = [0u128; 2];
            for mode in if batch % 2 == 0 { [0, 1] } else { [1, 0] } {
                let mut vm = VM::new();
                platforms::register_platforms_all(&mut vm);
                let chunks = programs[mode].clone();
                let begin = Instant::now();
                let f = vm.run(chunks).unwrap();
                cold[mode].push(begin.elapsed().as_micros());
                let begin = Instant::now();
                let result = vm.invoke_callback(&f, &args);
                pair[mode] = begin.elapsed().as_micros();
                // Do not include full-output string equality in call timing.
                assert_eq!(result, expected, "n={n} rooted={rooted} mode={mode}");
                samples[mode].push(pair[mode]);
            }
            eprintln!(
                "compact path raw n={n} segment=4096 rooted={rooted} calls=1 batch={batch} old={}us compact={}us gain={:.4}x",
                pair[0],
                pair[1],
                pair[0] as f64 / pair[1] as f64
            );
        }
        for v in &mut samples {
            v.sort_unstable();
        }
        for v in &mut cold {
            v.sort_unstable();
        }
        eprintln!(
            "compact path median n={n} segment=4096 rooted={rooted} calls=1 old={}us compact={}us gain={:.4}x cold={}→{}us bytecode {}→{} locals {}→{}",
            samples[0][3],
            samples[1][3],
            samples[0][3] as f64 / samples[1][3] as f64,
            cold[0][3],
            cold[1][3],
            programs[0][1].code.len(),
            programs[1][1].code.len(),
            programs[0][1].local_count,
            programs[1][1].local_count
        );
    }
}

fn identity_function(index: usize, arity: u8) -> Value {
    object(Object {
        descriptor: None,
            kind: ObjectKind::Function(Function {
            name: Some("identity-probe".into()),
            arity,
            chunk_index: index,
            upvalues: vec![],
        }),
        ..Object::new()
    })
}
fn identity_program(mode: u8, n: u8) -> Vec<Chunk> {
    let mut root = Chunk::new("<path-identity-host-entry>");
    root.add_global_import(vybe_runtime::chunk::HOST_GLOBALS_MODULE, "identity_input");
    root.emit_op_u16(Op::REF_FUNC, 1, 9);
    root.emit(0, 9);
    root.emit_op(Op::RETURN, 9);
    let mut body = Chunk::new("<path-identity-body>");
    body.arity = n;
    body.local_count = n as u16;
    for i in 0..n {
        body.emit_op_u16(Op::LOCAL_GET, i as u16, 9);
    }
    if mode == 0 {
        baseline::emit_combine(&mut body, n, 9);
    } else if mode == 3 {
        candidate::emit_combine(&mut body, n, 9);
    } else {
        paths::emit_combine(&mut body, n, 9);
    }
    body.emit_op(Op::RETURN, 9);
    let mut eq = Chunk::new("<actual-ref-eq>");
    eq.arity = 2;
    eq.local_count = 2;
    eq.emit_op_u16(Op::LOCAL_GET, 0, 9);
    eq.emit_op_u16(Op::LOCAL_GET, 1, 9);
    eq.emit_op(Op::REF_EQ, 9);
    eq.emit_op(Op::RETURN, 9);
    let mut chunks = vec![
        root,
        body,
        eq,
        Chunk::new("<raw-finalized-input>"),
        Chunk::new("<raw-finalized-empty>"),
    ];
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    let input = chunks[0]
        .globals
        .iter()
        .position(|v| v == "identity_input")
        .unwrap();
    let empty_name =
        vybe_runtime::chunk::imported_global_key(vybe_runtime::chunk::STRING_CONSTANTS_MODULE, "");
    let empty = chunks[0]
        .globals
        .iter()
        .position(|v| v == &empty_name)
        .unwrap();
    // Operands are ALREADY module-global indices. Do not call emit_read after
    // normalization: it encodes constant-pool keys and would test wrong values.
    for (chunk, slot) in [(3, input), (4, empty)] {
        chunks[chunk].emit_op_u32(Op::GLOBAL_GET, slot as u32, 9);
        chunks[chunk].emit_op(Op::RETURN, 9);
    }
    chunks
}
#[test]
fn path_combine_preserves_single_string_and_empty_constant_ref_identity() {
    for n in [16u8, 32, 255] {
        for mode in [0, 1, 3] {
            let chunks = identity_program(mode, n);
            validate(&chunks);
            let input = Value::String(format!("/identity-é😀-{n}").into());
            let mut vm = VM::new();
            platforms::register_platforms_all(&mut vm);
            vm.set_global("identity_input", input.clone());
            let combine = vm.run(chunks).unwrap();
            let input = vm.invoke_callback(&identity_function(3, 0), &[]);
            assert_eq!(
                input,
                Value::String(format!("/identity-é😀-{n}").into()),
                "raw finalized input must be actual host value"
            );
            let eq = identity_function(2, 2);
            let same = |vm: &mut VM, a: &Value, b: &Value| {
                vm.invoke_callback(&eq, &[a.clone(), b.clone()])
            };
            for (pos, discarded) in [(0, false), (n as usize / 2, false), (n as usize / 2, true)] {
                let mut args = vec![Value::String("".into()); n as usize];
                args[pos] = input.clone();
                if discarded {
                    args[0] = Value::String("discarded-before-root".into());
                }
                let result = vm.invoke_callback(&combine, &args);
                assert_eq!(result, input);
                assert_eq!(
                    same(&mut vm, &result, &input),
                    Value::I32(1),
                    "single original String identity mode={mode} n={n} pos={pos} discarded={discarded}"
                );
            }
            let empty = vm.invoke_callback(&identity_function(4, 0), &[]);
            assert_eq!(empty, Value::String("".into()));
            let args = vec![Value::String("".into()); n as usize];
            let a = vm.invoke_callback(&combine, &args);
            let b = vm.invoke_callback(&combine, &args);
            assert_eq!(
                same(&mut vm, &a, &b),
                Value::I32(1),
                "repeated all-empty calls must retain constant mode={mode} n={n}"
            );
            assert_eq!(
                same(&mut vm, &a, &empty),
                Value::I32(1),
                "actual captured empty constant mode={mode} n={n}"
            );
            let mut args = vec![Value::String("".into()); n as usize];
            args[0] = input.clone();
            args[n as usize / 2] = Value::String("tail".into());
            let a = vm.invoke_callback(&combine, &args);
            let b = vm.invoke_callback(&combine, &args);
            assert_eq!(a, b);
            assert_ne!(a, input);
            assert_eq!(
                same(&mut vm, &a, &input),
                Value::I32(0),
                "multi-piece output must remain fresh"
            );
            assert_eq!(
                same(&mut vm, &a, &b),
                Value::I32(0),
                "separate multi-piece calls must remain fresh"
            );
        }
    }
}
