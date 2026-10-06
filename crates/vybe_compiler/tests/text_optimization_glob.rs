use vybe_compiler::primitives::{bundle, globals, platforms, strings};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/text_optimization_glob_equality_baseline.rs"]
mod baseline;
#[path = "fixtures/text_optimization_glob_baseline.rs"]
mod buggy_baseline;
fn program_with_retained(
    old: bool,
    matching: Option<bool>,
    calls: usize,
    retained: bool,
) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<glob-private-comparisons>")];
    let current = if retained {
        chunks[0].emit_op_u16(Op::REF_FUNC, 1, 0);
        chunks[0].emit(0, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        chunks.push(Chunk::new("<retained-glob-translator>"));
        1
    } else {
        0
    };
    for index in 0..calls {
        if matching.is_some() {
            globals::emit_read(&mut chunks[current], "name", 0);
        }
        globals::emit_read(&mut chunks[current], "pattern", 0);
        if let Some(fold_case) = matching {
            let options = strings::GlobOptions { fold_case };
            if old {
                baseline::emit_glob_match(&mut chunks, current, options, 0);
            } else {
                strings::emit_glob_match(&mut chunks, current, options, 0);
            }
        } else if old {
            baseline::emit_glob_to_regex(&mut chunks, current, 0);
        } else {
            strings::emit_glob_to_regex(&mut chunks, current, 0);
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
fn program(old: bool, matching: Option<bool>, calls: usize) -> Vec<Chunk> {
    program_with_retained(old, matching, calls, false)
}
fn runtime() -> VM {
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
fn run(vm: &mut VM, chunks: &[Chunk], pattern: Value, name: &str) -> Result<Value, String> {
    vm.set_global("pattern", pattern);
    vm.set_global("name", Value::String(name.into()));
    vm.run(chunks.to_vec()).map_err(|mut error| {
        for frame in &mut error.call_stack {
            frame.offset = 0;
        }
        error.to_string()
    })
}
#[test]
fn original_numeric_string_predicates_capture_the_bug_separately() {
    let mut buggy = vec![Chunk::new("<frozen-glob-string-predicate-bug>")];
    globals::emit_read(&mut buggy[0], "pattern", 0);
    buggy_baseline::emit_glob_to_regex(&mut buggy, 0, 0);
    buggy[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut buggy);
    globals::declare_free_globals(&mut buggy);
    globals::normalize_global_table(&mut buggy);
    validate(&buggy);
    let fixed = program(false, None, 1);
    validate(&fixed);
    let mut old_vm = runtime();
    let mut new_vm = runtime();
    assert_eq!(
        run(&mut old_vm, &buggy, Value::String("abc".into()), "").unwrap(),
        Value::String("^.*.*.*$".into())
    );
    assert_eq!(
        run(&mut new_vm, &fixed, Value::String("abc".into()), "").unwrap(),
        Value::String("^abc$".into())
    );
    assert_eq!(
        run(&mut new_vm, &fixed, Value::String("[!a-z]".into()), "").unwrap(),
        Value::String("^[^a-z]$".into())
    );
    assert_eq!(
        run(&mut new_vm, &fixed, Value::String("a*?".into()), "").unwrap(),
        Value::String("^a.*.$".into())
    );
    assert_eq!(
        run(&mut old_vm, &buggy, Value::String("[A-Z]".into()), "").unwrap(),
        Value::String("^.*.*.*.*.*$".into())
    );
    assert_eq!(
        run(&mut new_vm, &fixed, Value::String("[A-Z]".into()), "").unwrap(),
        Value::String("^[A-Z]$".into())
    );
    // These are interpreted VM outputs. Whole-Wasm validation above does
    // not establish execution of the emitted module in a native engine.
}

#[test]
fn frozen_translation_preserves_ascii_brackets_utf16_and_strict_errors() {
    let old = program(true, None, 1);
    let new = program(false, None, 1);
    validate(&old);
    validate(&new);
    let mut before = runtime();
    let mut after = runtime();
    let mut patterns = (0u8..=127)
        .map(|c| format!("x{}y", c as char))
        .collect::<Vec<_>>();
    patterns.extend(
        [
            "",
            "*",
            "?",
            "[]",
            "[!]",
            "[!a]",
            "[a-z]",
            "[A-Z]",
            "[abc",
            "[",
            "[!",
            "{{}}",
            "a\\*b",
            "é😀",
            "[😀]",
            "e\u{301}\0?",
            "\\.()+^${}|",
        ]
        .map(str::to_owned),
    );
    for pattern in patterns {
        let before = run(&mut before, &old, Value::String(pattern.clone().into()), "");
        let after = run(&mut after, &new, Value::String(pattern.clone().into()), "");
        assert_eq!(before, after, "pattern={pattern:?}");
        if pattern == "é😀" {
            assert_eq!(after.unwrap(), Value::String("^é��$".into()));
        }
    }
    for input in [
        Value::Null,
        Value::Undefined,
        Value::I32(7),
        Value::F64(7.0),
        Value::Bool(true),
        Value::Object(vybe_runtime::heap::alloc(vybe_runtime::value::Object::new())),
    ] {
        let old_result = run(&mut before, &old, input.clone(), "");
        let new_result = run(&mut after, &new, input, "");
        assert_eq!(old_result, new_result);
        assert!(new_result.unwrap_err().contains("not a string"));
    }
}
#[test]
fn strict_length_failure_precedes_object_string_coercion() {
    use std::sync::{
        Arc,
        atomic::{AtomicUsize, Ordering},
    };
    for old in [true, false] {
        let chunks = program(old, None, 1);
        validate(&chunks);
        let mut vm = runtime();
        let calls = Arc::new(AtomicUsize::new(0));
        let observed = calls.clone();
        vm.register_host_fn(
            "test:glob",
            "stringify",
            Box::new(move |_, _| {
                observed.fetch_add(1, Ordering::SeqCst);
                Value::String("a*".into())
            }),
        );
        let index = vm
            .resolve_host_function_index("test:glob", "stringify")
            .unwrap();
        let mut function = vybe_runtime::value::Object::new();
        function.kind = vybe_runtime::value::ObjectKind::HostFunction(index);
        let mut object = vybe_runtime::value::Object::new();
        object.properties.insert(
            "toString".into(),
            Value::Object(vybe_runtime::heap::alloc(function)),
        );
        let error = run(
            &mut vm,
            &chunks,
            Value::Object(vybe_runtime::heap::alloc(object)),
            "",
        )
        .unwrap_err();
        assert!(error.contains("wasm:js-string.length"));
        assert_eq!(calls.load(Ordering::SeqCst), 0);
    }
}

#[test]
fn glob_matching_preserves_exact_folded_classes_and_regex_errors() {
    for folded in [false, true] {
        let old = program(true, Some(folded), 1);
        let new = program(false, Some(folded), 1);
        validate(&old);
        validate(&new);
        let mut before = runtime();
        let mut after = runtime();
        for (pattern, name) in [
            ("gr[ae]y", "grey"),
            ("[!a]", "B"),
            ("[A-Z]", "a"),
            ("a*b?", "axxbz"),
            ("*", "😀"),
            ("?", "😀"),
            ("é😀", "é😀"),
            ("[", "x"),
            ("[!]", "!"),
        ] {
            let old_result = run(&mut before, &old, Value::String(pattern.into()), name);
            let new_result = run(&mut after, &new, Value::String(pattern.into()), name);
            assert_eq!(
                old_result, new_result,
                "pattern={pattern:?} name={name:?} folded={folded}"
            );
            if pattern == "[A-Z]" {
                assert_eq!(new_result.unwrap(), Value::Bool(folded));
            }
        }
    }
}
fn probes(chunks: &[Chunk]) -> usize {
    chunks
        .iter()
        .map(|chunk| {
            let mut offset = 0;
            let mut count = 0;
            while offset < chunk.code.len() {
                let op = Op::from_code_at(&chunk.code, offset);
                if op == Op::CALL {
                    let index = u16::from_be_bytes([chunk.code[offset + 4], chunk.code[offset + 5]])
                        as usize;
                    if chunk.imports[index].module == "wasm:js-number"
                        && chunk.imports[index].name == "test"
                    {
                        count += 1;
                    }
                }
                offset += 4 + op.operand_format().size_in(&chunk.code, offset + 4);
            }
            count
        })
        .sum()
}
#[test]
fn repeated_translation_reduces_real_emitted_probes_and_native_size() {
    for calls in [1, 16] {
        let old = program(true, None, calls);
        let new = program(false, None, calls);
        let old_native = validate(&old);
        let new_native = validate(&new);
        assert!(probes(&new) < probes(&old));
        let old_code: usize = old.iter().map(|c| c.code.len()).sum();
        let new_code: usize = new.iter().map(|c| c.code.len()).sum();
        assert!(new_code < old_code);
        assert!(new_native < old_native);
        println!(
            "glob calls={calls} code={old_code}->{new_code} wasm={old_native}->{new_native} static_number_probes={}->{}",
            probes(&old),
            probes(&new)
        );
    }
}
#[test]
#[ignore = "interleaved opt-level-0 cold and loaded glob diagnostic"]
fn benchmark_glob_private_comparisons() {
    use std::time::Instant;
    for (name, pattern, calls) in [
        ("literal96", "abcxyz".repeat(16), 1),
        ("literal2048", "a".repeat(2048), 1),
        ("wildcard2048", "a*?b".repeat(512), 1),
        ("class2048", "[!a-z]xx".repeat(256), 1),
        ("repeat16", "abc*?xyz".repeat(12), 16),
    ] {
        let old = program(true, None, calls);
        let new = program(false, None, calls);
        validate(&old);
        validate(&new);
        let mut old_vm = runtime();
        let mut new_vm = runtime();
        let expected = run(&mut old_vm, &old, Value::String(pattern.clone().into()), "").unwrap();
        let mut times = [vec![], vec![]];
        for sample in 0..7 {
            for variant in if sample % 2 == 0 { [0, 1] } else { [1, 0] } {
                let (vm, chunks) = if variant == 0 {
                    (&mut old_vm, &old)
                } else {
                    (&mut new_vm, &new)
                };
                vm.set_global("pattern", Value::String(pattern.clone().into()));
                let prepared = chunks.clone();
                let start = Instant::now();
                let actual = vm.run(prepared).unwrap();
                let elapsed = start.elapsed().as_nanos();
                assert_eq!(actual, expected);
                times[variant].push(elapsed);
            }
        }
        for list in &mut times {
            list.sort();
        }
        println!(
            "glob {name} cold old={}ns new={}ns gain={:.2}x",
            times[0][3],
            times[1][3],
            times[0][3] as f64 / times[1][3] as f64
        );
        let old_loaded = program_with_retained(true, None, calls, true);
        let new_loaded = program_with_retained(false, None, calls, true);
        validate(&old_loaded);
        validate(&new_loaded);
        let old_ref = old_vm.run(old_loaded).unwrap();
        let new_ref = new_vm.run(new_loaded).unwrap();
        let mut loaded = [vec![], vec![]];
        for sample in 0..7 {
            for variant in if sample % 2 == 0 { [0, 1] } else { [1, 0] } {
                let (vm, callback) = if variant == 0 {
                    (&mut old_vm, &old_ref)
                } else {
                    (&mut new_vm, &new_ref)
                };
                let start = Instant::now();
                let actual = vm.invoke_callback(callback, &[]);
                let elapsed = start.elapsed().as_nanos();
                assert_eq!(actual, expected);
                loaded[variant].push(elapsed);
            }
        }
        for list in &mut loaded {
            list.sort();
        }
        println!(
            "glob {name} loaded old={}ns new={}ns gain={:.2}x",
            loaded[0][3],
            loaded[1][3],
            loaded[0][3] as f64 / loaded[1][3] as f64
        );
    }
}
