use vybe_compiler::primitives::{bundle, globals, platforms, string_similarity};
use vybe_runtime::{Chunk, Op, VM, Value};

#[path = "fixtures/uncached_levenshtein.rs"]
mod baseline;

fn program(old: bool, calls: usize) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<levenshtein>")];
    for call in 0..calls {
        globals::emit_read(&mut chunks[0], "first", 0);
        globals::emit_read(&mut chunks[0], "second", 0);
        if old {
            baseline::emit_levenshtein(&mut chunks, 0, 2, 0);
        } else {
            string_similarity::emit_levenshtein(&mut chunks, 0, 2, 0);
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

fn vm(a: &str, b: &str) -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.set_global("first", Value::String(a.into()));
    vm.set_global("second", Value::String(b.into()));
    vm
}

#[test]
fn character_caching_preserves_existing_distance_and_unicode_contracts() {
    for (a, b, expected) in [
        ("", "", 0),
        ("kitten", "sitting", 3),
        ("a", "", 1),
        ("", "abc", 3),
        ("same", "same", 0),
        ("ab", "ba", 2),
    ] {
        for old in [true, false] {
            assert_eq!(
                vm(a, b).run(program(old, 1)).unwrap().as_f64(),
                expected as f64
            );
        }
    }
    let strings = [
        "", "a", "é", "😀", "A😀B", "é😀é", "\u{fffd}", "e\u{301}", "\0x",
    ];
    for a in strings {
        for b in strings {
            let old = vm(a, b).run(program(true, 1)).unwrap();
            let new = vm(a, b).run(program(false, 1)).unwrap();
            assert_eq!(old, new, "{a:?} / {b:?}");
        }
    }
    for old in [true, false] {
        let chunks = program(old, 2);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
            .unwrap();
    }
}

#[test]
fn caching_reduces_character_queries_to_linear_count() {
    use std::sync::{
        Arc,
        atomic::{AtomicUsize, Ordering},
    };
    for (a, b) in [("abcd", "uvwxyz"), ("", "abcdef"), ("abcdef", "")] {
        for old in [true, false] {
            let count = Arc::new(AtomicUsize::new(0));
            let observed = count.clone();
            let mut runtime = vm(a, b);
            runtime.register_host_fn(
                "ecma:string",
                "charAt",
                Box::new(move |_ctx, args| {
                    observed.fetch_add(1, Ordering::Relaxed);
                    let text = args[0].as_str();
                    let index = args[1].as_i32() as usize;
                    let units: Vec<_> = text.encode_utf16().collect();
                    Value::String(
                        units
                            .get(index)
                            .map(|u| String::from_utf16_lossy(&[*u]))
                            .unwrap_or_default()
                            .into(),
                    )
                }),
            );
            runtime.run(program(old, 1)).unwrap();
            let expected = if a.is_empty() || b.is_empty() {
                0
            } else if old {
                2 * a.len() * b.len()
            } else {
                a.len() + b.len()
            };
            assert_eq!(count.load(Ordering::Relaxed), expected);
        }
    }
}

#[test]
#[ignore = "interleaved debug distance execution benchmark"]
fn cached_distance_execution_benchmark() {
    for (kind, a, b) in [
        ("ascii", "abcd".repeat(32), "abce".repeat(32)),
        ("unicode", "é世".repeat(64), "é界".repeat(64)),
        ("unicode-long", "é世".repeat(128), "é界".repeat(128)),
    ] {
        let programs = [program(true, 1), program(false, 1)];
        let mut times = [vec![], vec![]];
        let mut results = [Value::Null, Value::Null];
        for round in 0..5 {
            for which in [round % 2, 1 - round % 2] {
                let mut runtime = vm(&a, &b);
                let chunks = programs[which].clone();
                let start = std::time::Instant::now();
                results[which] = runtime.run(chunks).unwrap();
                times[which].push(start.elapsed());
            }
        }
        assert_eq!(results[0], results[1]);
        for samples in &mut times {
            samples.sort();
        }
        println!(
            "{kind} units={}: before={}us after={}us speedup={:.2}x",
            a.encode_utf16().count(),
            times[0][2].as_micros(),
            times[1][2].as_micros(),
            times[0][2].as_secs_f64() / times[1][2].as_secs_f64()
        );
    }
}

#[test]
fn repeated_distance_calls_share_one_valid_helper() {
    let old = program(true, 20);
    let mut new = program(false, 20);
    assert_eq!(
        new.iter()
            .filter(|c| c.name == "__stdlib_levenshtein")
            .count(),
        1
    );
    let count = new.len();
    bundle::finalize_with_runtime_helpers(&mut new);
    assert_eq!(new.len(), count);
    let old_bytes = vybe_platform_wasm::write_wasm(&old);
    let new_bytes = vybe_platform_wasm::write_wasm(&new);
    for bytes in [&old_bytes, &new_bytes] {
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(bytes)
            .unwrap();
    }
    assert!(new[0].code.len() < old[0].code.len() / 10);
    assert_eq!(vm("kitten", "sitting").run(new).unwrap().as_f64(), 3.0);
    println!(
        "distance 20 calls: bytecode {} -> {}; wasm {} -> {}",
        old[0].code.len(),
        program(false, 20)[0].code.len(),
        old_bytes.len(),
        new_bytes.len()
    );
}

#[test]
#[ignore = "interleaved opt-level-0 debug wide Unicode benchmark"]
fn wide_unicode_distance_execution_benchmark() {
    let a = "é世".repeat(16);
    let b = "é界".repeat(1024);
    let programs = [program(true, 1), program(false, 1)];
    let mut times = [vec![], vec![]];
    let mut results = [Value::Null, Value::Null];
    for round in 0..3 {
        for which in [round % 2, 1 - round % 2] {
            let mut runtime = vm(&a, &b);
            let chunks = programs[which].clone();
            let start = std::time::Instant::now();
            results[which] = runtime.run(chunks).unwrap();
            times[which].push(start.elapsed());
        }
    }
    assert_eq!(results[0], results[1]);
    for samples in &mut times {
        samples.sort();
    }
    println!(
        "Unicode 32x2048 units, 3 interleaved samples: before={}us after={}us speedup={:.2}x",
        times[0][1].as_micros(),
        times[1][1].as_micros(),
        times[0][1].as_secs_f64() / times[1][1].as_secs_f64()
    );
}

fn validate_private_distance(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}
fn alias_distance_program(old: bool, body_name_collision: bool) -> Vec<Chunk> {
    use vybe_runtime::chunk::{LocalName, LocalOrigin};
    let mut root = Chunk::new("<ordinary-distance-alias>");
    let mut user = Chunk::new(if body_name_collision {
        "__stdlib_levenshtein"
    } else {
        "ordinary-source-function"
    });
    user.arity = 2;
    user.local_count = 2;
    // Even a Source debugger record with the private spelling cannot authorize
    // generated-body discovery. The ordinary function remains its own callable.
    user.local_names.push(LocalName::new(
        0,
        "<private-helper-body:__stdlib_levenshtein>",
        LocalOrigin::Source,
    ));
    user.emit_f64_const(42.0, 1);
    user.emit_op(Op::RETURN, 1);
    vybe_compiler::primitives::functions::emit_ref_func(&mut root, 1, 0, 1);
    globals::emit_write(&mut root, "__vybe_levenshtein", 1);
    root.emit_string_const("kitten", 2);
    root.emit_string_const("sitting", 2);
    let mut chunks = vec![root, user];
    if old {
        baseline::emit_levenshtein(&mut chunks, 0, 2, 2);
    } else {
        string_similarity::emit_levenshtein(&mut chunks, 0, 2, 2);
    }
    chunks[0].emit_op(Op::RETURN, 2);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    validate_private_distance(&chunks);
    chunks
}
#[test]
fn private_distance_preserves_legal_source_alias_and_ordinary_body_name_collision() {
    for collision in [false, true] {
        for old in [true, false] {
            let mut runtime = vm("", "");
            let chunks = alias_distance_program(old, collision);
            if !old {
                assert!(
                    !chunks[0]
                        .global_inits
                        .iter()
                        .any(|g| g.name == "__vybe_levenshtein")
                );
                assert_eq!(
                    chunks
                        .iter()
                        .filter(|c| c
                            .local_names
                            .iter()
                            .any(|n| n.origin == vybe_runtime::chunk::LocalOrigin::Compiler
                                && n.name == "<private-helper-body:__stdlib_levenshtein>"))
                        .count(),
                    1
                );
            }
            assert_eq!(runtime.run(chunks).unwrap(), Value::F64(3.0));
            let alias = runtime.global("__vybe_levenshtein").unwrap().clone();
            assert_eq!(
                runtime.invoke_callback(
                    &alias,
                    &[Value::String("a".into()), Value::String("b".into())]
                ),
                Value::F64(42.0)
            );
        }
    }
}
#[test]
fn private_distance_is_required_despite_public_exclusion_and_finalizes_idempotently() {
    let mut chunks = vec![Chunk::new("<excluded-distance>")];
    chunks[0].emit_string_const("kitten", 37);
    chunks[0].emit_string_const("sitting", 37);
    string_similarity::emit_levenshtein(&mut chunks, 0, 2, 37);
    chunks[0].emit_op(Op::RETURN, 37);
    bundle::finalize_with_runtime_helpers_excluding(&mut chunks, &["__stdlib_levenshtein"]);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    validate_private_distance(&chunks);
    assert!(
        !chunks[0]
            .global_inits
            .iter()
            .any(|g| g.name == "__vybe_levenshtein")
    );
    let before: Vec<_> = chunks.iter().map(|c| c.code.clone()).collect();
    bundle::finalize_with_runtime_helpers_excluding(&mut chunks, &["__stdlib_levenshtein"]);
    assert_eq!(
        before,
        chunks.iter().map(|c| c.code.clone()).collect::<Vec<_>>()
    );
    assert_eq!(vm("", "").run(chunks).unwrap(), Value::F64(3.0));
}
fn distance_boundary_placeholder() -> Chunk {
    let mut chunk = Chunk::new("unused-distance-boundary");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk
}
#[test]
fn private_distance_exact_u16_boundary_preserves_nonzero_argument_slots_and_native_fallback() {
    for initial in [u16::MAX as usize, u16::MAX as usize + 1] {
        let mut caller = Chunk::new("<distance-boundary>");
        caller.local_count = 7; // Inline body must use spilled args, not slots 0/1.
        caller.emit_string_const("A😀B", 37);
        caller.emit_string_const("A😀C", 37);
        let mut chunks = vec![caller];
        string_similarity::emit_levenshtein(&mut chunks, 0, 2, 37);
        chunks[0].emit_op(Op::RETURN, 37);
        chunks.resize(initial, distance_boundary_placeholder());
        bundle::finalize_with_runtime_helpers(&mut chunks);
        let fits = initial == u16::MAX as usize;
        assert_eq!(chunks.len(), initial + usize::from(fits));
        assert_eq!(
            chunks
                .iter()
                .filter(|c| c
                    .local_names
                    .iter()
                    .any(|n| n.name == "<private-helper-body:__stdlib_levenshtein>"))
                .count(),
            usize::from(fits)
        );
        assert_eq!(chunks[0].code.len(), chunks[0].lines.len());
        globals::declare_free_globals(&mut chunks);
        globals::normalize_global_table(&mut chunks);
        validate_private_distance(&chunks);
        let code = chunks[0].code.clone();
        bundle::finalize_with_runtime_helpers(&mut chunks);
        assert_eq!(chunks[0].code, code);
        let expected = vm("A😀B", "A😀C").run(program(true, 1)).unwrap();
        assert_eq!(vm("", "").run(chunks).unwrap(), expected);
    }
}
#[test]
fn private_distance_retained_callable_survives_cumulative_loads_and_alias_reassignment() {
    let mut root = Chunk::new("<retain-distance>");
    let mut caller = Chunk::new("source-distance");
    caller.arity = 2;
    caller.local_count = 2;
    caller.local_names.push(vybe_runtime::chunk::LocalName::new(
        0,
        "a",
        vybe_runtime::chunk::LocalOrigin::Source,
    ));
    caller.local_names.push(vybe_runtime::chunk::LocalName::new(
        1,
        "b",
        vybe_runtime::chunk::LocalOrigin::Source,
    ));
    caller.emit_op_u16(Op::LOCAL_GET, 0, 37);
    caller.emit_op_u16(Op::LOCAL_GET, 1, 37);
    vybe_compiler::primitives::functions::emit_ref_func(&mut root, 1, 0, 1);
    root.emit_op(Op::RETURN, 1);
    let mut chunks = vec![root, caller];
    string_similarity::emit_levenshtein(&mut chunks, 1, 2, 37);
    chunks[1].emit_op(Op::RETURN, 37);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    validate_private_distance(&chunks);
    let mut runtime = vm("", "");
    runtime
        .run(vec![distance_boundary_placeholder(); 65536])
        .unwrap();
    let retained = runtime.run(chunks).unwrap();
    let args = [
        Value::String("kitten".into()),
        Value::String("sitting".into()),
    ];
    assert_eq!(runtime.invoke_callback(&retained, &args), Value::F64(3.0));
    runtime.run(alias_distance_program(false, true)).unwrap();
    assert_eq!(runtime.invoke_callback(&retained, &args), Value::F64(3.0));
    let alias = runtime.global("__vybe_levenshtein").unwrap().clone();
    assert_eq!(runtime.invoke_callback(&alias, &args), Value::F64(42.0));
}
