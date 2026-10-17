#![cfg(feature = "primitive-benchmarks")]
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/levenshtein_two_rows.rs"]
mod before;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;

fn output(name: &str, value: &serde_json::Value) {
    let root = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::write(
        root.join(format!("LEVEN_ROW_{name}.json")),
        serde_json::to_vec_pretty(value).unwrap(),
    )
    .unwrap();
}
fn program(old: bool) -> Vec<Chunk> {
    let mut c = vec![Chunk::new("leven-row-owner"), Chunk::new("entry")];
    c[1].arity = 2;
    c[1].param_count = 2;
    c[1].local_count = 2;
    for slot in [0, 1] {
        c[1].emit_op_u16(Op::LOCAL_GET, slot, 0);
    }
    primitives::string_similarity::emit_levenshtein(&mut c, 1, 2, 0);
    c[1].emit_op(Op::RETURN, 0);
    primitives::functions::emit_ref_func(&mut c[0], 1, 0, 0);
    c[0].emit_op(Op::RETURN, 0);
    primitives::bundle::finalize_with_runtime_helpers(&mut c);
    if old {
        let i = c
            .iter()
            .position(|x| x.name == "__stdlib_levenshtein")
            .unwrap();
        let mut chunk = primitives::functions::create_function_chunk("__stdlib_levenshtein", 2);
        chunk.local_count = 2;
        chunk.local_names = c[i].local_names.clone();
        let mut body = vec![chunk];
        before::emit_levenshtein_body(&mut body, 0, 0, 1, 0);
        body[0].emit_op(Op::RETURN, 0);
        c[i] = body.remove(0);
    }
    primitives::globals::declare_free_globals(&mut c);
    primitives::globals::normalize_global_table(&mut c);
    c
}
fn worker(c: Vec<Chunk>) -> (VM, Value) {
    let mut v = VM::new();
    primitives::platforms::register_platforms_all(&mut v);
    let f = v.run(c).unwrap();
    (v, f)
}
fn units(s: &str) -> Vec<u16> {
    s.encode_utf16()
        .map(|u| {
            if (0xd800..=0xdfff).contains(&u) {
                0xfffd
            } else {
                u
            }
        })
        .collect()
}
fn oracle(a: &str, b: &str) -> usize {
    let (a, b) = (units(a), units(b));
    let mut matrix = vec![vec![0; b.len() + 1]; a.len() + 1];
    for i in 0..=a.len() {
        matrix[i][0] = i;
    }
    for j in 0..=b.len() {
        matrix[0][j] = j;
    }
    for i in 1..=a.len() {
        for j in 1..=b.len() {
            matrix[i][j] = (matrix[i - 1][j] + 1)
                .min(matrix[i][j - 1] + 1)
                .min(matrix[i - 1][j - 1] + usize::from(a[i - 1] != b[j - 1]));
        }
    }
    matrix[a.len()][b.len()]
}
fn native(c: &[Chunk]) -> Vec<u8> {
    let b = vybe_platform_wasm::write_wasm_with_options(
        c,
        vybe_platform_wasm::WasmWriteOptions {
            include_name_section: false,
            include_vybe_metadata: false,
        },
    );
    if let Err(e) =
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&b)
    {
        let h = c.iter().find(|x| x.name == "__stdlib_levenshtein").unwrap();
        let dir = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
        std::fs::write(dir.join(format!("leven-failure-{}.wasm", h.code.len())), &b).unwrap();
        panic!("helper bytes {}: {e}", h.code.len());
    }
    b
}
#[test]
fn complete_helper_matches_frozen_rows_and_independent_matrix() {
    let old = program(true);
    let new = program(false);
    let oi = old
        .iter()
        .position(|x| x.name == "__stdlib_levenshtein")
        .unwrap();
    let ni = new
        .iter()
        .position(|x| x.name == "__stdlib_levenshtein")
        .unwrap();
    assert_eq!(old[oi].local_count, new[ni].local_count);
    assert!(new[ni].code.len() < old[oi].code.len());
    let bytes = [native(&old), native(&new)];
    native_export::export_chunks(&old, "leven-row-before");
    native_export::export_chunks(&new, "leven-row-after");
    output(
        "SIZES",
        &serde_json::json!({"old_body":old[oi].code.len(),"new_body":new[ni].code.len(),"old_locals":old[oi].local_count,"new_locals":new[ni].local_count,"old_native":bytes[0].len(),"new_native":bytes[1].len()}),
    );
    let (mut a, fa) = worker(old);
    let (mut b, fb) = worker(new);
    let mut strings = vec![String::new()];
    for n in 1..=4 {
        for mask in 0..1usize << n {
            strings.push(
                (0..n)
                    .map(|i| if (mask >> i) & 1 == 0 { 'a' } else { 'b' })
                    .collect(),
            );
        }
    }
    strings.extend(
        [
            "é",
            "€",
            "😀",
            "🦀",
            "A😀B",
            "e\u{301}",
            "\0x",
            "\u{fffd}",
            "kitten",
            "sitting",
            "________ab",
            "________ba",
        ]
        .map(str::to_owned),
    );
    let mut pairs: Vec<_> = strings
        .iter()
        .flat_map(|x| strings.iter().map(move |y| (x.clone(), y.clone())))
        .collect();
    for n in [0, 1, 7, 8, 16, 31, 32, 33, 40, 64] {
        for m in [0, 1, 8, 33, 40] {
            for alphabet in ["ab", "é€", "😀🦀"] {
                pairs.push((
                    alphabet.chars().cycle().take(n).collect(),
                    alphabet.chars().cycle().skip(1).take(m).collect(),
                ));
            }
        }
    }
    let mut checks = 0;
    let mut corpus = vec![];
    for (x, y) in pairs {
        let expected = oracle(&x, &y) as f64;
        let args = [
            Value::String(x.clone().into()),
            Value::String(y.clone().into()),
        ];
        assert_eq!(
            a.invoke(&fa, &args).unwrap(),
            Value::F64(expected),
            "old {x:?}/{y:?}"
        );
        assert_eq!(
            b.invoke(&fb, &args).unwrap(),
            Value::F64(expected),
            "new {x:?}/{y:?}"
        );
        checks += 2;
        corpus.push(serde_json::json!({"a":x,"b":y,"vm":expected}));
    }
    for args in [
        [Value::Null, Value::String("a".into())],
        [Value::String("a".into()), Value::Undefined],
        [Value::I32(1), Value::Null],
    ] {
        let trap = |e: vybe_runtime::VMError| e.to_string().lines().next().unwrap().to_owned();
        assert_eq!(
            a.invoke(&fa, &args).map_err(trap),
            b.invoke(&fb, &args).map_err(trap)
        );
        checks += 1;
    }
    output("GATES", &serde_json::json!({"checks":checks}));
    output("CORPUS", &serde_json::json!({"queries":corpus}));
}

#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (label, x, y, control) in [
        ("latin64", "ab".repeat(32), "ba".repeat(32), false),
        ("unicode64", "é€".repeat(32), "€é".repeat(32), false),
        ("short-row", "ab".repeat(64), "ba".into(), false),
        ("myers-control", "ab".repeat(8), "ba".repeat(8), false),
        ("empty", String::new(), "a".repeat(128), false),
        (
            "identical-code-control",
            "ab".repeat(32),
            "ba".repeat(32),
            true,
        ),
    ] {
        for repeats in [1] {
            let (va, fa) = worker(program(true));
            let (vb, fb) = worker(program(control));
            let mut vs = [va, vb];
            let fs = [fa, fb];
            let args = [
                Value::String(x.clone().into()),
                Value::String(y.clone().into()),
            ];
            for side in 0..2 {
                assert_eq!(
                    vs[side].invoke(&fs[side], &args).unwrap(),
                    Value::F64(oracle(&x, &y) as f64)
                );
            }
            let mut wall = [vec![], vec![]];
            let mut cpu = [vec![], vec![]];
            let mut calls = vec![];
            for pair in 0..7 {
                let mut secs = [0., 0.];
                let mut cpus = [0., 0.];
                let mut counts = [0usize, 0];
                let mut turn = pair;
                while secs[0] < 0.1 || secs[1] < 0.1 || counts.iter().any(|n| *n < 1) {
                    for side in if turn % 2 == 0 { [0, 1] } else { [1, 0] } {
                        if secs[side] >= 0.1 && counts[side] >= 1 {
                            continue;
                        }
                        let c = thread_cpu_seconds();
                        let t = Instant::now();
                        std::hint::black_box(vs[side].invoke(&fs[side], &args).unwrap());
                        secs[side] += t.elapsed().as_secs_f64();
                        cpus[side] += thread_cpu_seconds() - c;
                        counts[side] += 1;
                    }
                    turn += 1;
                }
                for side in 0..2 {
                    wall[side].push(secs[side] / counts[side] as f64);
                    cpu[side].push(cpus[side] / counts[side] as f64);
                }
                calls.push(counts);
            }
            let mut wr: Vec<_> = wall[0].iter().zip(&wall[1]).map(|(a, b)| a / b).collect();
            wr.sort_by(f64::total_cmp);
            let mut cr: Vec<_> = cpu[0].iter().zip(&cpu[1]).map(|(a, b)| a / b).collect();
            cr.sort_by(f64::total_cmp);
            rows.push(serde_json::json!({"label":label,"repeats":repeats,"median_speedup":wr[3],"median_thread_cpu_speedup":cr[3],"wall":wall,"cpu":cpu,"calls":calls}));
        }
    }
    output(
        "MEASUREMENTS",
        &serde_json::json!({"pairs":7,"minimum_ms_per_side":100,"separate_vms":true,"cases":rows}),
    );
}
#[cfg(target_os = "macos")]
fn thread_cpu_seconds() -> f64 {
    #[repr(C)]
    struct Timespec {
        sec: i64,
        nsec: i64,
    }
    unsafe extern "C" {
        fn clock_gettime(clock: i32, result: *mut Timespec) -> i32;
    }
    let mut t = Timespec { sec: 0, nsec: 0 };
    assert_eq!(unsafe { clock_gettime(16, &mut t) }, 0);
    t.sec as f64 + t.nsec as f64 * 1e-9
}
#[cfg(not(target_os = "macos"))]
fn thread_cpu_seconds() -> f64 {
    0.0
}

fn array_program(element: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("array-new-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = 2;
    c.param_count = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::ARRAY_NEW, 0, 0);
    if element {
        c.emit_i32_const(0, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    } else {
        c.emit_op(Op::ARRAY_LENGTH, 0);
    }
    c.emit_op(Op::RETURN, 0);
    vec![root, c]
}
#[test]
fn standard_array_new_unboxes_only_length_and_preserves_initializer() {
    let mut checks = 0;
    for element in [false, true] {
        let c = array_program(element);
        native(&c);
        native_export::export_chunks(
            &c,
            if element {
                "array-new-element"
            } else {
                "array-new-length"
            },
        );
        let (mut v, f) = worker(c);
        for value in [
            Value::I32(-7),
            Value::String("fill".into()),
            Value::Bool(true),
            Value::Null,
        ] {
            for n in [0, 1, 2, 256] {
                if element && n == 0 {
                    continue;
                }
                let result = v.invoke(&f, &[value.clone(), Value::I32(n)]).unwrap();
                assert_eq!(
                    result,
                    if element {
                        value.clone()
                    } else {
                        Value::I32(n)
                    }
                );
                checks += 1;
            }
        }
    }
    output(
        "ARRAY_NEW_GATES",
        &serde_json::json!({"checks":checks,"complete_native_modules":2}),
    );
}

#[test]
fn strict_query_order_and_wrapped_length_fallback_match() {
    use std::sync::{Arc, Mutex};
    let mut checks = 0;
    for m in [-1, 0, 1, 3] {
        for n in [-1, 0, 1, 3] {
            let mut results = vec![];
            for old in [true, false] {
                let mut v = VM::new();
                primitives::platforms::register_platforms_all(&mut v);
                let trace = Arc::new(Mutex::new(vec![]));
                let log = trace.clone();
                v.register_host_fn(
                    "test:leven-length",
                    "length",
                    Box::new(move |_, args| {
                        let s = args[0].as_str();
                        log.lock().unwrap().push(format!("length:{s}"));
                        Value::I32(if s == "aaa" { m } else { n })
                    }),
                );
                let log = trace.clone();
                v.register_host_fn(
                    "ecma:string",
                    "charAt",
                    Box::new(move |_, args| {
                        let s = args[0].as_str();
                        let i = args[1].as_i32();
                        log.lock().unwrap().push(format!("char:{s}:{i}"));
                        let unit = s.encode_utf16().nth(i as usize);
                        Value::String(
                            unit.map(|u| String::from_utf16_lossy(&[u]))
                                .unwrap_or_default()
                                .into(),
                        )
                    }),
                );
                let f = v.run(diagnostic_program(old)).unwrap();
                let result = v
                    .invoke(
                        &f,
                        &[Value::String("aaa".into()), Value::String("bbb".into())],
                    )
                    .map_err(|e| e.to_string().lines().next().unwrap().to_owned());
                let seen = trace.lock().unwrap().clone();
                assert!(
                    seen.starts_with(&["length:aaa".to_owned(), "length:bbb".to_owned()]),
                    "override was not called: {seen:?}"
                );
                results.push((result, seen));
            }
            assert_eq!(results[0], results[1], "m={m},n={n}");
            checks += 1;
        }
    }
    output("QUERY_ORDER_GATES", &serde_json::json!({"checks":checks}));
}

// The VM intrinsifies core String.length, so use an explicit diagnostic
// import to observe and inject length results. Native gates keep real builtins.
fn diagnostic_program(old: bool) -> Vec<Chunk> {
    let mut chunks = program(old);
    for c in &mut chunks {
        for i in c.imports.iter_mut() {
            if i.module == "wasm:js-string" && i.name == "length" {
                i.module = "test:leven-length".into();
            }
        }
    }
    chunks
}
