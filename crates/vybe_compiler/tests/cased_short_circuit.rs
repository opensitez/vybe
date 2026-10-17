#![cfg(feature = "primitive-benchmarks")]
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/cased_short_before.rs"]
mod before;
#[path = "fixtures/ascii_class_before.rs"]
mod before_ascii_fold;
#[path = "fixtures/cased_short_candidate.rs"]
mod candidate;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
type Emit = fn(&mut [Chunk], usize, u32);
const OLD: [Emit; 6] = [
    before::emit_is_digit,
    before::emit_is_alpha,
    before::emit_is_alnum,
    before::emit_is_space,
    before::emit_is_upper,
    before::emit_is_lower,
];
const MEASURED: [Emit; 6] = [
    candidate::emit_is_digit,
    candidate::emit_is_alpha,
    candidate::emit_is_alnum,
    candidate::emit_is_space,
    candidate::emit_is_upper,
    candidate::emit_is_lower,
];
const NEW: [Emit; 6] = [
    primitives::strings::emit_is_digit,
    primitives::strings::emit_is_alpha,
    primitives::strings::emit_is_alnum,
    primitives::strings::emit_is_space,
    primitives::strings::emit_is_upper,
    primitives::strings::emit_is_lower,
];
fn program(fast: bool, class: usize) -> Vec<Chunk> {
    program_emit((if fast { NEW } else { OLD })[class])
}
fn program_emit(emit: Emit) -> Vec<Chunk> {
    let mut root = Chunk::new("classifier-short-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = 1;
    c.param_count = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let mut cs = vec![root, c];
    emit(&mut cs, 1, 0);
    cs[1].emit_op(Op::RETURN, 0);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    cs
}
fn function_index(f: &Value) -> usize {
    let Value::Object(o) = f else {
        panic!("expected function")
    };
    let o = o.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Function(f) = &o.kind else {
        panic!("expected guest function")
    };
    f.chunk_index
}
fn vm() -> VM {
    let mut v = VM::new();
    primitives::platforms::register_platforms_all(&mut v);
    v
}
fn pred(class: usize, c: char) -> bool {
    match class {
        0 => c.is_ascii_digit(),
        1 => c.is_alphabetic(),
        2 => c.is_alphabetic() || c.is_ascii_digit(),
        3 => c.is_whitespace(),
        4 => c.is_uppercase(),
        5 => c.is_lowercase(),
        _ => unreachable!(),
    }
}
fn oracle(class: usize, s: &str) -> bool {
    if class < 4 {
        !s.is_empty() && s.chars().all(|c| pred(class, c))
    } else {
        s.chars().any(|c| pred(class, c))
            && !s.chars().any(|c| pred(if class == 4 { 5 } else { 4 }, c))
    }
}
fn output(n: &str, v: &serde_json::Value) {
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join(format!("../../CASED_SHORT_{n}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
#[test]
fn production_preserves_measured_candidate_lineage() {
    let frozen: [Emit; 6] = [
        before_ascii_fold::emit_is_digit,
        before_ascii_fold::emit_is_alpha,
        before_ascii_fold::emit_is_alnum,
        before_ascii_fold::emit_is_space,
        before_ascii_fold::emit_is_upper,
        before_ascii_fold::emit_is_lower,
    ];
    for class in 0..6 {
        let a = program(true, class);
        let b = program_emit(MEASURED[class]);
        let frozen_program = program_emit(frozen[class]);
        for (old, measured) in frozen_program.iter().zip(&b) {
            assert_eq!(old.code, measured.code);
            assert_eq!(old.imports, measured.imports);
            assert_eq!(old.constants, measured.constants);
        }
        for (index, (a, b)) in a.iter().zip(&b).enumerate() {
            // ASCII case folding is verified independently in ascii_class_fold.
            if index == 1 && matches!(class, 1 | 2) {
                assert!(a.code.len() < b.code.len());
            } else {
                assert_eq!(a.code, b.code);
            }
            assert_eq!(a.constants, b.constants);
            assert_eq!(a.imports, b.imports);
            assert_eq!(a.global_imports, b.global_imports);
            assert_eq!(a.local_count, b.local_count);
            assert_eq!(a.arity, b.arity);
            assert_eq!(a.param_count, b.param_count);
        }
    }
}
#[test]
fn values_sizes_reads_and_unicode_boundaries() {
    let mut ranges = vec![];
    let mut points = std::collections::BTreeSet::new();
    points.extend(0..128);
    for class in 0..6 {
        let mut rs = vec![];
        let mut start = None;
        for cp in 0..=0x110000u32 {
            let yes = char::from_u32(cp).is_some_and(|c| pred(class, c));
            match (start, yes) {
                (None, true) => start = Some(cp),
                (Some(lo), false) => {
                    rs.push([lo, cp]);
                    start = None;
                    for p in [lo.saturating_sub(1), lo, cp.saturating_sub(1), cp] {
                        points.insert(p);
                    }
                }
                _ => {}
            }
        }
        ranges.push(rs);
    }
    // All predicate transition boundaries are exercised in VMs; native gates
    // additionally cover every Unicode scalar for both cased predicates.
    let mut texts: Vec<String> = points
        .into_iter()
        .filter_map(char::from_u32)
        .map(|c| c.to_string())
        .collect();
    for s in [
        "",
        "A",
        "a",
        "123",
        "😀",
        "中",
        "A\u{301}",
        "ΩΩΩ",
        "ΩAΩ",
        "ΩΜΕΓΑΚΟΣΜΟΣ",
        "ωμέγακόσμος",
        "ÉéÉ",
        "ééé",
        "ßß",
        "𐐀𐐀",
        "𐐨𐐨",
        "Ω 😀",
        "abc123",
        "A-a",
        "\u{10ffff}",
    ] {
        texts.push(s.into());
        texts.push(format!("A{s}"));
        texts.push(format!("{s}a"));
        texts.push(format!("a{s}"));
        texts.push(format!("{s}A"));
    }
    for n in [2, 3, 31, 32, 127] {
        for s in ["ABC! ", "abc! ", "ΩÉ𐐀", "ωé𐐨", "中😀1"] {
            texts.push(s.repeat(n));
        }
    }
    let mut checks = 0;
    let mut rows = vec![];
    for class in 0..6 {
        let mut v = vm();
        let fs = [
            v.run(program(false, class)).unwrap(),
            v.run(program(true, class)).unwrap(),
        ];
        let indices = [function_index(&fs[0]), function_index(&fs[1])];
        assert_ne!(indices[0], indices[1]);
        for (side, fast) in [false, true].into_iter().enumerate() {
            let expected = program(fast, class);
            assert_eq!(v.chunks[indices[side]].local_count, expected[1].local_count);
            assert_eq!(v.chunks[indices[side]].code.len(), expected[1].code.len());
        }
        for s in &texts {
            let expected = Value::Bool(oracle(class, s));
            for f in &fs {
                assert_eq!(
                    v.invoke(f, &[Value::String(s.clone().into())]).unwrap(),
                    expected,
                    "class={class} text={s:?}"
                );
                checks += 1;
            }
        }
        for x in [
            Value::Null,
            Value::Undefined,
            Value::I32(1),
            Value::Bool(false),
        ] {
            let mut outcomes = vec![];
            for fast in [false, true] {
                let mut bad = vm();
                let f = bad.run(program(fast, class)).unwrap();
                outcomes.push(bad.invoke(&f, &[x.clone()]).map_err(|e| e.message));
            }
            assert_eq!(outcomes[0], outcomes[1]);
            assert!(outcomes[0].is_err());
            checks += 1;
        }
        let old = program(false, class);
        let actual = program(true, class);
        assert_eq!(old[1].imports, actual[1].imports);
        assert_eq!(old[1].global_imports, actual[1].global_imports);
        assert_eq!(old[1].constants, actual[1].constants);
        assert_eq!(
            old[1].local_count - actual[1].local_count,
            if class < 4 { 1 } else { 3 }
        );
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = program(fast, class);
            let b = vybe_platform_wasm::write_wasm_with_options(
                &cs,
                vybe_platform_wasm::WasmWriteOptions {
                    include_name_section: false,
                    include_vybe_metadata: false,
                },
            );
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&b)
                .unwrap();
            native_export::export_chunks(&cs, &format!("cased-short-{class}-{fast}"));
            sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":b.len()}));
        }
        assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
        assert!(sizes[1]["locals"].as_u64() < sizes[0]["locals"].as_u64());
        assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
        let mut work = vec![];
        for s in [
            "ABCXYZ".repeat(16),
            "ΩÉ𐐀".repeat(16),
            "ωé𐐨".repeat(16),
            "中😀1".repeat(16),
        ] {
            let mut counts = vec![];
            for f in &fs {
                v.record_runtime_perf(true);
                v.invoke(f, &[Value::String(s.clone().into())]).unwrap();
                counts.push(v.take_runtime_perf().unwrap().dispatches);
            }
            work.push(serde_json::json!({"text":s,"dispatches":counts}));
        }
        rows.push(
            serde_json::json!({"class":class,"chunk_indices":indices,"sizes":sizes,"work":work}),
        );
    }
    let artifact = std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap();
    std::fs::write(
        std::path::Path::new(&artifact).join("cased-oracle.json"),
        serde_json::to_vec(&ranges).unwrap(),
    )
    .unwrap();
    output(
        "GATES",
        &serde_json::json!({"checks":checks,"texts":texts.len(),"rows":rows}),
    );
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (label, class, text, control) in [
        ("control", 4, "ABCXYZ".repeat(32), true),
        ("digit", 0, "1234567890".repeat(20), false),
        ("alpha-ascii", 1, "Alphabetic".repeat(20), false),
        ("alpha-unicode", 1, "ΩÉ中𐐀".repeat(20), false),
        ("alnum", 2, "abc123Ω".repeat(20), false),
        ("space", 3, " \t\n".repeat(60), false),
        ("empty", 4, "".into(), false),
        ("tiny", 4, "Ab".into(), false),
        ("ascii-upper", 4, "ABCXYZ".repeat(32), false),
        ("ascii-lower", 5, "abcxyz".repeat(32), false),
        ("unicode-upper", 4, "ΩÉ𐐀".repeat(32), false),
        ("unicode-lower", 5, "ωé𐐨".repeat(32), false),
        ("unicode-repeat", 4, "Ω".repeat(192), false),
        ("caseless", 4, "中😀1".repeat(32), false),
        ("early-opposite", 4, format!("a{}", "ΩÉ𐐀".repeat(32)), false),
        ("late-opposite", 4, format!("{}a", "ΩÉ𐐀".repeat(32)), false),
    ] {
        let mut v = vm();
        let a = v.run(program(false, class)).unwrap();
        let b = if control {
            a.clone()
        } else {
            v.run(program(true, class)).unwrap()
        };
        let fs = [a, b];
        if !control {
            assert_ne!(function_index(&fs[0]), function_index(&fs[1]));
        }
        let args = [Value::String(text.into())];
        for f in &fs {
            for _ in 0..8 {
                std::hint::black_box(v.invoke(f, &args).unwrap());
            }
        }
        let mut samples = [vec![], vec![]];
        let mut cpu_samples = [vec![], vec![]];
        for pair in 0..7 {
            let mut secs = [0.0, 0.0];
            let mut cpu_secs = [0.0, 0.0];
            let mut n = 0;
            let mut turn = pair;
            while secs[0] < 0.5 || secs[1] < 0.5 || n < 8 {
                for side in if turn % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let cpu = thread_cpu_seconds();
                    let t = Instant::now();
                    std::hint::black_box(v.invoke(&fs[side], &args).unwrap());
                    secs[side] += t.elapsed().as_secs_f64();
                    if let (Some(a), Some(b)) = (cpu, thread_cpu_seconds()) {
                        cpu_secs[side] += b - a;
                    }
                }
                n += 1;
                turn += 1;
            }
            for side in 0..2 {
                samples[side].push(secs[side] / n as f64);
                cpu_samples[side].push(cpu_secs[side] / n as f64);
            }
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        let mut cpu_ratios: Vec<_> = cpu_samples[0]
            .iter()
            .zip(&cpu_samples[1])
            .filter(|(_, b)| **b > 0.0)
            .map(|(a, b)| a / b)
            .collect();
        cpu_ratios.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"label":label,"class":class,"control":control,"median_speedup":ratios[3],"median_thread_cpu_speedup":cpu_ratios.get(cpu_ratios.len()/2),"seconds_per_call":samples,"thread_cpu_seconds_per_call":cpu_samples}));
    }
    output(
        "MEASUREMENTS",
        &serde_json::json!({"pairs":7,"same_vm":true,"equal_calls_interleaved":true,"minimum_sample_ms":500,"cases":rows}),
    );
}

// Diagnostic only: the local macOS SDK defines CLOCK_THREAD_CPUTIME_ID=16.
// Wall time remains recorded; thread CPU time excludes descheduling by other jobs.
#[cfg(target_os = "macos")]
fn thread_cpu_seconds() -> Option<f64> {
    #[repr(C)]
    struct Timespec {
        sec: i64,
        nsec: i64,
    }
    unsafe extern "C" {
        fn clock_gettime(clock: i32, result: *mut Timespec) -> i32;
    }
    let mut result = Timespec { sec: 0, nsec: 0 };
    assert_eq!(unsafe { clock_gettime(16, &mut result) }, 0);
    Some(result.sec as f64 + result.nsec as f64 * 1e-9)
}
#[cfg(not(target_os = "macos"))]
fn thread_cpu_seconds() -> Option<f64> {
    None
}
