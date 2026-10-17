use vybe_compiler::primitives::{globals, platforms, strings};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/ascii_class_before.rs"]
mod before;
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
const NEW: [Emit; 6] = [
    strings::emit_is_digit,
    strings::emit_is_alpha,
    strings::emit_is_alnum,
    strings::emit_is_space,
    strings::emit_is_upper,
    strings::emit_is_lower,
];
fn output(n: &str, v: &serde_json::Value) {
    let d = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::write(
        d.join(format!("ASCII_CLASS_{n}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
fn program(fast: bool, class: usize) -> Vec<Chunk> {
    let mut root = Chunk::new("ascii-class-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = 1;
    c.param_count = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let mut cs = vec![root, c];
    (if fast { NEW } else { OLD })[class](&mut cs, 1, 0);
    cs[1].emit_op(Op::RETURN, 0);
    globals::declare_free_globals(&mut cs);
    globals::normalize_global_table(&mut cs);
    cs
}
fn worker(fast: bool, class: usize) -> (VM, Value) {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    let f = vm.run(program(fast, class)).unwrap();
    (vm, f)
}
fn native(cs: &[Chunk]) -> Vec<u8> {
    let b = vybe_platform_wasm::write_wasm_with_options(
        cs,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: false,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&b)
        .unwrap();
    b
}
fn oracle(class: usize, s: &str) -> bool {
    !s.is_empty()
        && s.chars()
            .all(|c| c.is_alphabetic() || (class == 2 && c.is_ascii_digit()))
}
#[test]
fn unchanged_classifiers_keep_the_exact_original_emission() {
    for class in [0, 3, 4, 5] {
        for (a, b) in program(false, class).iter().zip(program(true, class)) {
            assert_eq!(a.code, b.code);
            assert_eq!(a.constants, b.constants);
            assert_eq!(a.imports, b.imports);
            assert_eq!(a.local_count, b.local_count);
        }
    }
}
#[test]
fn ascii_pairs_unicode_boundaries_errors_and_native_sizes() {
    let mut texts = vec![String::new()];
    for a in 0..128 {
        texts.push(char::from_u32(a).unwrap().to_string());
        for b in 0..128 {
            texts.push(format!(
                "{}{}",
                char::from_u32(a).unwrap(),
                char::from_u32(b).unwrap()
            ));
        }
    }
    let mut points = std::collections::BTreeSet::new();
    let mut previous = false;
    for cp in 0..=0x110000u32 {
        let yes = char::from_u32(cp).is_some_and(char::is_alphabetic);
        if yes != previous {
            for p in [cp.saturating_sub(1), cp, cp + 1] {
                points.insert(p);
            }
            previous = yes;
        }
    }
    for cp in points {
        if let Some(c) = char::from_u32(cp) {
            texts.push(c.to_string());
            texts.push(format!("A{c}9"));
            texts.push(format!("{c}a"));
        }
    }
    for s in [
        "aAZz09",
        "AZaz",
        "[\\]^_`{|}~",
        "Ωé中𐐀𐐨",
        "éAé",
        "A😀A",
        "a\u{301}",
        "\0",
        "\u{10ffff}",
    ] {
        for n in [1, 2, 31, 127] {
            texts.push(s.repeat(n));
        }
    }
    let mut checks = 0;
    let mut sizes = vec![];
    let mut corpus = vec![];
    for class in [1, 2] {
        let old = program(false, class);
        let new = program(true, class);
        let bytes = [native(&old), native(&new)];
        assert!(new[1].code.len() < old[1].code.len());
        assert_eq!(old[1].local_count, new[1].local_count);
        assert_eq!(old[1].imports, new[1].imports);
        assert_eq!(old[1].constants, new[1].constants);
        native_export::export_chunks(&old, &format!("ascii-class-{class}-before"));
        native_export::export_chunks(&new, &format!("ascii-class-{class}-after"));
        sizes.push(serde_json::json!({"class":class,"old_body":old[1].code.len(),"new_body":new[1].code.len(),"old_native":bytes[0].len(),"new_native":bytes[1].len(),"old_locals":old[1].local_count,"new_locals":new[1].local_count}));
        let (va, fa) = worker(false, class);
        let (vb, fb) = worker(true, class);
        let mut vs = [va, vb];
        let fs = [fa, fb];
        for s in &texts {
            let expected = oracle(class, s);
            let x = Value::String(s.clone().into());
            for side in 0..2 {
                assert_eq!(
                    vs[side].invoke(&fs[side], &[x.clone()]).unwrap(),
                    Value::Bool(expected),
                    "class={class} {s:?}"
                );
                checks += 1;
            }
            corpus.push(serde_json::json!({"class":class,"text":s,"expected":expected}));
        }
        for input in [
            Value::Null,
            Value::Undefined,
            Value::I32(7),
            Value::Bool(false),
        ] {
            let mut outcomes = vec![];
            for fast in [false, true] {
                let (mut v, f) = worker(fast, class);
                outcomes.push(v.invoke(&f, &[input.clone()]).map_err(|e| e.message));
            }
            assert_eq!(outcomes[0], outcomes[1]);
            assert!(outcomes[0].is_err());
            checks += 2;
        }
    }
    output(
        "GATES",
        &serde_json::json!({"checks":checks,"texts":texts.len(),"sizes":sizes}),
    );
    output("CORPUS", &serde_json::json!({"queries":corpus}));
}

#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (class, label, text, control) in [
        (1, "alpha-short", "AbCdZz".into(), false),
        (1, "alpha-long", "AbCdZz".repeat(128), false),
        (1, "alpha-reject", "[a".into(), false),
        (1, "alpha-unicode", "Ωé中𐐀".repeat(32), false),
        (2, "alnum-short", "A1b2Z9".into(), false),
        (2, "alnum-long", "A1b2Z9".repeat(128), false),
        (2, "alnum-reject", "_a".into(), false),
        (2, "alnum-unicode", "Ωé中𐐀1".repeat(32), false),
        (1, "control", "AbCdZz".repeat(128), true),
    ] {
        let (va, fa) = worker(false, class);
        let (vb, fb) = worker(!control, class);
        let mut vs = [va, vb];
        let fs = [fa, fb];
        let x = Value::String(text.into());
        let mut wall = [vec![], vec![]];
        let mut cpu = [vec![], vec![]];
        let mut calls = vec![];
        for pair in 0..7 {
            let mut secs = [0., 0.];
            let mut cpus = [0., 0.];
            let mut counts = [0usize, 0];
            let mut turn = pair;
            while secs[0] < 0.1 || secs[1] < 0.1 || counts.iter().any(|n| *n < 8) {
                for side in if turn % 2 == 0 { [0, 1] } else { [1, 0] } {
                    if secs[side] >= 0.1 && counts[side] >= 8 {
                        continue;
                    }
                    let c = thread_cpu_seconds();
                    let t = Instant::now();
                    std::hint::black_box(
                        vs[side]
                            .invoke(&fs[side], std::slice::from_ref(&x))
                            .unwrap(),
                    );
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
        rows.push(serde_json::json!({"class":class,"label":label,"median_speedup":wr[3],"median_thread_cpu_speedup":cr[3],"wall":wall,"cpu":cpu,"calls":calls}));
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
