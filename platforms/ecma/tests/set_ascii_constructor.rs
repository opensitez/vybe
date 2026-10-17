#![cfg(feature = "primitive-benchmarks")]
use std::sync::Arc;
use vybe_runtime::{
    VM, Value,
    value::{Object, ObjectKind},
};
fn output(name: &str, v: &serde_json::Value) {
    let d = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::create_dir_all(&d).unwrap();
    std::fs::write(
        d.join(format!("SET_ASCII_{name}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
fn worker(fast: bool) -> VM {
    let mut v = VM::new();
    vybe_platform_ecma::set::register(&mut v);
    vybe_platform_ecma::set::benchmark_set_ascii_constructor(&mut v, fast);
    v
}
fn function(v: &VM, op: usize) -> Value {
    v.func_table[v
        .resolve_host_function_index("ecma:set", if op == 0 { "new" } else { "fromIterable" })
        .unwrap()]
    .clone()
}
fn values(v: &Value) -> Vec<Value> {
    let Value::Object(o) = v else { panic!("set") };
    let o = o.lock().unwrap();
    let ObjectKind::Set(s) = &o.kind else {
        panic!("Set backing")
    };
    s.iter().cloned().collect()
}
fn capacity(v: &Value) -> usize {
    let Value::Object(o) = v else { panic!("set") };
    let o = o.lock().unwrap();
    let ObjectKind::Set(s) = &o.kind else {
        panic!("Set backing")
    };
    s.capacity()
}
fn equal(a: &[Value], b: &[Value], ascii: bool) {
    assert_eq!(a, b);
    for (x, y) in a.iter().zip(b) {
        match (x, y) {
            (Value::F64(x), Value::F64(y)) => assert_eq!(x.to_bits(), y.to_bits()),
            (Value::Object(x), Value::Object(y)) => assert!(Arc::ptr_eq(x, y)),
            (Value::String(x), Value::String(y)) if ascii => assert!(Arc::ptr_eq(x, y)),
            _ => {}
        }
    }
}
fn check(ws: &mut [VM; 2], s: &str, record: bool, corpus: &mut Vec<serde_json::Value>) -> usize {
    let arg = Value::String(s.into());
    let mut expected: Vec<Value> = vec![];
    for ch in s.chars() {
        let v = Value::String(ch.to_string().into());
        if !expected.contains(&v) {
            expected.push(v);
        }
    }
    for op in 0..2 {
        let mut results = vec![];
        for v in &mut *ws {
            let f = function(v, op);
            let r = v.invoke(&f, &[arg.clone()]).unwrap();
            let actual = values(&r);
            assert_eq!(actual, expected);
            results.push(actual);
            if s.len() >= 16 && s.is_ascii() && results.len() == 2 {
                assert!(capacity(&r) <= 256);
            }
        }
        equal(&results[0], &results[1], s.is_ascii());
        if record {
            corpus.push(serde_json::json!({"op":op,"source":s,"result":results[1].iter().map(|v|v.as_str().to_owned()).collect::<Vec<_>>()}));
        }
    }
    4
}
#[test]
fn ordered_ascii_domain_unicode_and_original_fallbacks() {
    vybe_runtime::heap::enable_tracking();
    let mut ws = [worker(false), worker(true)];
    let baseline = vybe_runtime::heap::snapshot();
    let mut checks = 0;
    let mut corpus = vec![];
    for a in 0..128u8 {
        for b in 0..128u8 {
            let s = format!("{}{}", a as char, b as char).repeat(20);
            checks += check(
                &mut ws,
                &s,
                (a as usize * 128 + b as usize) % 257 == 0,
                &mut corpus,
            );
            if checks % 128 == 0 {
                vybe_runtime::heap::restore(&baseline);
            }
        }
    }
    for rotate in 0..128 {
        let s: String = (0..128)
            .map(|i| ((i + rotate) % 128) as u8 as char)
            .collect();
        checks += check(&mut ws, &s.repeat(4), true, &mut corpus);
        vybe_runtime::heap::restore(&baseline);
    }
    for s in [
        "".to_string(),
        "x".into(),
        "a".repeat(15),
        "a".repeat(16),
        "Ω".repeat(1024),
        "😀éΩ中\0A".repeat(64),
        "a".repeat(2048) + "😀",
        (0..128)
            .map(|i| i as u8 as char)
            .collect::<String>()
            .repeat(4)
            + "😀",
        (0..128)
            .map(|i| i as u8 as char)
            .collect::<String>()
            .repeat(512),
    ] {
        checks += check(&mut ws, &s, true, &mut corpus);
        vybe_runtime::heap::restore(&baseline);
    }
    let chars = ['A', 'B', '\0', '\n', 'é', 'Ω', '中', '😀', '𐐀', '\u{200d}'];
    let mut state = 0x123456789abcdef0u64;
    for _ in 0..512 {
        let mut s = String::new();
        for _ in 0..40 {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            s.push(chars[state as usize % chars.len()]);
        }
        checks += check(&mut ws, &s, true, &mut corpus);
        vybe_runtime::heap::restore(&baseline);
    }
    for arg in [
        Value::Undefined,
        Value::Null,
        Value::I32(7),
        Value::Bool(true),
        Value::Object(vybe_runtime::heap::alloc(Object::new())),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
            Value::F64(-0.),
            Value::I32(0),
            Value::F64(f64::NAN),
            Value::String("x".into()),
        ]))),
    ] {
        for op in 0..2 {
            let mut results = vec![];
            for v in &mut ws {
                let f = function(v, op);
                results.push(
                    v.invoke(&f, &[arg.clone()])
                        .map(|x| values(&x))
                        .map_err(|e| e.message),
                );
            }
            assert_eq!(results[0], results[1]);
            checks += 2;
        }
    }
    output(
        "GATES",
        &serde_json::json!({"checks":checks,"exhaustive_ascii_pairs":16384,"ascii_rotations":128,"random_unicode":512}),
    );
    output("JS_CORPUS", &serde_json::json!({"queries":corpus}));
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    vybe_runtime::heap::enable_tracking();
    let mut rows = vec![];
    for (label, source, op, control) in [
        ("ascii-repeat-64k", "a".repeat(65536), 0, false),
        (
            "ascii-text-64k",
            "the quick brown fox jumps over the lazy dog ".repeat(1500),
            0,
            false,
        ),
        (
            "ascii-domain-64k",
            (0..128)
                .map(|i| i as u8 as char)
                .collect::<String>()
                .repeat(512),
            0,
            false,
        ),
        ("ascii-fromIterable", "abcdef".repeat(1024), 1, false),
        ("ascii-16", "a".repeat(16), 0, false),
        (
            "ascii-unique-32",
            (32..64).map(|i| i as u8 as char).collect::<String>(),
            0,
            false,
        ),
        ("short", "abc".into(), 0, false),
        ("empty", "".into(), 0, false),
        ("unicode-repeat", "Ω😀中".repeat(2048), 0, false),
        (
            "unicode-unique",
            (0x400..0x480)
                .map(|x| char::from_u32(x).unwrap())
                .collect::<String>(),
            0,
            false,
        ),
        ("control", "abcdef".repeat(1024), 0, true),
    ] {
        let args = [Value::String(source.clone().into())];
        let va = worker(false);
        let vb = worker(!control);
        let fs = [function(&va, op), function(&vb, op)];
        let mut vs = [va, vb];
        let heap_baseline = vybe_runtime::heap::snapshot();
        let mut since_collection = 0usize;
        let a = vs[0].invoke(&fs[0], &args).unwrap();
        let b = vs[1].invoke(&fs[1], &args).unwrap();
        assert_eq!(values(&a), values(&b));
        let capacities = [capacity(&a), capacity(&b)];
        drop((a, b));
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
                    std::hint::black_box(vs[side].invoke(&fs[side], &args).unwrap());
                    secs[side] += t.elapsed().as_secs_f64();
                    cpus[side] += thread_cpu_seconds() - c;
                    counts[side] += 1;
                    since_collection += 1;
                    if since_collection == 8 {
                        // Reclaim discarded result/iterator cycles outside the timed call.
                        vybe_runtime::heap::restore(&heap_baseline);
                        since_collection = 0;
                    }
                }
                turn += 1;
            }
            for side in 0..2 {
                wall[side].push(secs[side] / counts[side] as f64);
                cpu[side].push(cpus[side] / counts[side] as f64);
            }
            calls.push(counts);
        }
        vybe_runtime::heap::restore(&heap_baseline);
        let mut ratios: Vec<_> = wall[0].iter().zip(&wall[1]).map(|(a, b)| a / b).collect();
        ratios.sort_by(f64::total_cmp);
        let mut cr: Vec<_> = cpu[0].iter().zip(&cpu[1]).map(|(a, b)| a / b).collect();
        cr.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"label":label,"op":op,"source_bytes":source.len(),"capacities":capacities,"median_speedup":ratios[3],"median_thread_cpu_speedup":cr[3],"wall":wall,"cpu":cpu,"calls":calls}));
    }
    output(
        "MEASUREMENTS",
        &serde_json::json!({"pairs":7,"minimum_ms_per_side":100,"separate_vms":true,"cases":rows,"discarded_result_generation_collection_every":8,"collection_outside_timed_calls":true}),
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
