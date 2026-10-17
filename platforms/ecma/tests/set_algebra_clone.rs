#![cfg(feature = "primitive-benchmarks")]
use std::sync::Arc;
use vybe_runtime::{
    VM, Value,
    value::{Object, ObjectKind},
};
const OPS: [&str; 4] = ["union", "intersection", "difference", "symmetricDifference"];
fn output(name: &str, v: &serde_json::Value) {
    let d = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::create_dir_all(&d).unwrap();
    std::fs::write(
        d.join(format!("SET_CLONE_{name}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
fn worker(fast: bool) -> VM {
    let mut v = VM::new();
    vybe_platform_ecma::set::register(&mut v);
    vybe_platform_ecma::set::benchmark_set_algebra_clone(&mut v, fast);
    v
}
fn function(v: &VM, op: usize) -> Value {
    v.func_table[v.resolve_host_function_index("ecma:set", OPS[op]).unwrap()].clone()
}
fn set(items: Vec<Value>) -> Value {
    vybe_platform_ecma::set::make_set(items.into_iter().collect())
}
fn values(v: &Value) -> Vec<Value> {
    let Value::Object(o) = v else { panic!("set") };
    let o = o.lock().unwrap();
    let ObjectKind::Set(s) = &o.kind else {
        panic!("Set backing")
    };
    s.iter().cloned().collect()
}
fn same_values(a: &[Value], b: &[Value]) {
    assert_eq!(a, b);
    for (x, y) in a.iter().zip(b) {
        match (x, y) {
            (Value::F64(x), Value::F64(y)) => assert_eq!(x.to_bits(), y.to_bits()),
            (Value::String(x), Value::String(y)) => assert!(Arc::ptr_eq(x, y)),
            (Value::Object(x), Value::Object(y)) => assert!(Arc::ptr_eq(x, y)),
            _ => assert_eq!(std::mem::discriminant(x), std::mem::discriminant(y)),
        }
    }
}
fn distinct(a: &Value, b: &Value) {
    if let (Value::Object(a), Value::Object(b)) = (a, b) {
        assert!(!Arc::ptr_eq(a, b));
    }
}
fn oracle(op: usize, a: &[Value], b: &[Value]) -> Vec<Value> {
    match op {
        0 => {
            let mut x = a.to_vec();
            for v in b {
                if !x.contains(v) {
                    x.push(v.clone());
                }
            }
            x
        }
        1 => a.iter().filter(|x| b.contains(x)).cloned().collect(),
        2 => a.iter().filter(|x| !b.contains(x)).cloned().collect(),
        _ => a
            .iter()
            .filter(|x| !b.contains(x))
            .chain(b.iter().filter(|x| !a.contains(x)))
            .cloned()
            .collect(),
    }
}
#[test]
fn fresh_ordered_results_identity_hashes_aliases_and_fallbacks() {
    let mut vs = [worker(false), worker(true)];
    let obj = Value::Object(vybe_runtime::heap::alloc(Object::new()));
    let other = Value::Object(vybe_runtime::heap::alloc(Object::new()));
    let pool = vec![
        Value::I32(1),
        Value::F64(1.),
        Value::F64(-0.),
        Value::I32(0),
        Value::F64(f64::NAN),
        Value::F64(f64::from_bits(0xfff8000000000042)),
        Value::String("long".repeat(100).into()),
        Value::String("x".into()),
        Value::Bool(true),
        Value::Undefined,
        Value::Null,
        obj.clone(),
        other.clone(),
    ];
    let mut checks = 0;
    for mask in 0..128 {
        let a = set(pool
            .iter()
            .enumerate()
            .filter(|(i, _)| (mask >> (i % 7)) & 1 != 0)
            .map(|(_, x)| x.clone())
            .collect());
        let b = set(pool
            .iter()
            .enumerate()
            .rev()
            .filter(|(i, _)| (mask >> ((i + 3) % 7)) & 1 != 0)
            .map(|(_, x)| x.clone())
            .collect());
        for right in [b.clone(), a.clone(), set(vec![])] {
            for op in 0..4 {
                let av = values(&a);
                let bv = values(&right);
                let want = oracle(op, &av, &bv);
                for v in &mut vs {
                    let f = function(v, op);
                    let r = v.invoke(&f, &[a.clone(), right.clone()]).unwrap();
                    same_values(&values(&r), &want);
                    distinct(&r, &a);
                    distinct(&r, &right);
                    same_values(&values(&a), &av);
                    same_values(&values(&right), &bv);
                    let Value::Object(o) = &r else { unreachable!() };
                    let o = o.lock().unwrap();
                    assert!(o.properties.contains_key("__proto__"));
                    assert_eq!(
                        o.properties.get("__type"),
                        Some(&Value::String("Set".into()))
                    );
                    let Value::Object(iterator) = o.properties.get("iterator").unwrap() else {
                        panic!("iterator")
                    };
                    let iterator = iterator.lock().unwrap();
                    let Value::Object(bound) = iterator.properties.get("__bound_args").unwrap()
                    else {
                        panic!("bound args")
                    };
                    let bound = bound.lock().unwrap();
                    let ObjectKind::Array(bound) = &bound.kind else {
                        panic!("args")
                    };
                    let (Value::Object(receiver), Value::Object(result)) = (&bound[0], &r) else {
                        panic!("receiver")
                    };
                    assert!(Arc::ptr_eq(receiver, result));
                    checks += 1;
                }
            }
        }
    }
    for bad in [
        Value::Undefined,
        Value::Null,
        Value::I32(1),
        Value::Object(vybe_runtime::heap::alloc(Object::new())),
    ] {
        for op in 0..4 {
            for args in [
                vec![],
                vec![bad.clone()],
                vec![bad.clone(), set(vec![obj.clone(), Value::I32(2)])],
                vec![set(vec![obj.clone()]), bad.clone()],
            ] {
                let mut outcomes = vec![];
                for v in &mut vs {
                    let f = function(v, op);
                    outcomes.push(
                        v.invoke(&f, &args)
                            .map(|x| values(&x))
                            .map_err(|e| e.message),
                    );
                }
                assert_eq!(outcomes[0], outcomes[1]);
                checks += 2;
            }
        }
    }
    // Mutating a cloned result changes no source membership or object identity.
    let a = set(vec![obj.clone(), Value::I32(1)]);
    let mut v = worker(true);
    let r = v
        .invoke(&function(&v, 0), &[a.clone(), set(vec![])])
        .unwrap();
    let Value::Object(o) = &r else { unreachable!() };
    {
        let mut o = o.lock().unwrap();
        let ObjectKind::Set(s) = &mut o.kind else {
            unreachable!()
        };
        s.shift_remove(&Value::I32(1));
        s.insert(Value::I32(3));
    }
    assert_eq!(values(&a), vec![obj.clone(), Value::I32(1)]);
    assert_eq!(values(&r), vec![obj, Value::I32(3)]);
    output(
        "GATES",
        &serde_json::json!({"checks":checks,"source_alias_and_result_mutation":true}),
    );
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    vybe_runtime::heap::enable_tracking();
    let mut rows = vec![];
    for (op, label, count, width, alias, empty, control) in [
        (0, "union-numbers", 256, 0, false, false, false),
        (0, "union-string-first", 128, 2048, false, true, false),
        (0, "union-overlap", 128, 2048, false, false, false),
        (0, "union-tiny", 2, 0, false, true, false),
        (0, "control", 128, 2048, false, true, true),
        (1, "intersection-self", 128, 2048, true, false, false),
        (1, "intersection-distinct", 128, 0, false, false, false),
        (2, "difference-self", 128, 2048, true, false, false),
        (2, "difference-empty", 128, 2048, false, true, false),
        (3, "symmetric-self", 128, 2048, true, false, false),
        (3, "symmetric-empty", 128, 2048, false, true, false),
    ] {
        let items = |shift: i32| -> Vec<Value> {
            (0..count)
                .map(|i| {
                    if width == 0 {
                        Value::I32(i + shift)
                    } else {
                        Value::String(format!("{}-{}", "x".repeat(width), i + shift).into())
                    }
                })
                .collect()
        };
        let a = set(items(0));
        let b = if alias {
            a.clone()
        } else if empty {
            set(vec![])
        } else {
            set(items(count / 2))
        };
        let args = [a, b];
        let va = worker(false);
        let vb = worker(!control);
        let fs = [function(&va, op), function(&vb, op)];
        let mut vs = [va, vb];
        let heap_baseline = vybe_runtime::heap::snapshot();
        let mut since_collection = 0usize;
        let expected = values(&vs[0].invoke(&fs[0], &args).unwrap());
        assert_eq!(values(&vs[1].invoke(&fs[1], &args).unwrap()), expected);
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
                    if since_collection == 128 {
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
        rows.push(serde_json::json!({"label":label,"op":op,"count":count,"width":width,"alias":alias,"empty":empty,"median_speedup":ratios[3],"median_thread_cpu_speedup":cr[3],"wall":wall,"cpu":cpu,"calls":calls}));
    }
    output(
        "MEASUREMENTS",
        &serde_json::json!({"pairs":7,"minimum_ms_per_side":100,"separate_vms":true,"cases":rows,"discarded_result_generation_collection_every":128,"collection_outside_timed_calls":true}),
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
fn encoded(v: &Value) -> serde_json::Value {
    match v {
        Value::I32(x) => serde_json::json!(["number", *x]),
        Value::F64(x) => serde_json::json!([
            "number",
            if x.is_nan() {
                serde_json::json!("NaN")
            } else if x.is_infinite() {
                serde_json::json!(if x.is_sign_negative() {
                    "-Infinity"
                } else {
                    "Infinity"
                })
            } else {
                serde_json::json!(if *x == 0. { 0. } else { *x })
            }
        ]),
        Value::String(x) => serde_json::json!(["string", x.as_ref()]),
        Value::Bool(x) => serde_json::json!(["boolean", x]),
        Value::Null => serde_json::json!(["null"]),
        Value::Undefined => serde_json::json!(["undefined"]),
        Value::Object(x) => serde_json::json!(["object", format!("{:p}", Arc::as_ptr(x))]),
        _ => panic!("corpus key"),
    }
}
#[test]
fn actual_adapter_results_for_js_oracle() {
    let mut vs = [worker(false), worker(true)];
    let pool = vec![
        Value::I32(0),
        Value::F64(-0.),
        Value::I32(1),
        Value::F64(1.),
        Value::F64(f64::NAN),
        Value::F64(f64::INFINITY),
        Value::F64(f64::NEG_INFINITY),
        Value::String("Ω😀".into()),
        Value::String("".into()),
        Value::String("x".repeat(200).into()),
        Value::Bool(false),
        Value::Bool(true),
        Value::Undefined,
        Value::Null,
        Value::Object(vybe_runtime::heap::alloc(Object::new())),
        Value::Object(vybe_runtime::heap::alloc(Object::new())),
    ];
    let mut state = 0x123456789abcdef0u64;
    let mut rows = vec![];
    for case in 0..512 {
        let mut next = || {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            state
        };
        let a = set((0..next() % 24)
            .map(|_| pool[next() as usize % pool.len()].clone())
            .collect());
        let b = if case % 7 == 0 {
            a.clone()
        } else {
            set((0..next() % 24)
                .map(|_| pool[next() as usize % pool.len()].clone())
                .collect())
        };
        for op in 0..4 {
            let mut results = vec![];
            for v in &mut vs {
                let f = function(v, op);
                results.push(values(&v.invoke(&f, &[a.clone(), b.clone()]).unwrap()));
            }
            same_values(&results[0], &results[1]);
            rows.push(serde_json::json!({"op":OPS[op],"a":values(&a).iter().map(encoded).collect::<Vec<_>>(),"b":values(&b).iter().map(encoded).collect::<Vec<_>>(),"result":results[1].iter().map(encoded).collect::<Vec<_>>()}));
        }
    }
    output("JS_CORPUS", &serde_json::json!({"queries":rows}));
}
