#![cfg(feature = "primitive-benchmarks")]
use std::sync::{Arc, Mutex};
use vybe_runtime::{
    VM, Value,
    value::{Object, ObjectKind},
};
fn output(n: &str, v: &serde_json::Value) {
    let d = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::create_dir_all(&d).unwrap();
    std::fs::write(
        d.join(format!("SPLICE_HOLES_{n}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
fn array(v: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(v)))
}
fn object(v: &Value) -> Arc<Mutex<Object>> {
    let Value::Object(o) = v else { panic!("array") };
    o.clone()
}
fn values(v: &Value) -> Vec<Value> {
    let o = object(v);
    let o = o.lock().unwrap();
    let ObjectKind::Array(a) = &o.kind else {
        panic!("array")
    };
    a.clone()
}
fn worker(fast: bool) -> (VM, Value) {
    let mut v = VM::new();
    vybe_platform_ecma::array::register(&mut v);
    vybe_platform_ecma::array::benchmark_array_splice_holes(&mut v, fast);
    let f = v.func_table[v
        .resolve_host_function_index("ecma:array", "splice")
        .unwrap()]
    .clone();
    (v, f)
}
fn make(len: usize, holes: Option<Vec<Value>>) -> (Value, Option<Value>) {
    let a = array((0..len).map(|i| Value::I32(i as i32)).collect());
    let marker = holes.map(array);
    let o = object(&a);
    let mut o = o.lock().unwrap();
    o.properties.insert("custom".into(), Value::I32(77));
    if let Some(m) = &marker {
        o.properties.insert("__holes".into(), m.clone());
    }
    drop(o);
    (a, marker)
}
fn raw_holes(a: &Value) -> Option<Vec<Value>> {
    object(a)
        .lock()
        .unwrap()
        .properties
        .get("__holes")
        .map(values)
}
fn snap(a: &Value) -> (Vec<Value>, Option<Vec<Value>>, Option<Value>, Option<Value>) {
    let o = object(a);
    let o = o.lock().unwrap();
    let ObjectKind::Array(v) = &o.kind else {
        panic!()
    };
    (
        v.clone(),
        o.properties.get("__holes").map(values),
        o.properties.get("length").cloned(),
        o.properties.get("custom").cloned(),
    )
}
fn visible(a: &Value) -> serde_json::Value {
    let data = values(a);
    let holes: Vec<usize> = raw_holes(a)
        .unwrap_or_default()
        .iter()
        .filter_map(|v| {
            let n = v.as_i64();
            (n >= 0).then_some(n as usize)
        })
        .filter(|n| *n < data.len())
        .collect();
    serde_json::json!({"values":data.iter().enumerate().map(|(i,v)|if holes.contains(&i){None}else{Some(v.as_i32())}).collect::<Vec<_>>(),"holes":holes})
}
#[test]
fn sparse_dense_and_zero_shift_match_frozen_provider_and_aliases() {
    let (mut a, fa) = worker(false);
    let (mut b, fb) = worker(true);
    let mut corpus = vec![];
    let mut checks = 0;
    for len in 0..=5 {
        for mask in 0..1usize << len {
            let holes: Vec<_> = (0..len)
                .filter(|i| mask & (1 << i) != 0)
                .map(|i| Value::I32(i as i32))
                .collect();
            for start in [-9, -1, 0, 1, 4, 99] {
                for delete in [-2, 0, 1, 3, 99] {
                    for insert in [0, 1, 4] {
                        let mut results = vec![];
                        for (v, f) in [(&mut a, &fa), (&mut b, &fb)] {
                            let (input, marker) = make(len, Some(holes.clone()));
                            let alias = input.clone();
                            let mut args =
                                vec![input.clone(), Value::I32(start), Value::I32(delete)];
                            args.extend((0..insert).map(|i| Value::I32(1000 + i)));
                            let removed = v.invoke(f, &args).unwrap();
                            assert!(v.last_exception.is_none());
                            assert!(Arc::ptr_eq(&object(&input), &object(&alias)));
                            assert!(!Arc::ptr_eq(&object(&input), &object(&removed)));
                            if object(&input)
                                .lock()
                                .unwrap()
                                .properties
                                .get("__holes")
                                .is_some()
                            {
                                let retained = object(&input)
                                    .lock()
                                    .unwrap()
                                    .properties
                                    .get("__holes")
                                    .unwrap()
                                    .clone();
                                assert!(Arc::ptr_eq(
                                    &object(marker.as_ref().unwrap()),
                                    &object(&retained)
                                ));
                            }
                            results.push((
                                snap(&input),
                                snap(&removed),
                                marker.as_ref().map(values),
                            ));
                            if results.len() == 2 {
                                corpus.push(serde_json::json!({"len":len,"holes":holes.iter().map(Value::as_i32).collect::<Vec<_>>(),"start":start,"delete":delete,"insert":insert,"after":visible(&input),"removed":visible(&removed)}));
                            }
                            checks += 1;
                        }
                        assert_eq!(
                            results[0], results[1],
                            "len={len},mask={mask},start={start},delete={delete},insert={insert}"
                        );
                    }
                }
            }
        }
    }
    output(
        "GATES",
        &serde_json::json!({"checks":checks,"cases":corpus.len()}),
    );
    output("CORPUS", &serde_json::json!({"queries":corpus}));
}
#[test]
fn malformed_metadata_omission_frozen_and_object_identity_match() {
    let mut checks = 0;
    for omit in [false, true] {
        for frozen in [false, true] {
            for holes in [
                None,
                Some(vec![]),
                Some(vec![
                    Value::I32(-1),
                    Value::I32(3),
                    Value::I32(3),
                    Value::I64(7),
                    Value::I64(1000),
                ]),
            ] {
                let mut out = vec![];
                let item = array(vec![Value::I32(42)]);
                for fast in [false, true] {
                    let (mut v, f) = worker(fast);
                    let (a, marker) = make(8, holes.clone());
                    if frozen {
                        object(&a)
                            .lock()
                            .unwrap()
                            .properties
                            .insert("__vybe_frozen".into(), Value::Bool(true));
                    }
                    let mut args = vec![a.clone(), Value::I32(1)];
                    if !omit {
                        args.push(Value::Undefined);
                        args.push(item.clone());
                    }
                    let invocation = v.invoke(&f, &args);
                    let thrown = invocation
                        .as_ref()
                        .err()
                        .map(|e| e.to_string().lines().next().unwrap().to_owned());
                    let removed = invocation.unwrap_or(Value::Undefined);
                    let error = v.last_exception.as_ref().map(|e| {
                        let o = object(e);
                        let o = o.lock().unwrap();
                        (
                            o.properties.get("name").cloned(),
                            o.properties.get("message").cloned(),
                        )
                    });
                    out.push((
                        snap(&a),
                        if matches!(removed, Value::Object(_)) {
                            Some(snap(&removed))
                        } else {
                            None
                        },
                        marker.as_ref().map(values),
                        error,
                        thrown,
                    ));
                    checks += 1;
                }
                assert_eq!(out[0], out[1]);
            }
        }
    }
    output("EDGE_GATES", &serde_json::json!({"checks":checks}));
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

fn restore(a: &Value, seed: &[Value], marker: &Value, holes: &[Value]) {
    let o = object(a);
    let mut o = o.lock().unwrap();
    let ObjectKind::Array(v) = &mut o.kind else {
        panic!()
    };
    v.clone_from(&seed.to_vec());
    o.properties
        .insert("length".into(), Value::F64(seed.len() as f64));
    if holes.is_empty() {
        o.properties.shift_remove("__holes");
    } else {
        object(marker).lock().unwrap().kind = ObjectKind::Array(holes.to_vec());
        o.properties.insert("__holes".into(), marker.clone());
    }
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (label, len, count, start, delete, insert, control) in [
        ("dense-delete", 32768, 0, 8192, 8192, 0, false),
        ("few-holes-delete", 32768, 4, 8192, 8192, 0, false),
        ("all-holes-replace", 8192, 8192, 2048, 1024, 1024, false),
        ("all-holes-whole", 8192, 8192, 0, 8192, 8192, false),
        ("all-holes-narrow", 8192, 8192, 2048, 1, 1, false),
        ("queue-pop", 32768, 0, 0, 1, 0, false),
        ("queue-insert", 32768, 0, 16384, 0, 1, false),
        ("dense-small", 8, 0, 2, 2, 2, false),
        ("control", 8192, 8192, 2048, 1024, 1024, true),
    ] {
        for repeats in [1] {
            let (mut va, fa) = worker(false);
            let (mut vb, fb) = worker(!control);
            let fs = [fa, fb];
            let mut vs = [va, vb];
            let seeds: Vec<_> = (0..len).map(Value::I32).collect();
            let template: Vec<_> = (0..count).map(|i| Value::I32(i * len / count)).collect();
            let arrays = [array(seeds.clone()), array(seeds.clone())];
            let markers = [array(template.clone()), array(template.clone())];
            let args = arrays
                .iter()
                .map(|a| {
                    let mut v = vec![a.clone(), Value::I32(start), Value::I32(delete)];
                    v.extend((0..insert).map(|i| Value::I32(100000 + i)));
                    v
                })
                .collect::<Vec<_>>();
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
                        restore(&arrays[side], &seeds, &markers[side], &template);
                        let c = thread_cpu_seconds();
                        let t = Instant::now();
                        std::hint::black_box(vs[side].invoke(&fs[side], &args[side]).unwrap());
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
