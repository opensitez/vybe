#![cfg(feature = "primitive-benchmarks")]
use std::sync::Arc;
use vybe_runtime::{
    VM, Value,
    value::{Object, ObjectKind, TypedElemKind},
};
fn output(name: &str, v: &serde_json::Value) {
    let d = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::create_dir_all(&d).unwrap();
    std::fs::write(
        d.join(format!("FILL_HOLES_{name}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
}
fn worker(fast: bool) -> VM {
    let mut v = VM::new();
    vybe_platform_ecma::array::register(&mut v);
    vybe_platform_ecma::array::benchmark_array_fill_holes(&mut v, fast);
    v
}
fn function(v: &VM) -> Value {
    v.func_table[v.resolve_host_function_index("ecma:array", "fill").unwrap()].clone()
}
fn array(values: Vec<Value>, holes: Option<Vec<Value>>) -> Value {
    let mut o = Object::new_array(values);
    o.properties.insert("custom".into(), Value::I32(77));
    if let Some(h) = holes {
        o.properties.insert(
            "__holes".into(),
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(h))),
        );
    }
    Value::Object(vybe_runtime::heap::alloc(o))
}
fn values(v: &Value) -> Vec<Value> {
    let Value::Object(o) = v else { panic!("array") };
    let o = o.lock().unwrap();
    let ObjectKind::Array(a) = &o.kind else {
        panic!("array")
    };
    a.clone()
}
fn holes(v: &Value) -> Option<Vec<Value>> {
    let Value::Object(o) = v else { panic!("array") };
    let o = o.lock().unwrap();
    o.properties.get("__holes").map(values)
}
fn same(a: &[Value], b: &[Value]) {
    assert_eq!(a, b);
    for (x, y) in a.iter().zip(b) {
        match (x, y) {
            (Value::F64(x), Value::F64(y)) => assert_eq!(x.to_bits(), y.to_bits()),
            (Value::Object(x), Value::Object(y)) => assert!(Arc::ptr_eq(x, y)),
            _ => {}
        }
    }
}
fn receiver(a: &Value, b: &Value) {
    let (Value::Object(a), Value::Object(b)) = (a, b) else {
        panic!("self")
    };
    assert!(Arc::ptr_eq(a, b));
}
#[test]
fn ranges_metadata_values_identity_and_typed_fallbacks() {
    let mut vs = [worker(false), worker(true)];
    let object = Value::Object(vybe_runtime::heap::alloc(Object::new()));
    let mut checks = 0;
    let mut corpus = vec![];
    for len in [0, 1, 2, 8, 31] {
        for mode in 0..6 {
            let h = match mode {
                0 => None,
                1 => Some(vec![]),
                2 => Some((0..len).map(Value::I32).collect()),
                3 => Some((0..len).step_by(3).map(Value::I32).collect()),
                4 => Some(vec![
                    Value::I32(0),
                    Value::I32(0),
                    Value::I32(-1),
                    Value::F64(1.),
                    Value::I64(2),
                    Value::I32(len + 10),
                ]),
                _ => Some(vec![Value::I32(0), Value::I32(len + 1)]),
            };
            for (start, end) in [
                (-100, 100),
                (-1, 100),
                (0, 0),
                (0, 100),
                (2, 1),
                (1, 3),
                (i32::MIN, i32::MAX),
                (0, -1),
                (2, -2),
                (100, 100),
            ] {
                for val in [
                    Value::I32(7),
                    Value::F64(f64::from_bits(0xfff8000000000042)),
                    object.clone(),
                ] {
                    let mut results = vec![];
                    for v in &mut vs {
                        let a = array((0..len).map(Value::I32).collect(), h.clone());
                        let old_hole = if let Value::Object(o) = &a {
                            o.lock().unwrap().properties.get("__holes").cloned()
                        } else {
                            None
                        };
                        let args = [a.clone(), val.clone(), Value::I32(start), Value::I32(end)];
                        let r = v.invoke(&function(v), &args).unwrap();
                        receiver(&a, &r);
                        let Value::Object(o) = &a else { unreachable!() };
                        {
                            let o = o.lock().unwrap();
                            assert_eq!(o.properties.get("custom"), Some(&Value::I32(77)));
                            if let (Some(old), Some(now)) = (&old_hole, o.properties.get("__holes"))
                            {
                                receiver(old, now);
                            }
                        }
                        results.push((values(&a), holes(&a)));
                        checks += 1;
                    }
                    same(&results[0].0, &results[1].0);
                    assert_eq!(results[0].1, results[1].1);
                    if mode < 4 && matches!(val, Value::I32(_)) {
                        corpus.push(serde_json::json!({"length":len,"holes":h.as_ref().map(|x|x.iter().map(Value::as_i32).collect::<Vec<_>>()).unwrap_or_default(),"start":start,"end":end,"value":7,"result":results[1].0.iter().map(Value::as_i32).collect::<Vec<_>>(),"remaining":results[1].1.as_ref().map(|x|x.iter().map(Value::as_i32).collect::<Vec<_>>()).unwrap_or_default()}));
                    }
                }
            }
        }
    }
    for end in [
        None,
        Some(Value::Undefined),
        Some(Value::Null),
        Some(Value::F64(f64::NAN)),
        Some(Value::String("2".into())),
    ] {
        let mut out = vec![];
        for v in &mut vs {
            let a = array(vec![Value::I32(0); 5], Some(vec![Value::I32(1)]));
            let mut args = vec![a.clone(), Value::I32(9), Value::I32(-2)];
            if let Some(e) = &end {
                args.push(e.clone());
            }
            v.invoke(&function(v), &args).unwrap();
            out.push((values(&a), holes(&a)));
            checks += 1;
        }
        assert_eq!(out[0], out[1]);
    }
    for kind in [TypedElemKind::U8, TypedElemKind::F64] {
        let mut out = vec![];
        for v in &mut vs {
            let a = vybe_platform_ecma::typedarray::new_typed_array(kind, 16);
            v.invoke(
                &function(v),
                &[a.clone(), Value::I32(7), Value::I32(-10), Value::I32(-1)],
            )
            .unwrap();
            let Value::Object(o) = a else { unreachable!() };
            let o = o.lock().unwrap();
            let ObjectKind::TypedArray(ta) = &o.kind else {
                unreachable!()
            };
            out.push(
                (0..16)
                    .map(|i| vybe_platform_ecma::typedarray::read_element(ta, i))
                    .collect::<Vec<_>>(),
            );
            checks += 1;
        }
        assert_eq!(out[0], out[1]);
    }
    output("GATES", &serde_json::json!({"checks":checks}));
    output("JS_CORPUS", &serde_json::json!({"queries":corpus}));
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (label, len, count, start, end, objects, control) in [
        ("dense-32k", 32768, 0, 0, 32768, false, false),
        ("few-holes-32k", 32768, 4, 0, 32768, false, false),
        ("all-holes-32k", 32768, 32768, 0, 32768, false, false),
        ("many-holes-narrow", 32768, 32768, 10, 11, false, false),
        ("empty-range", 32768, 4, 20, 10, false, false),
        ("object-values", 32768, 0, 0, 32768, true, false),
        ("dense-small", 8, 0, 0, 8, false, false),
        ("control", 32768, 0, 0, 32768, false, true),
    ] {
        let val = if objects {
            Value::Object(vybe_runtime::heap::alloc(Object::new()))
        } else {
            Value::I32(7)
        };
        let template: Vec<_> = (0..count).map(|i| Value::I32(i * len / count)).collect();
        let arrays = [
            array(vec![val.clone(); len as usize], None),
            array(vec![val.clone(); len as usize], None),
        ];
        let markers = [array(template.clone(), None), array(template.clone(), None)];
        let va = worker(false);
        let vb = worker(!control);
        let fs = [function(&va), function(&vb)];
        let mut vs = [va, vb];
        let args = [
            vec![
                arrays[0].clone(),
                val.clone(),
                Value::I32(start),
                Value::I32(end),
            ],
            vec![arrays[1].clone(), val, Value::I32(start), Value::I32(end)],
        ];
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
                    // Restore the original sparse metadata before every timed call.
                    if count > 0 {
                        let Value::Object(marker) = &markers[side] else {
                            unreachable!()
                        };
                        marker.lock().unwrap().kind = ObjectKind::Array(template.clone());
                        let Value::Object(a) = &arrays[side] else {
                            unreachable!()
                        };
                        a.lock()
                            .unwrap()
                            .properties
                            .insert("__holes".into(), markers[side].clone());
                    }
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
        let mut ratios: Vec<_> = wall[0].iter().zip(&wall[1]).map(|(a, b)| a / b).collect();
        ratios.sort_by(f64::total_cmp);
        let mut cr: Vec<_> = cpu[0].iter().zip(&cpu[1]).map(|(a, b)| a / b).collect();
        cr.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"label":label,"length":len,"holes":count,"start":start,"end":end,"objects":objects,"median_speedup":ratios[3],"median_thread_cpu_speedup":cr[3],"wall":wall,"cpu":cpu,"calls":calls}));
    }
    output(
        "MEASUREMENTS",
        &serde_json::json!({"pairs":7,"minimum_ms_per_side":100,"separate_vms":true,"cases":rows,"original_holes_restored_outside_every_timed_call":true}),
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
