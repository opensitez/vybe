use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::{collections, platforms};
use vybe_runtime::{
    Chunk, Op, VM, Value,
    value::{Object, ObjectKind},
};
#[path = "fixtures/collection_stack_before.rs"]
mod before;
fn output(n: &str, v: &serde_json::Value) {
    let d = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap());
    std::fs::create_dir_all(&d).unwrap();
    std::fs::write(
        d.join(format!("COLLECTION_STACK_{n}.json")),
        serde_json::to_vec_pretty(v).unwrap(),
    )
    .unwrap();
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
fn program(fast: bool, kind: usize, reps: usize) -> Vec<Chunk> {
    program_scoped(fast, kind, reps, false)
}
fn program_scoped(fast: bool, kind: usize, reps: usize, prefix: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("collection-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = 1;
    c.param_count = 1;
    c.local_count = 1;
    let mut cs = vec![root, c];
    for i in 0..reps {
        if prefix {
            cs[1].emit_i32_const(17, 0);
        }
        cs[1].emit_op_u16(Op::LOCAL_GET, 0, 0);
        match (kind, fast) {
            (0, false) => before::emit_rank(&mut cs, 1, 0),
            (0, true) => collections::emit_rank(&mut cs, 1, 0),
            (_, false) => before::emit_pop(&mut cs, 1, 0),
            (_, true) => collections::emit_pop(&mut cs, 1, 0),
        }
        if prefix {
            cs[1].emit_op(Op::I32_ADD, 0);
        }
        if i + 1 < reps {
            cs[1].emit_op(Op::DROP, 0);
        }
    }
    cs[1].emit_op(Op::RETURN, 0);
    cs[0].imports = cs[1].imports.clone();
    cs
}
fn worker(cs: Vec<Chunk>) -> (VM, Value) {
    let mut v = VM::new();
    platforms::register_platforms_all(&mut v);
    let f = v.run(cs).unwrap();
    (v, f)
}
fn arr(xs: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(xs)))
}
fn map(xs: Vec<(Value, Value)>) -> Value {
    let mut o = Object::new();
    o.kind = ObjectKind::Map(xs.into_iter().collect());
    Value::Object(vybe_runtime::heap::alloc(o))
}
fn values(v: &Value) -> Vec<Value> {
    let Value::Object(o) = v else { return vec![] };
    let o = o.lock().unwrap();
    match &o.kind {
        ObjectKind::Array(xs) => xs.clone(),
        ObjectKind::Map(xs) => xs.values().cloned().collect(),
        _ => vec![],
    }
}
fn same(a: &Value, b: &Value) {
    match (a, b) {
        (Value::F64(a), Value::F64(b)) => assert_eq!(a.to_bits(), b.to_bits()),
        _ => assert_eq!(a, b),
    }
}
#[test]
fn rank_depth_and_pop_mutation_identity_and_native_validation() {
    let d = std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap())
        .join("collection-native");
    std::fs::create_dir_all(&d).unwrap();
    let mut sizes = vec![];
    for k in 0..2 {
        for fast in [false, true] {
            let cs = program(fast, k, 1);
            let b = native(&cs);
            std::fs::write(d.join(format!("{k}-{fast}.wasm")), &b).unwrap();
            sizes.push(serde_json::json!({"kind":k,"fast":fast,"body":cs[1].code.len(),"locals":cs[1].local_count,"native":b.len()}));
        }
    }
    let mut checks = 0;
    let mut ranks = vec![];
    for seed in [
        Value::Null,
        Value::Undefined,
        Value::I32(1),
        Value::String("x".into()),
        map(vec![]),
    ] {
        let mut x = seed;
        for depth in 0..6 {
            ranks.push((x.clone(), depth.min(3)));
            x = arr(vec![x, Value::I32(99)]);
        }
    }
    ranks.extend([
        (arr(vec![]), 1),
        (arr(vec![arr(vec![])]), 2),
        (arr(vec![arr(vec![arr(vec![])])]), 3),
    ]);
    for fast in [false, true] {
        let (mut v, f) = worker(program(fast, 0, 1));
        for (x, rank) in &ranks {
            assert_eq!(v.invoke(&f, &[x.clone()]).unwrap().as_i32(), *rank);
            checks += 1;
        }
    }
    let object = Value::Object(vybe_runtime::heap::alloc(Object::new()));
    for kind in 0..2 {
        for len in [0, 1, 2, 8] {
            for val in [
                Value::Null,
                Value::Undefined,
                Value::I32(7),
                Value::F64(f64::from_bits(0xfff8000000000042)),
                object.clone(),
            ] {
                let mut outcomes = vec![];
                for fast in [false, true] {
                    let x = if kind == 0 {
                        arr(vec![val.clone(); len])
                    } else {
                        map((0..len)
                            .map(|i| (Value::String(format!("k{i}").into()), val.clone()))
                            .collect())
                    };
                    let (mut v, f) = worker(program(fast, 1, 1));
                    let result = v.invoke(&f, &[x.clone()]).unwrap();
                    if len > 0 {
                        same(&result, &val);
                    } else if kind == 0 {
                        assert_eq!(result, Value::Undefined);
                    } else {
                        assert_eq!(result, Value::Null);
                    }
                    assert_eq!(values(&x).len(), len.saturating_sub(1));
                    outcomes.push(result);
                    checks += 1;
                }
                same(&outcomes[0], &outcomes[1]);
            }
        }
    }
    output("GATES", &serde_json::json!({"checks":checks,"sizes":sizes}));
}
#[test]
fn map_get_result_stays_rooted_across_delete_and_call_order_matches() {
    let mut traces = vec![];
    for fast in [false, true] {
        let object = vybe_runtime::heap::alloc(Object::new());
        let weak = Arc::downgrade(&object);
        let x = map(vec![(
            Value::String("key".into()),
            Value::Object(object.clone()),
        )]);
        drop(object);
        let (mut v, f) = worker(program(fast, 1, 1));
        let trace = Arc::new(Mutex::new(Vec::new()));
        let t = trace.clone();
        v.register_host_fn(
            "ecma:map",
            "keys",
            Box::new(move |_, args| {
                t.lock().unwrap().push("keys");
                let Value::Object(o) = &args[0] else { panic!() };
                let o = o.lock().unwrap();
                let ObjectKind::Map(m) = &o.kind else {
                    panic!()
                };
                arr(m.keys().cloned().collect())
            }),
        );
        let t = trace.clone();
        v.register_host_fn(
            "ecma:map",
            "get",
            Box::new(move |_, args| {
                t.lock().unwrap().push("get");
                let Value::Object(o) = &args[0] else { panic!() };
                let o = o.lock().unwrap();
                let ObjectKind::Map(m) = &o.kind else {
                    panic!()
                };
                m.get(&args[1]).cloned().unwrap()
            }),
        );
        let t = trace.clone();
        let w = weak.clone();
        v.register_host_fn(
            "ecma:map",
            "delete",
            Box::new(move |_, args| {
                t.lock().unwrap().push("delete");
                let Value::Object(o) = &args[0] else { panic!() };
                let mut o = o.lock().unwrap();
                let ObjectKind::Map(m) = &mut o.kind else {
                    panic!()
                };
                drop(m.shift_remove(&args[1]));
                assert!(
                    w.upgrade().is_some(),
                    "fetched object must remain a root after deletion"
                );
                Value::Bool(true)
            }),
        );
        let result = v.invoke(&f, &[x]).unwrap();
        let Value::Object(o) = &result else { panic!() };
        assert!(Arc::ptr_eq(o, &weak.upgrade().unwrap()));
        traces.push(trace.lock().unwrap().clone());
    }
    assert_eq!(traces[0], vec!["keys", "get", "delete"]);
    assert_eq!(traces[0], traces[1]);
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for repetitions in [1, 16] {
        for (kind, label, input, control) in [
            (0, "rank-scalar", Value::I32(7), false),
            (0, "rank-one", arr(vec![Value::I32(7)]), false),
            (
                0,
                "rank-three",
                arr(vec![arr(vec![arr(vec![Value::I32(7)])])]),
                false,
            ),
            (
                0,
                "rank-control",
                arr(vec![arr(vec![arr(vec![Value::I32(7)])])]),
                true,
            ),
            (1, "pop-array", arr(vec![Value::I32(7); repetitions]), false),
            (
                1,
                "pop-map",
                map((0..repetitions)
                    .map(|i| (Value::I32(i as i32), Value::I32(7)))
                    .collect()),
                false,
            ),
            (1, "pop-empty-map", map(vec![]), false),
            (
                1,
                "pop-control",
                map((0..repetitions)
                    .map(|i| (Value::I32(i as i32), Value::I32(7)))
                    .collect()),
                true,
            ),
        ] {
            let old = program(false, kind, repetitions);
            let new = program(!control, kind, repetitions);
            let sizes = serde_json::json!({"old_guest":old[1].code.len(),"new_guest":new[1].code.len(),"old_native":native(&old).len(),"new_native":native(&new).len(),"old_locals":old[1].local_count,"new_locals":new[1].local_count});
            let (va, fa) = worker(old);
            let (vb, fb) = worker(new);
            let mut vs = [va, vb];
            let fs = [fa, fb];
            let inputs = if kind == 0 {
                [input.clone(), input]
            } else {
                let Value::Object(o) = input else {
                    unreachable!()
                };
                let original = o.lock().unwrap().kind.clone();
                [
                    Value::Object(vybe_runtime::heap::alloc({
                        let mut o = Object::new();
                        o.kind = original.clone();
                        o
                    })),
                    Value::Object(vybe_runtime::heap::alloc({
                        let mut o = Object::new();
                        o.kind = original;
                        o
                    })),
                ]
            };
            let template = if kind == 1 {
                let Value::Object(o) = &inputs[0] else {
                    unreachable!()
                };
                Some(o.lock().unwrap().kind.clone())
            } else {
                None
            };
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
                        if let Some(template) = &template {
                            let Value::Object(o) = &inputs[side] else {
                                unreachable!()
                            };
                            o.lock().unwrap().kind = template.clone();
                        }
                        let c = thread_cpu_seconds();
                        let t = Instant::now();
                        std::hint::black_box(
                            vs[side].invoke(&fs[side], &[inputs[side].clone()]).unwrap(),
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
            let mut ratios: Vec<_> = wall[0].iter().zip(&wall[1]).map(|(a, b)| a / b).collect();
            ratios.sort_by(f64::total_cmp);
            let mut cr: Vec<_> = cpu[0].iter().zip(&cpu[1]).map(|(a, b)| a / b).collect();
            cr.sort_by(f64::total_cmp);
            rows.push(serde_json::json!({"label":label,"kind":kind,"repetitions":repetitions,"sizes":sizes,"median_speedup":ratios[3],"median_thread_cpu_speedup":cr[3],"wall":wall,"cpu":cpu,"calls":calls}));
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

#[test]
fn helpers_keep_a_lower_live_operand_in_the_callers_stack() {
    for fast in [false, true] {
        for (kind, input, expected) in [
            (0, arr(vec![arr(vec![arr(vec![])])]), 20),
            (1, arr(vec![Value::I32(7)]), 24),
            (
                1,
                map(vec![(Value::String("key".into()), Value::I32(7))]),
                24,
            ),
        ] {
            let cs = program_scoped(fast, kind, 1, true);
            native(&cs);
            let (mut v, f) = worker(cs);
            assert_eq!(v.invoke(&f, &[input]).unwrap().as_i32(), expected);
        }
    }
}
