use super::*;
include!("concat_capacity_before.rs");
fn function(vm: &VM, name: &str) -> Value {
    let mut o = Object::new();
    o.kind = ObjectKind::HostFunction(vm.resolve_host_function_index("ecma:array", name).unwrap());
    Value::Object(vybe_runtime::heap::alloc(o))
}
fn values(value: Value) -> Vec<Value> {
    let Value::Object(o) = value else {
        panic!("array")
    };
    let o = o.lock().unwrap();
    let ObjectKind::Array(v) = &o.kind else {
        panic!("array")
    };
    v.clone()
}
fn runtime() -> (VM, [Value; 2]) {
    let mut vm = VM::new();
    register(&mut vm);
    register_before(&mut vm);
    let funcs = [function(&vm, "concat-before"), function(&vm, "concat")];
    (vm, funcs)
}
#[test]
fn registered_concat_keeps_shapes_identity_and_spread_policy() {
    let (mut vm, f) = runtime();
    for len in [0, 1, 255, 256, 4096] {
        let left = make_array((0..len).map(Value::I32).collect());
        for flag in [
            None,
            Some(Value::Bool(false)),
            Some(Value::Bool(true)),
            Some(Value::I32(1)),
            Some(Value::Null),
            Some(Value::Undefined),
        ] {
            let mut right = Object::new_array(vec![Value::I32(9), Value::Null]);
            if let Some(flag) = flag {
                right.properties.insert("isconcatspreadable".into(), flag);
            }
            let right = Value::Object(vybe_runtime::heap::alloc(right));
            for args in [
                [left.clone(), right],
                [left.clone(), left.clone()],
                [left.clone(), Value::String("abc".into())],
            ] {
                let old = values(vm.invoke(&f[0], &args).unwrap());
                let new = values(vm.invoke(&f[1], &args).unwrap());
                assert_eq!(old, new);
            }
        }
    }
}
#[test]
fn capacity_hint_is_optional_and_does_not_read_elements() {
    let mut array = Object::new_array(vec![Value::Null; 1024]);
    assert_eq!(concat_array_capacity_hint(&array), 1024);
    array
        .properties
        .insert("isconcatspreadable".into(), Value::Bool(false));
    assert_eq!(concat_array_capacity_hint(&array), 0);
    assert_eq!(concat_array_capacity_hint(&Object::new()), 0);
}
#[test]
#[ignore = "Seven paired opt0 registered concat batches"]
fn debug_pairs() {
    let (mut vm, f) = runtime();
    let mut rows = vec![];
    for len in [8, 4096, 65536] {
        let args = [
            make_array(vec![Value::I32(7); len]),
            make_array(vec![Value::I32(9); len]),
        ];
        let mut counts = [1, 1];
        let mut ratios = vec![];
        let mut raw = vec![];
        for pair in 0..7 {
            let mut ns = [0, 0];
            for side in [pair % 2, 1 - pair % 2] {
                loop {
                    let started = std::time::Instant::now();
                    for _ in 0..counts[side] {
                        std::hint::black_box(vm.invoke(&f[side], &args).unwrap());
                    }
                    ns[side] = started.elapsed().as_nanos();
                    if ns[side] >= 100_000_000 {
                        break;
                    }
                    counts[side] *= 2;
                }
            }
            ratios.push((ns[0] as f64 / counts[0] as f64) / (ns[1] as f64 / counts[1] as f64));
            raw.push(serde_json::json!({"counts":counts,"ns":ns}));
        }
        let mut sorted = ratios.clone();
        sorted.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"length":len,"median":sorted[3],"ratios":ratios,"raw":raw}));
    }
    if let Ok(path) = std::env::var("CONCAT_CAPACITY_MEASUREMENTS_PATH") {
        std::fs::write(path, serde_json::to_vec_pretty(&rows).unwrap()).unwrap();
    } else {
        eprintln!("{rows:?}");
    }
}
