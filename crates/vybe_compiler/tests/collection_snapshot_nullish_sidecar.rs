//! Separate correctness foundation: nullish snapshot sidecar, actual public emitters.
use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::{bundle, platforms, sets};
use vybe_runtime::{
    Chunk, Op, VM, Value,
    value::{Object, ObjectKind},
};
fn object(o: Object) -> Value {
    Value::Object(vybe_runtime::heap::alloc(o))
}
fn set() -> Value {
    let mut o = Object::new();
    o.kind = ObjectKind::Set(Default::default());
    object(o)
}
fn map() -> Value {
    let mut o = Object::new();
    o.kind = ObjectKind::Map(Default::default());
    object(o)
}
fn same(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Object(a), Value::Object(b)) => Arc::ptr_eq(a, b),
        _ => a == b,
    }
}
fn prop(s: &Value) -> Value {
    let Value::Object(o) = s else { panic!("set") };
    o.lock()
        .unwrap()
        .properties
        .get(sets::SNAPSHOT_KEYS_KEY)
        .cloned()
        .unwrap_or(Value::Undefined)
}
fn put(s: &Value, key: &str, v: Value) {
    let Value::Object(o) = s else { panic!("set") };
    o.lock().unwrap().properties.insert(key.into(), v);
}
fn module(operation: u8) -> Vec<Chunk> {
    let mut e = Chunk::new("entry");
    e.emit_op_u16(Op::REF_FUNC, 1, 0);
    e.emit(0, 0);
    e.emit_op(Op::RETURN, 0);
    let mut b = Chunk::new("snapshot-owner");
    b.arity = 2;
    b.local_count = 2;
    for i in [0, 1] {
        b.emit_op_u16(Op::LOCAL_GET, i, 0);
    }
    let mut c = vec![e, b];
    match operation {
        0 => sets::emit_add_snapshot(&mut c, 1, 0),
        1 => sets::emit_has_snapshot(&mut c, 1, 0),
        2 => sets::emit_delete_snapshot(&mut c, 1, 0),
        _ => unreachable!(),
    }
    c[1].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut c);
    c
}
fn machine(operation: u8) -> (VM, Value) {
    let mut v = VM::new();
    platforms::register_platforms_all(&mut v);
    let f = v.run(module(operation)).unwrap();
    (v, f)
}
#[test]
fn absent_and_explicit_null_initialize_once_and_deduplicate() {
    for initial in [None, Some(Value::Undefined), Some(Value::Null)] {
        let s = set();
        if let Some(v) = initial {
            put(&s, sets::SNAPSHOT_KEYS_KEY, v);
        }
        let (mut v, f) = machine(0);
        assert_eq!(
            v.invoke(&f, &[s.clone(), Value::I32(4)]).unwrap(),
            Value::Bool(true)
        );
        let keys = prop(&s);
        assert!(
            matches!(&keys,Value::Object(o)if matches!(o.lock().unwrap().kind,ObjectKind::Map(_)))
        );
        assert_eq!(
            v.invoke(&f, &[s.clone(), Value::I32(4)]).unwrap(),
            Value::Bool(false)
        );
        assert!(same(&keys, &prop(&s)));
        let (mut v, f) = machine(2);
        assert_eq!(
            v.invoke(&f, &[s.clone(), Value::I32(4)]).unwrap(),
            Value::Bool(true)
        );
        assert_eq!(
            v.invoke(&f, &[s.clone(), Value::I32(4)]).unwrap(),
            Value::Bool(false)
        );
        assert!(same(&keys, &prop(&s)));
    }
}
#[test]
fn non_null_falsey_sidecars_keep_identity_and_membership_behavior() {
    for input in [
        Value::Bool(false),
        Value::I32(0),
        Value::String("".into()),
        Value::String("existing".into()),
        map(),
    ] {
        for op in [0, 1, 2] {
            let initial = if matches!(&input, Value::Object(_)) {
                map()
            } else {
                input.clone()
            };
            let s = set();
            put(&s, sets::SNAPSHOT_KEYS_KEY, initial.clone());
            let (mut v, f) = machine(op);
            let result = v.invoke(&f, &[s.clone(), Value::I32(4)]).unwrap();
            assert_eq!(result, Value::Bool(op == 0));
            assert!(
                same(&initial, &prop(&s)),
                "non-null sidecar replaced op={op}"
            );
            assert!(v.last_exception.is_none());
        }
    }
}
#[test]
fn real_getter_observed_once_before_membership_and_keeps_map() {
    {
        let (mut v, f) = machine(0);
        let trace = Arc::new(Mutex::new(Vec::new()));
        let t = trace.clone();
        let keys = map();
        let returned = keys.clone();
        v.register_host_fn(
            "test:snapshot",
            "get",
            Box::new(move |_, _| {
                t.lock().unwrap().push("get");
                returned.clone()
            }),
        );
        let idx = v
            .resolve_host_function_index("test:snapshot", "get")
            .unwrap();
        let s = set();
        put(&s, "__get___snapshot_keys", v.func_table[idx].clone());
        assert_eq!(
            v.invoke(&f, &[s.clone(), Value::I32(4)]).unwrap(),
            Value::Bool(true)
        );
        assert_eq!(
            v.invoke(&f, &[s.clone(), Value::I32(4)]).unwrap(),
            Value::Bool(false)
        );
        assert_eq!(*trace.lock().unwrap(), vec!["get", "get"]);
        assert_eq!(
            prop(&s),
            Value::Undefined,
            "existing getter not overwritten"
        );
        assert!(v.last_exception.is_none());
    }
}
#[test]
fn full_public_snapshot_native_validation() {
    {
        for operation in [0, 1, 2] {
            let bytes = vybe_platform_wasm::write_wasm(&module(operation));
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes)
                .unwrap_or_else(|e| panic!("operation={operation}: {e}"));
        }
    }
}
#[test]
fn throwing_getter_is_observed_once_and_exception_remains_visible() {
    let (mut v, f) = machine(1);
    let trace = Arc::new(Mutex::new(Vec::new()));
    let t = trace.clone();
    v.register_host_fn(
        "test:snapshot",
        "throw",
        Box::new(move |ctx, _| {
            t.lock().unwrap().push("get");
            ctx.throw_value(Value::String("sidecar-getter-exception".into()));
            Value::Bool(false)
        }),
    );
    let idx = v
        .resolve_host_function_index("test:snapshot", "throw")
        .unwrap();
    let s = set();
    put(&s, "__get___snapshot_keys", v.func_table[idx].clone());
    let r = v.invoke(&f, &[s.clone(), Value::I32(4)]);
    assert_eq!(*trace.lock().unwrap(), vec!["get"]);
    assert_eq!(prop(&s), Value::Undefined);
    assert!(
        r.is_err() || v.last_exception.is_some(),
        "getter exception remains observable"
    );
}
#[test]
fn actual_snapshot_constructor_seeds_membership_for_adopted_iterable() {
    let mut e = Chunk::new("entry");
    e.emit_op_u16(Op::REF_FUNC, 1, 0);
    e.emit(0, 0);
    e.emit_op(Op::RETURN, 0);
    let mut b = Chunk::new("snapshot-construct");
    b.arity = 1;
    b.local_count = 1;
    b.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let mut c = vec![e, b];
    sets::emit_from_iterable_snapshot(&mut c, 1, 0);
    c[1].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut c);
    let mut v = VM::new();
    platforms::register_platforms_all(&mut v);
    let f = v.run(c).unwrap();
    let a = object(Object::new_array(vec![
        Value::I32(1),
        Value::I32(1),
        Value::I32(2),
    ]));
    let copied = v.invoke(&f, &[a]).unwrap();
    let Value::Object(o) = &copied else {
        panic!("snapshot constructor")
    };
    assert!(
        matches!(&o.lock().unwrap().kind,ObjectKind::Set(s)if s.iter().cloned().collect::<Vec<_>>()==vec![Value::I32(1),Value::I32(2)])
    );
    assert!(
        matches!(prop(&copied),Value::Object(o)if matches!(o.lock().unwrap().kind,ObjectKind::Map(_)))
    );
    let (mut v, has) = machine(1);
    assert_eq!(
        v.invoke(&has, &[copied.clone(), Value::I32(1)]).unwrap(),
        Value::Bool(true)
    );
    assert_eq!(
        v.invoke(&has, &[copied, Value::I32(9)]).unwrap(),
        Value::Bool(false)
    );
}
