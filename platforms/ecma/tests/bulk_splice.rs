use std::{
    collections::BTreeSet,
    sync::{Arc, Mutex},
};
use vybe_runtime::{
    VM, Value,
    value::{Object, ObjectKind},
};
fn vm() -> VM {
    let mut v = VM::new();
    vybe_platform_ecma::array::register(&mut v);
    v
}
fn host(v: &VM, name: &str) -> Value {
    v.func_table[v
        .resolve_host_function_index("ecma:array", name)
        .expect("actual registered provider")]
    .clone()
}
fn arr(v: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(v)))
}
fn values(a: &Value) -> Vec<Value> {
    let Value::Object(o) = a else { panic!("array") };
    let o = o.lock().unwrap();
    let ObjectKind::Array(v) = &o.kind else {
        panic!("actual Array kind")
    };
    v.clone()
}
fn obj(a: &Value) -> Arc<Mutex<Object>> {
    let Value::Object(o) = a else {
        panic!("object")
    };
    o.clone()
}
fn holes(a: &Value) -> Vec<Value> {
    obj(a)
        .lock()
        .unwrap()
        .properties
        .get("__holes")
        .map(values)
        .unwrap_or_default()
}
#[test]
fn bulk_ranges_keep_original_array_alias_order_deleted_identity_and_length() {
    for len in [0, 1, 2, 7, 64] {
        for start in [-100, -1, 0, 1, 100] {
            for delete in [-1, 0, 1, 100] {
                for insert in [0, 1, 16] {
                    let mut v = vm();
                    let f = host(&v, "splice");
                    let original: Vec<Value> = (0..len).map(Value::I32).collect();
                    let a = arr(original.clone());
                    let alias = a.clone();
                    let replacement: Vec<Value> =
                        (0..insert).map(|i| Value::I32(1000 + i)).collect();
                    let mut args = vec![a.clone(), Value::I32(start), Value::I32(delete)];
                    args.extend(replacement.clone());
                    let out = v.invoke(&f, &args).unwrap();
                    assert!(v.last_exception.is_none());
                    let idx = if start < 0 {
                        (len + start).max(0) as usize
                    } else {
                        start.min(len) as usize
                    };
                    let end = (idx + delete.max(0) as usize).min(len as usize);
                    let mut expected = original[..idx].to_vec();
                    expected.extend(replacement);
                    expected.extend_from_slice(&original[end..]);
                    assert_eq!(values(&a), expected);
                    assert_eq!(values(&alias), expected);
                    assert_eq!(values(&out), original[idx..end]);
                    assert!(!Arc::ptr_eq(&obj(&a), &obj(&out)));
                    assert_eq!(
                        obj(&a).lock().unwrap().properties.get("length"),
                        Some(&Value::F64(expected.len() as f64))
                    );
                }
            }
        }
    }
}
#[test]
fn sparse_holes_metadata_and_omitted_vs_explicit_undefined_remain_exact() {
    let mut v = vm();
    let f = host(&v, "splice");
    let a = arr((0..8).map(Value::I32).collect());
    vybe_platform_ecma::array::store_hole_indices(
        &mut obj(&a).lock().unwrap(),
        &BTreeSet::from([1, 3, 6]),
    );
    let out = v
        .invoke(
            &f,
            &[a.clone(), Value::I32(2), Value::I32(3), Value::I32(99)],
        )
        .unwrap();
    assert_eq!(
        values(&out),
        vec![Value::I32(2), Value::I32(3), Value::I32(4)]
    );
    assert_eq!(holes(&out), vec![Value::I32(1)]);
    assert_eq!(holes(&a), vec![Value::I32(1), Value::I32(4)]);
    for omitted in [false, true] {
        let a = arr((0..4).map(Value::I32).collect());
        let mut args = vec![a.clone(), Value::I32(1)];
        if !omitted {
            args.push(Value::Undefined);
        }
        let out = v.invoke(&f, &args).unwrap();
        assert_eq!(values(&out).len(), if omitted { 3 } else { 0 });
        assert_eq!(values(&a).len(), if omitted { 1 } else { 4 });
    }
}
#[test]
fn repeated_object_inputs_deleted_values_and_weak_lifetimes_keep_identity() {
    let mut v = vm();
    let f = host(&v, "splice");
    let item = vybe_runtime::heap::alloc(Object::new());
    let weak = Arc::downgrade(&item);
    let a = arr(vec![Value::Object(item.clone()), Value::I32(1)]);
    drop(item);
    let out = v
        .invoke(&f, &[a.clone(), Value::I32(0), Value::I32(1)])
        .unwrap();
    assert!(weak.upgrade().is_some());
    let removed = values(&out);
    assert!(Arc::ptr_eq(&weak.upgrade().unwrap(), &obj(&removed[0])));
    drop(removed);
    drop(out);
    assert!(weak.upgrade().is_none());
    let item = Value::Object(vybe_runtime::heap::alloc(Object::new()));
    let out = v
        .invoke(
            &f,
            &[
                a.clone(),
                Value::I32(0),
                Value::I32(0),
                item.clone(),
                item.clone(),
                a.clone(),
            ],
        )
        .unwrap();
    assert!(values(&out).is_empty());
    let got = values(&a);
    assert!(Arc::ptr_eq(&obj(&got[0]), &obj(&item)));
    assert!(Arc::ptr_eq(&obj(&got[1]), &obj(&item)));
    assert!(Arc::ptr_eq(&obj(&got[2]), &obj(&a)));
    // Break the deliberately inserted cycle before leaving the test.
    obj(&a).lock().unwrap().kind = ObjectKind::Array(Vec::new());
}
#[test]
fn frozen_receiver_throws_before_mutation_and_scalar_queue_shapes_remain() {
    let mut v = vm();
    let f = host(&v, "splice");
    let a = arr(vec![Value::I32(1), Value::I32(2)]);
    obj(&a)
        .lock()
        .unwrap()
        .properties
        .insert("__vybe_frozen".into(), Value::Bool(false));
    let result = v.invoke(
        &f,
        &[a.clone(), Value::I32(0), Value::I32(1), Value::I32(9)],
    );
    assert!(
        result.is_err() || v.last_exception.is_some(),
        "real frozen marker guard must throw"
    );
    assert_eq!(values(&a), vec![Value::I32(1), Value::I32(2)]);
    v.last_exception = None;
    let a = arr(vec![Value::I32(1), Value::I32(3)]);
    let out = v
        .invoke(
            &f,
            &[a.clone(), Value::I32(1), Value::I32(0), Value::I32(2)],
        )
        .unwrap();
    assert!(values(&out).is_empty());
    let out = v
        .invoke(&f, &[a.clone(), Value::I32(0), Value::I32(1)])
        .unwrap();
    assert_eq!(values(&out), vec![Value::I32(1)]);
    assert_eq!(values(&a), vec![Value::I32(2), Value::I32(3)]);
}

#[test]
fn backing_store_operation_keeps_existing_raw_property_policy_and_metadata() {
    let mut v = vm();
    let calls = Arc::new(Mutex::new(0));
    let c = calls.clone();
    v.register_host_fn(
        "test:splice",
        "observable",
        Box::new(move |_, _| {
            *c.lock().unwrap() += 1;
            Value::I32(71)
        }),
    );
    let observer = v.func_table[v
        .resolve_host_function_index("test:splice", "observable")
        .unwrap()]
    .clone();
    // Positive registration/dispatch check, then zero calls is meaningful: this
    // existing provider uses the locked Vec and never dispatches index accessors.
    assert_eq!(v.invoke(&observer, &[]).unwrap(), Value::I32(71));
    assert_eq!(*calls.lock().unwrap(), 1);
    *calls.lock().unwrap() = 0;
    let a = arr((0..16).map(Value::I32).collect());
    let prototype = Value::Object(vybe_runtime::heap::alloc(Object::new()));
    {
        let object = obj(&a);
        let mut o = object.lock().unwrap();
        o.properties.insert("__get_0".into(), observer.clone());
        o.properties.insert("__set_0".into(), observer.clone());
        o.properties.insert("__proto__".into(), prototype.clone());
        o.properties
            .insert("annotation".into(), Value::String("preserved".into()));
        o.type_id = 123;
    }
    let f = host(&v, "splice");
    let out = v
        .invoke(
            &f,
            &[a.clone(), Value::I32(0), Value::I32(8), Value::I32(99)],
        )
        .unwrap();
    assert_eq!(*calls.lock().unwrap(), 0);
    let object = obj(&a);
    let o = object.lock().unwrap();
    assert_eq!(o.type_id, 123);
    assert_eq!(
        o.properties.get("annotation"),
        Some(&Value::String("preserved".into()))
    );
    assert!(Arc::ptr_eq(
        &obj(o.properties.get("__proto__").unwrap()),
        &obj(&prototype)
    ));
    assert_eq!(values(&out), (0..8).map(Value::I32).collect::<Vec<_>>());
}

#[test]
fn nonarray_carriers_keep_original_empty_deleted_result_and_no_mutation() {
    let mut v = vm();
    let f = host(&v, "splice");
    for input in [
        Value::Null,
        Value::Undefined,
        Value::String("abc".into()),
        Value::Object(vybe_runtime::heap::alloc(Object::new())),
    ] {
        let out = v
            .invoke(
                &f,
                &[input.clone(), Value::I32(0), Value::I32(2), Value::I32(7)],
            )
            .unwrap();
        assert!(v.last_exception.is_none());
        assert!(values(&out).is_empty());
        if let Value::Object(o) = input {
            assert!(!Arc::ptr_eq(&o, &obj(&out)));
            assert!(matches!(o.lock().unwrap().kind, ObjectKind::Ordinary));
        }
    }
}
