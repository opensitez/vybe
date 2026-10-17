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
fn arr(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}
fn obj(a: &Value) -> Arc<Mutex<Object>> {
    let Value::Object(o) = a else {
        panic!("object")
    };
    o.clone()
}
fn values(a: &Value) -> Vec<Value> {
    let o = obj(a);
    let o = o.lock().unwrap();
    let ObjectKind::Array(v) = &o.kind else {
        panic!("physical Array")
    };
    v.clone()
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
fn zero_single_bulk_prepend_preserves_order_alias_and_returned_length() {
    let mut v = vm();
    let f = host(&v, "unshift");
    for len in [0, 1, 2, 16, 128] {
        for count in [0, 1, 2, 16, 256] {
            let original: Vec<_> = (0..len).map(Value::I32).collect();
            let a = arr(original.clone());
            let alias = a.clone();
            let items: Vec<_> = (0..count).map(|i| Value::I32(1000 + i)).collect();
            let mut args = vec![a.clone()];
            args.extend(items.clone());
            assert_eq!(v.invoke(&f, &args).unwrap(), Value::I32(len + count));
            assert!(v.last_exception.is_none());
            let mut expected = items;
            expected.extend(original);
            assert_eq!(values(&a), expected);
            assert_eq!(values(&alias), expected);
            assert_eq!(
                obj(&a).lock().unwrap().properties.get("length"),
                Some(&Value::F64((len + count) as f64))
            );
        }
    }
}
#[test]
fn holes_shift_once_and_explicit_undefined_insertions_are_present() {
    let mut v = vm();
    let f = host(&v, "unshift");
    for count in [0, 1, 2, 8] {
        let a = arr((0..8).map(Value::I32).collect());
        vybe_platform_ecma::array::store_hole_indices(
            &mut obj(&a).lock().unwrap(),
            &BTreeSet::from([1, 3, 6]),
        );
        let mut args = vec![a.clone()];
        args.extend((0..count).map(|_| Value::Undefined));
        assert_eq!(v.invoke(&f, &args).unwrap(), Value::I32(8 + count));
        assert_eq!(
            holes(&a),
            [1, 3, 6]
                .into_iter()
                .map(|x| Value::I32(x + count))
                .collect::<Vec<_>>()
        );
        for i in 0..count {
            assert!(!vybe_platform_ecma::array::is_array_hole(
                &obj(&a).lock().unwrap(),
                i as usize
            ));
            assert_eq!(values(&a)[i as usize], Value::Undefined);
        }
    }
}
#[test]
fn frozen_presence_returns_original_length_without_mutation_or_exception() {
    for mark in [Value::Bool(false), Value::Null, Value::Undefined] {
        for count in [0, 1, 8] {
            let mut v = vm();
            let f = host(&v, "unshift");
            let a = arr(vec![Value::I32(1), Value::I32(2)]);
            obj(&a)
                .lock()
                .unwrap()
                .properties
                .insert("__vybe_frozen".into(), mark.clone());
            let mut args = vec![a.clone()];
            args.extend((0..count).map(Value::I32));
            assert_eq!(v.invoke(&f, &args).unwrap(), Value::I32(2));
            assert_eq!(values(&a), vec![Value::I32(1), Value::I32(2)]);
            assert!(v.last_exception.is_none());
        }
    }
}
#[test]
fn repeated_references_keep_identity_and_weak_lifetime_without_hidden_roots() {
    let mut v = vm();
    let f = host(&v, "unshift");
    let a = arr(vec![Value::I32(9)]);
    let item = vybe_runtime::heap::alloc(Object::new());
    let weak = Arc::downgrade(&item);
    let value = Value::Object(item.clone());
    drop(item);
    let args = vec![a.clone(), value.clone(), value.clone()];
    assert_eq!(v.invoke(&f, &args).unwrap(), Value::I32(3));
    let got = values(&a);
    assert!(Arc::ptr_eq(&obj(&got[0]), &obj(&got[1])));
    assert!(Arc::ptr_eq(&obj(&got[0]), &obj(&value)));
    drop(got);
    drop(value);
    drop(args);
    assert!(weak.upgrade().is_some());
    drop(a);
    assert!(weak.upgrade().is_none());
    let a = arr(vec![]);
    assert_eq!(
        v.invoke(&f, &[a.clone(), a.clone(), a.clone()]).unwrap(),
        Value::I32(2)
    );
    let got = values(&a);
    assert!(Arc::ptr_eq(&obj(&a), &obj(&got[0])));
    assert!(Arc::ptr_eq(&obj(&a), &obj(&got[1])));
    obj(&a).lock().unwrap().kind = ObjectKind::Array(vec![]);
}
#[test]
fn original_raw_backing_store_property_policy_and_nonarray_carriers_stay() {
    let mut v = vm();
    let calls = Arc::new(Mutex::new(0));
    let c = calls.clone();
    v.register_host_fn(
        "test:unshift",
        "accessor",
        Box::new(move |_, _| {
            *c.lock().unwrap() += 1;
            Value::I32(7)
        }),
    );
    let callback = v.func_table[v
        .resolve_host_function_index("test:unshift", "accessor")
        .unwrap()]
    .clone();
    assert_eq!(v.invoke(&callback, &[]).unwrap(), Value::I32(7));
    assert_eq!(*calls.lock().unwrap(), 1);
    *calls.lock().unwrap() = 0;
    let a = arr(vec![Value::I32(9)]);
    let prototype = Value::Object(vybe_runtime::heap::alloc(Object::new()));
    {
        let o = obj(&a);
        let mut o = o.lock().unwrap();
        o.properties.insert("__get_0".into(), callback.clone());
        o.properties.insert("__set_0".into(), callback);
        o.properties.insert("__proto__".into(), prototype.clone());
        o.properties
            .insert("tag".into(), Value::String("unchanged".into()));
        o.type_id = 77;
    }
    let f = host(&v, "unshift");
    assert_eq!(
        v.invoke(&f, &[a.clone(), Value::I32(1), Value::I32(2)])
            .unwrap(),
        Value::I32(3)
    );
    assert_eq!(*calls.lock().unwrap(), 0);
    let o = obj(&a);
    let o = o.lock().unwrap();
    assert_eq!(o.type_id, 77);
    assert!(Arc::ptr_eq(
        &obj(o.properties.get("__proto__").unwrap()),
        &obj(&prototype)
    ));
    assert_eq!(
        o.properties.get("tag"),
        Some(&Value::String("unchanged".into()))
    );
    drop(o);
    for nonarray in [
        Value::Null,
        Value::Undefined,
        Value::String("abc".into()),
        Value::Object(vybe_runtime::heap::alloc(Object::new())),
    ] {
        assert_eq!(
            v.invoke(&f, &[nonarray, Value::I32(1), Value::I32(2)])
                .unwrap(),
            Value::I32(0)
        );
        assert!(v.last_exception.is_none());
    }
}
