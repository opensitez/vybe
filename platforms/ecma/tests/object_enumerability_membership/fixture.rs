use std::sync::{Arc, Mutex};
use vybe_runtime::{
    VM, Value,
    value::{Object, ObjectKind},
};
pub fn vm() -> VM {
    let mut v = VM::new();
    vybe_platform_ecma::array::register(&mut v);
    vybe_platform_ecma::error::register(&mut v);
    vybe_platform_ecma::boolean::register(&mut v);
    vybe_platform_ecma::object::register(&mut v);
    vybe_platform_ecma::proxy::register(&mut v);
    vybe_platform_ecma::symbol::register(&mut v);
    v
}
pub fn host(v: &VM, module: &str, name: &str) -> Value {
    v.func_table[v
        .resolve_host_function_index(module, name)
        .expect("actual registered provider")]
    .clone()
}
pub fn object(fields: impl IntoIterator<Item = (String, Value)>) -> Value {
    let mut o = Object::new();
    o.properties.extend(fields);
    Value::Object(vybe_runtime::heap::alloc(o))
}
pub fn obj(v: &Value) -> Arc<Mutex<Object>> {
    let Value::Object(o) = v else {
        panic!("Object")
    };
    o.clone()
}
pub fn array(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}
pub fn values(v: &Value) -> Vec<Value> {
    let o = obj(v);
    let o = o.lock().unwrap();
    let ObjectKind::Array(a) = &o.kind else {
        panic!("actual Array")
    };
    a.clone()
}
pub fn strings(v: &Value) -> Vec<String> {
    values(v)
        .iter()
        .map(|x| match x {
            Value::String(s) => s.to_string(),
            _ => panic!("actual string key"),
        })
        .collect()
}
pub fn define(v: &mut VM, source: &Value, key: Value, descriptor: Value) {
    let f = host(v, "ecma:object", "defineProperty");
    let result = v.invoke(&f, &[source.clone(), key, descriptor]).unwrap();
    assert!(v.last_exception.is_none());
    assert!(Arc::ptr_eq(&obj(&result), &obj(source)));
}
pub fn data_descriptor(value: Value, enumerable: bool) -> Value {
    object([
        ("value".into(), value),
        ("enumerable".into(), Value::Bool(enumerable)),
        ("configurable".into(), Value::Bool(true)),
        ("writable".into(), Value::Bool(true)),
    ])
}
pub fn fixture(v: &mut VM, n: usize, h: usize) -> Value {
    let source = object([]);
    for i in 0..n {
        define(
            v,
            &source,
            Value::String(format!("k{i:05}").into()),
            data_descriptor(Value::I32(i as i32), i >= h),
        );
    }
    source
}
pub fn key_list(n: usize, h: usize) -> Vec<String> {
    (h..n).map(|i| format!("k{i:05}")).collect()
}
pub fn marker_owners(source: &Value) -> Option<(usize, usize)> {
    let source = obj(source);
    let source = source.lock().unwrap();
    let Value::Object(marker) = source.properties.get("__nonenum")? else {
        return None;
    };
    Some((Arc::strong_count(marker), Arc::weak_count(marker)))
}
