//! Isolated public-API regression gate for Object.assign inherited setters.
//! Run separately: it temporarily adds one property to the realm prototype.
use std::sync::{Arc, atomic::{AtomicUsize, Ordering}};
use vybe_runtime::{VM, Value, value::{Object, ObjectKind}};

fn function(vm: &VM, module: &str, name: &str) -> Value {
    let mut object = Object::new();
    object.kind = ObjectKind::HostFunction(vm.resolve_host_function_index(module, name).unwrap());
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn main() {
    let mode = std::env::args().nth(1).unwrap_or_else(|| "tracked-fresh".into());
    assert!(matches!(mode.as_str(), "tracked-fresh" | "untracked-fresh" | "tracked-overwrite"));
    let overwrite = mode == "tracked-overwrite";
    let mut vm = VM::new();
    vybe_platform_ecma::object::register(&mut vm);
    let calls = Arc::new(AtomicUsize::new(0));
    let observed = calls.clone();
    vm.register_free_fn("probe:assign-inherited", "setter", Box::new(move |_, _| {
        observed.fetch_add(1, Ordering::Relaxed);
        Value::Undefined
    }));
    let key = "vegaInheritedAssignmentProbe";
    let mut target = Object::new();
    target.properties.insert("__keys".into(),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(if overwrite {
            vec![Value::String(key.into())]
        } else { Vec::new() }))));
    if overwrite { target.properties.insert(key.into(), Value::I32(0)); }
    let target = Value::Object(vybe_runtime::heap::alloc(target));
    let prototype = vm.invoke(&function(&vm, "ecma:object", "getPrototypeOf"), &[target.clone()])
        .expect("Object.getPrototypeOf failed");
    let mut descriptor = Object::new();
    descriptor.properties.insert("set".into(), function(&vm, "probe:assign-inherited", "setter"));
    descriptor.properties.insert("configurable".into(), Value::Bool(true));
    vm.invoke(&function(&vm, "ecma:object", "defineProperty"), &[
        prototype, Value::String(key.into()),
        Value::Object(vybe_runtime::heap::alloc(descriptor)),
    ]).expect("Object.defineProperty failed");
    let mut source = Object::new();
    if mode != "untracked-fresh" {
        source.properties.insert("__keys".into(),
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![Value::String(key.into())]))));
    }
    source.properties.insert(key.into(), Value::I32(42));
    vm.invoke(&function(&vm, "ecma:object", "assign"), &[
        target.clone(), Value::Object(vybe_runtime::heap::alloc(source)),
    ]).expect("Object.assign failed");
    let setter_calls = calls.load(Ordering::Relaxed);
    let Value::Object(target) = target else { unreachable!() };
    let target = target.lock().unwrap();
    let own_property = target.properties.contains_key(key);
    println!("mode={mode}, inherited setter calls={setter_calls}, target has own property={own_property}");
    if overwrite {
        assert_eq!(setter_calls, 0, "An own data property must shadow the inherited setter");
        assert!(matches!(target.properties.get(key), Some(Value::I32(42))));
    } else {
        assert_eq!(setter_calls, 1, "Object.assign must invoke the inherited setter exactly once");
        assert!(!own_property, "Calling an inherited setter must not create an own data property");
    }
}
