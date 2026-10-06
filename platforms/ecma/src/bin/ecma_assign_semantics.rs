//! Registered-adapter semantic gates for future Object.assign fast paths.
use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Value, VM};

fn object() -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new())))
}

fn callable(vm: &VM, module: &str, name: &str) -> Value {
    let index = vm.resolve_host_function_index(module, name).expect("registered gate function");
    let mut function = Object::new();
    function.kind = ObjectKind::HostFunction(index);
    Value::Object(Arc::new(Mutex::new(function)))
}

fn put(object: &Value, name: &str, value: Value) {
    let Value::Object(object) = object else { panic!("object fixture") };
    object.lock().unwrap().properties.insert(name.into(), value);
}

fn get(object: &Value, name: &str) -> Option<Value> {
    let Value::Object(object) = object else { panic!("object fixture") };
    object.lock().unwrap().properties.get(name).cloned()
}

fn key(name: &str) -> Value {
    Value::String(Arc::from(name))
}

fn descriptor(entries: &[(&str, Value)]) -> Value {
    let result = object();
    for (name, value) in entries {
        put(&result, name, value.clone());
    }
    result
}

fn main() {
    let mut vm = VM::new();
    vybe_platform_ecma::object::register(&mut vm);
    let assign = callable(&vm, "ecma:object", "assign");
    let define = callable(&vm, "ecma:object", "defineProperty");
    let prevent = callable(&vm, "ecma:object", "preventExtensions");
    let events = Arc::new(Mutex::new(Vec::<i32>::new()));

    let source = object();
    let target = object();
    let source_for_getter = source.clone();
    let getter_events = events.clone();
    vm.register_host_fn("ecma:assign-gate", "getter", Box::new(move |_, _| {
        getter_events.lock().unwrap().push(1);
        put(&source_for_getter, "second", Value::I32(42));
        Value::I32(7)
    }));
    let setter_events = events.clone();
    vm.register_host_fn("ecma:assign-gate", "setter", Box::new(move |_, _| {
        setter_events.lock().unwrap().push(2);
        Value::Undefined
    }));
    let getter = callable(&vm, "ecma:assign-gate", "getter");
    let setter = callable(&vm, "ecma:assign-gate", "setter");
    vm.invoke(&define, &[source.clone(), key("first"), descriptor(&[
        ("get", getter), ("enumerable", Value::Bool(true)), ("configurable", Value::Bool(true)),
    ])]).expect("define source getter");
    put(&source, "second", Value::I32(2));
    vm.invoke(&define, &[target.clone(), key("first"), descriptor(&[
        ("set", setter), ("enumerable", Value::Bool(true)), ("configurable", Value::Bool(true)),
    ])]).expect("define target setter");
    let result = vm.invoke(&assign, &[target.clone(), source.clone()]).expect("getter/setter assignment");
    let (Value::Object(result), Value::Object(target_object)) = (&result, &target) else { panic!("assign identity") };
    assert!(Arc::ptr_eq(result, target_object));
    assert_eq!(*events.lock().unwrap(), vec![1, 2]);
    assert!(matches!(get(&target, "second"), Some(Value::I32(42))));
    println!("getter/setter ordering and live source reads passed");

    let target = object();
    vm.invoke(&define, &[target.clone(), key("first"), descriptor(&[
        ("value", Value::I32(1)), ("writable", Value::Bool(false)),
        ("enumerable", Value::Bool(true)), ("configurable", Value::Bool(true)),
    ])]).expect("define readonly field");
    let source = object();
    put(&source, "first", Value::I32(2));
    put(&source, "second", Value::I32(3));
    assert!(vm.invoke(&assign, &[target.clone(), source]).is_err(), "readonly assignment must throw");
    assert!(matches!(get(&target, "first"), Some(Value::I32(1))));
    assert!(get(&target, "second").is_none(), "assignment must stop at first failure");
    println!("readonly failure and no later writes passed");

    let target = object();
    put(&target, "first", Value::I32(1));
    vm.invoke(&prevent, std::slice::from_ref(&target)).expect("prevent extensions");
    let source = object();
    put(&source, "first", Value::I32(2));
    put(&source, "second", Value::I32(3));
    assert!(vm.invoke(&assign, &[target.clone(), source]).is_err(), "new field on nonextensible target must throw");
    assert!(matches!(get(&target, "first"), Some(Value::I32(2))), "preceding write must remain");
    assert!(get(&target, "second").is_none());
    println!("nonextensible target and partial-write order passed");
}
