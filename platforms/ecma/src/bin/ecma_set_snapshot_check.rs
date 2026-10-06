use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::VM;

fn function(vm: &VM, module: &str, name: &str) -> Value {
    let mut object = Object::new();
    object.kind = ObjectKind::HostFunction(vm.resolve_host_function_index(module, name).unwrap());
    Value::Object(vybe_runtime::heap::alloc(object))
}

fn main() {
    let mut vm = VM::new();
    vybe_platform_ecma::object::register(&mut vm);
    vybe_platform_ecma::reflect::register(&mut vm);
    let target = Arc::new(Mutex::new(Object::new()));
    let captured = target.clone();
    vm.register_free_fn("ecma:set-check", "setter", Box::new(move |ctx, args| {
        let user_args = ctx.user_args(args, 0);
        let mut target = captured.try_lock().expect("target lock held during setter");
        assert!(matches!(user_args.first(), Some(Value::I32(17))),
            "setter value displaced: raw={args:?}, normalized={user_args:?}");
        target.properties.insert("observed".into(), Value::I32(17));
        Value::Undefined
    }));
    let mut descriptor = Object::new();
    descriptor.properties.insert("set".into(), function(&vm, "ecma:set-check", "setter"));
    descriptor.properties.insert("configurable".into(), Value::Bool(true));
    let define = function(&vm, "ecma:object", "defineProperty");
    vm.invoke(&define, &[
        Value::Object(target.clone()),
        vybe_platform_ecma::keys::string_value("field"),
        Value::Object(vybe_runtime::heap::alloc(descriptor)),
    ]).unwrap();
    let set = function(&vm, "ecma:reflect", "set");
    let result = vm.invoke(&set, &[
        Value::Object(target.clone()),
        vybe_platform_ecma::keys::string_value("field"),
        Value::I32(17),
    ]).unwrap();
    assert!(matches!(result, Value::Bool(true)));
    assert!(matches!(target.lock().unwrap().properties.get("observed"), Some(Value::I32(17))));
    println!("Registered setter arguments and target-lock release passed");
}
