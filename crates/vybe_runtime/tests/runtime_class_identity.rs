use vybe_runtime::value::Object;
use vybe_runtime::{Value, VM};

fn register_class(vm: &mut VM, name: &str, method: &str, result: i32) -> usize {
    vm.register_host_fn("identity-fixture", name, Box::new(move |_, _| Value::I32(result)));
    let host = vm.host_registry[&("identity-fixture".to_string(), name.to_string())];
    let mut definition = vm.type_registry.get(0).unwrap().clone();
    definition.name = name.to_string();
    let id = vm.type_registry.register(definition);
    vm.type_registry.add_host_method(id, method, host);
    id
}

#[test]
fn typed_receiver_cannot_borrow_another_classes_methods_from_metadata() {
    let mut vm = VM::new();
    let first = register_class(&mut vm, "Identity.First", "firstOnly", 11);
    register_class(&mut vm, "Identity.Second", "secondOnly", 22);
    let mut object = Object::new_typed(first);
    object.properties.insert("__type".into(), Value::String("Identity.Second".into()));
    let receiver = Value::Object(vybe_runtime::heap::alloc(object));

    let method = vm.resolve_property(&receiver, "firstOnly").unwrap();
    assert!(matches!(vm.invoke(&method, &[]).unwrap(), Value::I32(11)));
    assert!(matches!(vm.resolve_property(&receiver, "secondOnly").unwrap(), Value::Undefined));
}

#[test]
fn untyped_receiver_cannot_acquire_class_identity_from_metadata() {
    let mut vm = VM::new();
    register_class(&mut vm, "Identity.Registered", "registeredOnly", 33);
    for metadata in [Value::String("Identity.Registered".into()), Value::I32(1), Value::Null] {
        let mut object = Object::new();
        object.properties.insert("__type".into(), metadata);
        let receiver = Value::Object(vybe_runtime::heap::alloc(object));
        assert!(matches!(vm.resolve_property(&receiver, "registeredOnly").unwrap(), Value::Undefined));
    }
}

#[test]
fn removing_legacy_identity_preserves_explicit_instance_members() {
    let mut vm = VM::new();
    vm.register_host_fn("identity-fixture", "own", Box::new(|_, _| Value::I32(44)));
    let mut chunk = vybe_runtime::Chunk::new("identity-own-member");
    let own = chunk.add_import("identity-fixture", "own");
    chunk.emit_call(own, 0, 1);
    chunk.emit_op(vybe_runtime::Op::RETURN, 1);
    let value = vm.run(vec![chunk]).unwrap();
    let mut object = Object::new();
    object.properties.insert("__type".into(), Value::String("Identity.Missing".into()));
    object.properties.insert("ownMember".into(), value);
    let receiver = Value::Object(vybe_runtime::heap::alloc(object));
    assert!(matches!(vm.resolve_property(&receiver, "ownMember").unwrap(), Value::I32(44)));
}
