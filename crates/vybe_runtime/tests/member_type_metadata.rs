use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, Value, VM};

#[test]
fn member_resolution_uses_runtime_type_not_writable_type_metadata() {
    for registered in ["Fixture.Namespace.Type", "İ.Σ.中文", "Straße"] {
        for metadata in [registered.to_string(), registered.to_lowercase(), registered.to_uppercase(), "unknown".to_string()] {
            let mut vm = VM::new();
            vm.register_host_fn("fixture", "action", Box::new(|_, _| Value::I32(7)));
            let host_index = vm.host_registry[&("fixture".to_string(), "action".to_string())];
            let mut class = vm.type_registry.get(0).unwrap().clone();
            class.name = registered.into();
            let class_id = vm.type_registry.register(class);
            vm.type_registry.add_host_method(class_id, "fixtureAction", host_index);
            for actual_type in [0, class_id] {
            let mut object = Object::new_typed(actual_type);
            object.properties.insert("__type".into(), Value::String(metadata.clone().into()));
            let object = vybe_runtime::heap::alloc(object);
            vm.register_host_fn("fixture", "object", Box::new(move |_, _| Value::Object(object.clone())));
            let mut c = Chunk::new("member-type-metadata");
            let key = c.add_constant(Value::String("fixtureAction".into()));
            let make_object = c.add_import("fixture", "object");
            c.emit_call(make_object, 0, 1);
            c.emit_struct_field_op(Op::STRUCT_GET, 0, key, 1);
            c.emit_op(Op::RETURN, 1);
            let method = vm.run(vec![c]).unwrap();
            if actual_type == class_id {
                assert!(matches!(vm.invoke(&method, &[]).unwrap(), Value::I32(7)), "{registered:?} / {metadata:?}");
            } else {
                assert!(matches!(method, Value::Undefined), "{registered:?} / {metadata:?}");
            }
            }
        }
    }
}
