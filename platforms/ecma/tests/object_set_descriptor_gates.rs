use std::sync::{Arc, atomic::{AtomicUsize, Ordering}};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, Value, VM};

fn probe(accessor: Option<bool>, freeze: bool, strict: bool) -> (bool, usize, bool) {
    let mut vm = VM::new();
    vybe_platform_ecma::object::register(&mut vm);
    let calls = Arc::new(AtomicUsize::new(0));
    let recorded = calls.clone();
    vm.register_host_fn("fixture", "accessor", Box::new(move |_, _| {
        recorded.fetch_add(1, Ordering::Relaxed);
        Value::I32(7)
    }));
    let mut object = Object::new();
    if let Some(setter) = accessor {
        let index = vm.host_registry[&("fixture".to_string(), "accessor".to_string())];
        let mut function = Object::new();
        function.kind = ObjectKind::HostFunction(index);
        object.properties.insert(if setter { "__set_value" } else { "__get_value" }.into(), Value::Object(vybe_runtime::heap::alloc(function)));
    } else {
        object.properties.insert("value".into(), Value::I32(7));
    }
    let object = vybe_runtime::heap::alloc(object);
    let shared = object.clone();
    vm.register_host_fn("fixture", "object", Box::new(move |_, _| Value::Object(shared.clone())));
    vm.register_host_fn("fixture", "key", Box::new(|_, _| Value::String("value".into())));
    let mut c = Chunk::new("object-set-descriptor-gates");
    c.local_count = 1;
    let make_object = c.add_import("fixture", "object");
    let make_key = c.add_import("fixture", "key");
    let set = c.add_import("ecma:object", "set");
    c.emit_call(make_object, 0, 1);
    c.emit_op_u16(Op::LOCAL_SET, 0, 1);
    if freeze {
        let freeze_import = c.add_import("ecma:object", "freeze");
        c.emit_op_u16(Op::LOCAL_GET, 0, 1);
        c.emit_call(freeze_import, 1, 1);
        c.emit_op(Op::DROP, 1);
    }
    c.emit_op_u16(Op::LOCAL_GET, 0, 1);
    c.emit_call(make_key, 0, 1);
    c.emit_i32_const(99, 1);
    c.emit_i32_const(i32::from(strict), 1);
    c.emit_call(set, 4, 1);
    c.emit_op(Op::RETURN, 1);
    let failed = vm.run(vec![c]).is_err();
    let object = object.lock().unwrap();
    let overwritten = matches!(object.properties.get("value"), Some(Value::I32(99)));
    (failed, calls.load(Ordering::Relaxed), overwritten)
}

#[test]
fn frozen_data_property_is_not_overwritten() {
    assert_eq!(probe(None, true, false), (false, 0, false));
    assert_eq!(probe(None, true, true), (true, 0, false));
}

#[test]
fn own_setter_is_called_instead_of_creating_data_storage() {
    assert_eq!(probe(Some(true), false, false), (false, 1, false));
    assert_eq!(probe(Some(true), false, true), (false, 1, false));
}

#[test]
fn getter_only_property_remains_read_only() {
    assert_eq!(probe(Some(false), false, false), (false, 0, false));
    assert_eq!(probe(Some(false), false, true), (true, 0, false));
}
