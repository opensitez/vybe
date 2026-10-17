use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, Value, VM};

fn write_property(non_extensible: bool, existing: bool) {
    let mut object = Object::new();
    if existing {
        object.properties.insert("value".into(), Value::I32(7));
    }
    let object = vybe_runtime::heap::alloc(object);
    let shared = object.clone();
    let mut vm = VM::new();
    vybe_platform_ecma::object::register(&mut vm);
    vm.register_host_fn("fixture", "object", Box::new(move |_, _| Value::Object(shared.clone())));
    vm.register_host_fn("fixture", "key", Box::new(|_, _| Value::String("value".into())));
    let mut c = Chunk::new("object-set-extensibility");
    c.local_count = 1;
    let make_object = c.add_import("fixture", "object");
    let make_key = c.add_import("fixture", "key");
    let set = c.add_import("ecma:object", "set");
    c.emit_call(make_object, 0, 1);
    c.emit_op_u16(Op::LOCAL_SET, 0, 1);
    if non_extensible {
        let prevent_extensions = c.add_import("ecma:object", "preventExtensions");
        c.emit_op_u16(Op::LOCAL_GET, 0, 1);
        c.emit_call(prevent_extensions, 1, 1);
        c.emit_op(Op::DROP, 1);
    }
    c.emit_op_u16(Op::LOCAL_GET, 0, 1);
    c.emit_call(make_key, 0, 1);
    c.emit_i32_const(99, 1);
    c.emit_i32_const(0, 1); // non-strict assignment
    c.emit_call(set, 4, 1);
    c.emit_op(Op::DROP, 1);
    c.emit_op_u16(Op::LOCAL_GET, 0, 1);
    c.emit_op(Op::RETURN, 1);
    vm.run(vec![c]).unwrap();
    let object = object.lock().unwrap();
    if non_extensible && !existing {
        assert!(!object.properties.contains_key("value"));
    } else {
        assert!(matches!(object.properties.get("value"), Some(Value::I32(99))));
    }
}

#[test]
fn extensible_writes_allow_new_and_existing_properties() {
    write_property(false, false);
    write_property(false, true);
}

#[test]
fn non_extensible_writes_allow_existing_but_reject_new_properties() {
    write_property(true, false);
    write_property(true, true);
}
