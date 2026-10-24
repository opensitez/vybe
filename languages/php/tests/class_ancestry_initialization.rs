use vybe_compiler::primitives::{bundle, globals, platforms, reflection};
use vybe_runtime::{value::{Object, ObjectKind}, Chunk, Op, Value, VM};

#[test]
fn ancestry_stamps_initialize_absent_null_and_undefined_metadata() {
    let mut root = Chunk::new("ancestry-entry");
    root.emit_op_u16(Op::REF_FUNC, 1, 1);
    root.emit(0, 1);
    root.emit_op(Op::RETURN, 1);
    let mut body = Chunk::new("append-ancestry");
    body.arity = 1;
    body.local_count = 1;
    let mut chunks = vec![root, body];
    reflection::emit_instanceof_chain(&mut chunks, 1, 0, "Base", 1);
    reflection::emit_instanceof_chain(&mut chunks, 1, 0, "Marker", 1);
    chunks[1].emit_op_u16(Op::LOCAL_GET, 0, 1);
    chunks[1].emit_op(Op::RETURN, 1);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    let append = vm.run(chunks).unwrap();
    for initial in [None, Some(Value::Null), Some(Value::Undefined),
        Some(Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![Value::String("Ancestor".into())]))))] {
        let retained = matches!(&initial, Some(Value::Object(_)));
        let mut object = Object::new();
        if let Some(value) = initial { object.properties.insert("__types".into(), value); }
        let object = vybe_runtime::heap::alloc(object);
        let result = vm.invoke(&append, &[Value::Object(object.clone())]).unwrap();
        let Value::Object(result) = result else { panic!("receiver lost") };
        assert!(std::sync::Arc::ptr_eq(&result, &object));
        let types = object.lock().unwrap().properties.get("__types").cloned().unwrap();
        let Value::Object(types) = types else { panic!("ancestry array missing: {types:?}") };
        let types = types.lock().unwrap();
        let ObjectKind::Array(items) = &types.kind else { panic!("ancestry is not an array") };
        let mut expected = Vec::new();
        if retained { expected.push(Value::String("Ancestor".into())); }
        expected.extend([Value::String("Base".into()), Value::String("Marker".into())]);
        assert_eq!(items, &expected);
    }
}
