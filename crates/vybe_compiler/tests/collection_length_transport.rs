use vybe_compiler::primitives::collections;
use vybe_runtime::{Chunk, Value, VM, heap};
use vybe_runtime::opcode::Op;
use vybe_runtime::value::{Object, ObjectKind};

fn host(c: &mut Chunk, module: &str, name: &str, argc: u8) {
    let index = c.add_import(module, name);
    c.emit_call(index, argc, 0);
}
fn local(c: &mut Chunk, op: Op, slot: u16) { c.emit_op_u16(op, slot, 0); }
fn previous_len(c: &mut Chunk) {
    let value = c.alloc_scratch(1);
    local(c, Op::LOCAL_SET, value);
    local(c, Op::LOCAL_GET, value);
    host(c, "wasm:js-string", "test", 1);
    c.emit_if_value(0);
    local(c, Op::LOCAL_GET, value);
    host(c, "wasm:js-string", "length", 1);
    c.emit_else(0);
    local(c, Op::LOCAL_GET, value);
    host(c, "ecma:array", "length", 1);
    let array_length = c.alloc_scratch(1);
    local(c, Op::LOCAL_SET, array_length);
    local(c, Op::LOCAL_GET, array_length);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_if_value(0);
    local(c, Op::LOCAL_GET, value);
    host(c, "ecma:map", "size", 1);
    let map_length = c.alloc_scratch(1);
    local(c, Op::LOCAL_SET, map_length);
    local(c, Op::LOCAL_GET, map_length);
    c.emit_i32_const(0, 0);
    c.emit_op(Op::I32_NE, 0);
    c.emit_if_value(0);
    local(c, Op::LOCAL_GET, map_length);
    c.emit_else(0);
    local(c, Op::LOCAL_GET, value);
    host(c, "ecma:object", "keys", 1);
    host(c, "ecma:array", "length", 1);
    c.emit_end(0);
    c.emit_else(0);
    local(c, Op::LOCAL_GET, array_length);
    c.emit_end(0);
    c.emit_end(0);
}
fn execute(value: Value, candidate: bool) -> (Value, u64) {
    let mut vm = VM::new();
    vm.register_host_fn("test:collection", "input", Box::new(move |_, _| value.clone()));
    vm.register_host_fn("wasm:js-string", "test", Box::new(|_, args| {
        Value::I32(i32::from(matches!(args.first(), Some(Value::String(_)))))
    }));
    vm.register_host_fn("wasm:js-string", "length", Box::new(|_, args| {
        match args.first() {
            Some(Value::String(value)) => Value::I32(value.encode_utf16().count() as i32),
            _ => Value::Null,
        }
    }));
    vm.register_host_fn("ecma:array", "length", Box::new(|_, args| {
        if let Some(Value::Object(value)) = args.first() {
            if let ObjectKind::Array(items) = &value.lock().unwrap().kind {
                return Value::I32(items.len() as i32);
            }
        }
        Value::Null
    }));
    vm.register_host_fn("ecma:map", "size", Box::new(|_, args| {
        if let Some(Value::Object(value)) = args.first() {
            if let ObjectKind::Map(items) = &value.lock().unwrap().kind {
                return Value::I32(items.len() as i32);
            }
        }
        Value::I32(0)
    }));
    vm.register_host_fn("ecma:object", "keys", Box::new(|_, args| {
        let keys = if let Some(Value::Object(value)) = args.first() {
            value.lock().unwrap().properties.keys()
                .map(|key| Value::String(key.as_str().into())).collect()
        } else { Vec::new() };
        Value::Object(heap::alloc(Object::new_array(keys)))
    }));
    let mut chunks = vec![Chunk::new("length-transport")];
    host(&mut chunks[0], "test:collection", "input", 0);
    if candidate { collections::emit_len(&mut chunks, 0, 0); }
    else { previous_len(&mut chunks[0]); }
    chunks[0].emit_op(Op::RETURN, 0);
    vm.record_runtime_perf(true);
    let result = vm.run(chunks).unwrap();
    (result, vm.take_runtime_perf().unwrap().dispatches)
}
#[test]
fn length_transport_preserves_values_with_fewer_dispatches() {
    let mut map = Object::new();
    map.kind = ObjectKind::Map(Default::default());
    if let ObjectKind::Map(entries) = &mut map.kind {
        entries.insert(Value::I32(1), Value::Null);
    }
    let mut ordinary = Object::new();
    ordinary.properties.insert("first".into(), Value::Null);
    ordinary.properties.insert("second".into(), Value::Null);
    let fixtures = [
        (Value::String("abc".into()), 3, 1),
        (Value::Object(heap::alloc(Object::new_array(vec![]))), 0, 2),
        (Value::Object(heap::alloc(Object::new_array(vec![Value::Null; 2]))), 2, 2),
        (Value::Object(heap::alloc(map)), 1, 3),
        (Value::Object(heap::alloc(ordinary)), 2, 3),
        (Value::Null, 0, 3),
    ];
    for (value, expected, saved) in fixtures {
        let (before, old_dispatches) = execute(value.clone(), false);
        let (after, new_dispatches) = execute(value, true);
        assert!(matches!(before, Value::I32(value) if value == expected));
        assert!(matches!(after, Value::I32(value) if value == expected));
        assert_eq!(old_dispatches - new_dispatches, saved);
    }
}
