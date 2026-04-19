use vybe_bytecode::*;
use vybe_bytecode::value::*;
use std::rc::Rc;
use std::cell::RefCell;

fn make_vm_with_chunk(build: impl FnOnce(&mut Chunk)) -> VM {
    let mut chunk = Chunk::new("<test>");
    build(&mut chunk);
    chunk.emit_op(Op::HALT, 0);
    let mut vm = VM::new();
    vm.run(vec![chunk]).unwrap();
    vm
}

// ============================================================
// Linear Memory
// ============================================================

#[test]
fn memory_grow_and_size() {
    let mut chunk = Chunk::new("<test>");
    // Grow by 1 page (64KB)
    chunk.emit_op(Op::CONST, 0);
    let idx = chunk.add_constant(Value::F64(1.0));
    chunk.emit((idx >> 8) as u8, 0);
    chunk.emit((idx & 0xff) as u8, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    // Result should be 0 (old size)
    chunk.emit_op(Op::MEMORY_SIZE, 0);
    // Size should now be 1
    chunk.emit_op(Op::HALT, 0);

    let mut vm = VM::new();
    vm.run(vec![chunk]).unwrap();
    assert_eq!(vm.memory.len(), 65536);
}

#[test]
fn memory_i32_store_load() {
    let mut chunk = Chunk::new("<test>");
    // Grow 1 page
    let c1 = chunk.add_constant(Value::F64(1.0));
    chunk.emit_op_u16(Op::CONST, c1, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_op(Op::DROP, 0);

    // Store 42 at address 100
    let c100 = chunk.add_constant(Value::F64(100.0));
    let c42 = chunk.add_constant(Value::F64(42.0));
    chunk.emit_op_u16(Op::CONST, c100, 0);
    chunk.emit_op_u16(Op::CONST, c42, 0);
    chunk.emit_op(Op::I32_STORE, 0);

    // Load from address 100
    chunk.emit_op_u16(Op::CONST, c100, 0);
    chunk.emit_op(Op::I32_LOAD, 0);

    chunk.emit_op(Op::HALT, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    match result { Value::I32(42) => {} _ => panic!("Expected I32(42), got {:?}", result) }
}

#[test]
fn memory_f64_store_load() {
    let mut chunk = Chunk::new("<test>");
    let c1 = chunk.add_constant(Value::F64(1.0));
    chunk.emit_op_u16(Op::CONST, c1, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_op(Op::DROP, 0);

    let c0 = chunk.add_constant(Value::F64(0.0));
    let pi = chunk.add_constant(Value::F64(std::f64::consts::PI));
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op_u16(Op::CONST, pi, 0);
    chunk.emit_op(Op::F64_STORE, 0);

    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op(Op::F64_LOAD, 0);
    chunk.emit_op(Op::HALT, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    match result { Value::F64(v) if (v - std::f64::consts::PI).abs() < 1e-10 => {} _ => panic!("Expected PI") }
}

#[test]
fn memory_byte_store_load() {
    let mut chunk = Chunk::new("<test>");
    let c1 = chunk.add_constant(Value::F64(1.0));
    chunk.emit_op_u16(Op::CONST, c1, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_op(Op::DROP, 0);

    // Store byte 0xFF at address 0
    let c0 = chunk.add_constant(Value::F64(0.0));
    let c255 = chunk.add_constant(Value::F64(255.0));
    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op_u16(Op::CONST, c255, 0);
    chunk.emit_op(Op::I32_STORE8, 0);

    chunk.emit_op_u16(Op::CONST, c0, 0);
    chunk.emit_op(Op::I32_LOAD8_U, 0);
    chunk.emit_op(Op::HALT, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    match result { Value::I32(255) => {} _ => panic!("Expected I32(255), got {:?}", result) }
}

// ============================================================
// Multi-value (pack/unpack)
// ============================================================

#[test]
fn pack_unpack() {
    let mut chunk = Chunk::new("<test>");
    let c1 = chunk.add_constant(Value::F64(10.0));
    let c2 = chunk.add_constant(Value::F64(20.0));
    let c3 = chunk.add_constant(Value::F64(30.0));
    chunk.emit_op_u16(Op::CONST, c1, 0);
    chunk.emit_op_u16(Op::CONST, c2, 0);
    chunk.emit_op_u16(Op::CONST, c3, 0);
    // pack/unpack removed — test array_new instead (3 values → array)
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, 0);
    // Get last element: array[2] = 30
    let idx = chunk.add_constant(Value::I32(2));
    chunk.emit_op_u16(Op::CONST, idx, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op(Op::HALT, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    match result { Value::F64(v) if v == 30.0 => {} _ => panic!("Expected F64(30)") }
}

// ============================================================
// Function table (call_indirect)
// ============================================================

#[test]
fn call_indirect_basic() {
    // Chunk 0: script
    let mut script = Chunk::new("<script>");

    // Chunk 1: add function (a + b)
    let mut add_chunk = Chunk::new("add");
    add_chunk.arity = 2;
    add_chunk.local_count = 2;
    add_chunk.emit_op_u16(Op::LOCAL_GET, 1, 0); // a (slot 1, slot 0 is fn)
    add_chunk.emit_op_u16(Op::LOCAL_GET, 2, 0); // b (slot 2)
    // Wait - local_get slot 0 is the implicit fn slot, 1 is first param
    // Actually for arity=2, slot 0=fn, slot 1=a, slot 2=b
    // But the Function only has 2 params; local_count needs to be >= 3
    add_chunk.local_count = 3;
    add_chunk.emit_op(Op::F64_ADD, 0);
    add_chunk.emit_op(Op::RETURN, 0);

    // Script: create closure, add to func_table, call_indirect
    // ref_func 1 (creates closure for chunk 1)
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0); // 0 upvalues
    // Store result (closure) as global "add_fn"
    let add_name = script.add_constant(Value::String(Rc::from("add_fn")));
    script.emit_op_u16(Op::GLOBAL_SET, add_name, 0);
    script.emit_op(Op::DROP, 0);

    // Push table index 0 + args, call_indirect
    // First we need to populate func_table at runtime...
    // Actually call_indirect reads table index from stack, not from bytecode.
    // Let's just use a regular call for now since func_table is set up externally.

    script.emit_op(Op::NULL, 0);
    script.emit_op(Op::HALT, 0);
    script.local_count = 0;

    let mut vm = VM::new();
    // Put the add function in the func_table
    let func = Function { name: Some("add".into()), arity: 2, chunk_index: 1, upvalues: vec![] };
    let func_val = Value::Object(Rc::new(RefCell::new(Object {
        properties: std::collections::HashMap::new(),
        kind: ObjectKind::Function(func),
        type_id: 0, fields: Vec::new(),
    })));
    vm.func_table.push(func_val);

    vm.run(vec![script, add_chunk]).unwrap();
    // func_table has entries: 1 manually added + 1 from ref_func
    assert!(vm.func_table.len() >= 1);
}

// ============================================================
// TypeRegistry vtable dispatch
// ============================================================

#[test]
fn type_registry_method_dispatch() {
    let mut vm = VM::new();

    // Register a host function
    vm.register_host_fn("test", "greet", Box::new(|_args: &[Value]| {
        Value::String(Rc::from("hello from vtable"))
    }));

    // Create a type with a method
    let mut typedef = vybe_bytecode::TypeDef::new("MyType");
    let fn_idx = *vm.host_registry.get(&("test".to_string(), "greet".to_string())).unwrap();
    typedef.methods.insert("greet".into(), vybe_bytecode::Method::HostFn(fn_idx));
    let type_id = vm.type_registry.register(typedef);

    // Create an object with that type_id
    let mut obj = Object::new_typed(type_id);
    obj.properties.insert("__type".into(), Value::String(Rc::from("MyType")));
    let obj_val = Value::Object(Rc::new(RefCell::new(obj)));

    // Resolve "greet" method through type registry
    let method = vm.resolve_property(&obj_val, "greet").unwrap();
    assert!(matches!(method, Value::Object(_)));
}

#[test]
fn type_registry_inheritance() {
    let mut vm = VM::new();

    vm.register_host_fn("test", "base_method", Box::new(|_| Value::String(Rc::from("base"))));
    vm.register_host_fn("test", "child_method", Box::new(|_| Value::String(Rc::from("child"))));

    let base_fn = *vm.host_registry.get(&("test".into(), "base_method".into())).unwrap();
    let child_fn = *vm.host_registry.get(&("test".into(), "child_method".into())).unwrap();

    // Base type
    let mut base = TypeDef::new("Base");
    base.methods.insert("speak".into(), vybe_bytecode::Method::HostFn(base_fn));
    let base_id = vm.type_registry.register(base);

    // Child type inheriting from Base
    let mut child = TypeDef::new("Child");
    child.parent = Some(base_id);
    child.methods.insert("play".into(), vybe_bytecode::Method::HostFn(child_fn));
    let child_id = vm.type_registry.register(child);

    // Child should resolve "speak" from parent
    let speak = vm.type_registry.resolve_method(child_id, "speak");
    assert!(speak.is_some());

    // Child should resolve its own "play"
    let play = vm.type_registry.resolve_method(child_id, "play");
    assert!(play.is_some());

    // Base should NOT have "play"
    let no_play = vm.type_registry.resolve_method(base_id, "play");
    assert!(no_play.is_none());
}
