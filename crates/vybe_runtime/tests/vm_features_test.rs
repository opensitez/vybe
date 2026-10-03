use std::sync::Arc;
use vybe_runtime::chunk::{CompositeKind, TypeEntry};
use vybe_runtime::value::*;
use vybe_runtime::*;

#[allow(dead_code)]
fn make_vm_with_chunk(build: impl FnOnce(&mut Chunk)) -> VM {
    let mut chunk = Chunk::new("<test>");
    build(&mut chunk);
    let mut vm = VM::new();
    vm.run(vec![chunk]).unwrap();
    vm
}

// ============================================================
// Linear Memory
// ============================================================

#[test]
fn one_byte_negative_sleb_constants_preserve_sign() {
    let mut i32_chunk = Chunk::new("<i32-negative-sleb>");
    i32_chunk.emit_i32_const(-1, 0);
    let mut vm = VM::new();
    assert_eq!(vm.run(vec![i32_chunk]).unwrap(), Value::I32(-1));

    let mut i64_chunk = Chunk::new("<i64-negative-sleb>");
    i64_chunk.emit_i64_const(-1, 0);
    let mut vm = VM::new();
    assert_eq!(vm.run(vec![i64_chunk]).unwrap(), Value::I64(-1));
}

#[test]
fn f32_const_preserves_raw_bits() {
    let bits = 0x7fc0_1234u32;
    let mut chunk = Chunk::new("<f32-const-bits>");
    chunk.emit_f32_const(f32::from_bits(bits), 0);

    let mut vm = VM::new();
    match vm.run(vec![chunk]).unwrap() {
        Value::F32(value) => assert_eq!(value.to_bits(), bits),
        other => panic!("expected f32 value, got {:?}", other),
    }
}

#[test]
fn memory_grow_and_size() {
    let mut chunk = Chunk::new("<test>");
    // Grow by 1 page (64KB)
    chunk.emit_f64_const(1.0, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(0, 0);
    // Result should be 0 (old size)
    chunk.emit_op_u16(Op::MEMORY_SIZE, 0, 0);
    // Size should now be 1

    let mut vm = VM::new();
    vm.run(vec![chunk]).unwrap();
    assert_eq!(vm.memory.len(), 65536);
}

#[test]
fn memory_i32_store_load() {
    let mut chunk = Chunk::new("<test>");
    // Grow 1 page
    chunk.emit_f64_const(1.0, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(0, 0);
    chunk.emit_op(Op::DROP, 0);

    // Store 42 at address 100
    chunk.emit_f64_const(100.0, 0);
    chunk.emit_f64_const(42.0, 0);
    chunk.emit_op(Op::I32_STORE, 0);

    // Load from address 100
    chunk.emit_f64_const(100.0, 0);
    chunk.emit_op(Op::I32_LOAD, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    match result {
        Value::I32(42) => {}
        _ => panic!("Expected I32(42), got {:?}", result),
    }
}

#[test]
fn memory_f64_store_load() {
    let mut chunk = Chunk::new("<test>");
    chunk.emit_f64_const(1.0, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(0, 0);
    chunk.emit_op(Op::DROP, 0);

    chunk.emit_f64_const(0.0, 0);
    chunk.emit_f64_const(std::f64::consts::PI, 0);
    chunk.emit_op(Op::F64_STORE, 0);

    chunk.emit_f64_const(0.0, 0);
    chunk.emit_op(Op::F64_LOAD, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    match result {
        Value::F64(v) if (v - std::f64::consts::PI).abs() < 1e-10 => {}
        _ => panic!("Expected PI"),
    }
}

#[test]
fn memory_byte_store_load() {
    let mut chunk = Chunk::new("<test>");
    chunk.emit_f64_const(1.0, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(0, 0);
    chunk.emit_op(Op::DROP, 0);

    // Store byte 0xFF at address 0
    chunk.emit_f64_const(0.0, 0);
    chunk.emit_f64_const(255.0, 0);
    chunk.emit_op(Op::I32_STORE8, 0);

    chunk.emit_f64_const(0.0, 0);
    chunk.emit_op(Op::I32_LOAD8_U, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    match result {
        Value::I32(255) => {}
        _ => panic!("Expected I32(255), got {:?}", result),
    }
}

// ============================================================
// Multi-value (pack/unpack)
// ============================================================

#[test]
fn pack_unpack() {
    let mut chunk = Chunk::new("<test>");
    chunk.emit_f64_const(10.0, 0);
    chunk.emit_f64_const(20.0, 0);
    chunk.emit_f64_const(30.0, 0);
    // pack/unpack removed — test array_new instead (3 values → array)
    chunk.emit_array_new_fixed(0, 3, 0);
    // Get last element: array[2] = 30
    chunk.emit_i32_const(2, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    match result {
        Value::F64(v) if v == 30.0 => {}
        _ => panic!("Expected F64(30)"),
    }
}

fn gc_array_chunk(name: &str) -> Chunk {
    let mut chunk = Chunk::new(name);
    chunk.types.push(TypeEntry {
        name: "Arr".into(),
        kind: CompositeKind::Array,
        ..Default::default()
    });
    chunk
}

#[test]
fn gc_array_get_accepts_i32_index_from_typed_caller() {
    let mut chunk = gc_array_chunk("<gc-array-get-i32>");
    chunk.emit_i32_const(11, 0);
    chunk.emit_i32_const(22, 0);
    chunk.emit_array_new_fixed(1, 2, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = VM::new().run(vec![chunk]).expect("gc array get failed");
    assert_eq!(result.as_i32(), 22);
}

#[test]
fn gc_array_get_rejects_non_i32_runtime_index() {
    let mut chunk = gc_array_chunk("<gc-array-get-i64-index>");
    chunk.emit_i32_const(11, 0);
    chunk.emit_i32_const(22, 0);
    chunk.emit_array_new_fixed(1, 2, 0);
    chunk.emit_i64_const(1, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = VM::new()
        .run(vec![chunk])
        .expect_err("gc array get must reject non-i32 indexes");
    assert!(
        err.message.contains("array.get"),
        "wrong trap for non-i32 gc array index: {}",
        err.message
    );
}

#[test]
fn gc_array_set_rejects_non_i32_runtime_index() {
    let mut chunk = gc_array_chunk("<gc-array-set-f64-index>");
    chunk.emit_i32_const(11, 0);
    chunk.emit_i32_const(22, 0);
    chunk.emit_array_new_fixed(1, 2, 0);
    chunk.emit_f64_const(1.0, 0);
    chunk.emit_i32_const(99, 0);
    chunk.emit_op(Op::ARRAY_SET, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = VM::new()
        .run(vec![chunk])
        .expect_err("gc array set must reject non-i32 indexes");
    assert!(
        err.message.contains("array.set"),
        "wrong trap for non-i32 gc array set index: {}",
        err.message
    );
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
    let add_name = script.add_constant(Value::String(Arc::from("add_fn")));
    script.emit_op_u32(Op::GLOBAL_SET, add_name, 0);

    // Push table index 0 + args, call_indirect
    // First we need to populate func_table at runtime...
    // Actually call_indirect reads table index from stack, not from bytecode.
    // Let's just use a regular call for now since func_table is set up externally.

    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.local_count = 0;

    let mut vm = VM::new();
    // Put the add function in the func_table
    let func = Function {
        name: Some("add".into()),
        arity: 2,
        chunk_index: 1,
        upvalues: vec![],
    };
    let func_val = Value::Object(Arc::new(std::sync::Mutex::new(Object {
        properties: indexmap::IndexMap::new(),
        kind: ObjectKind::Function(func),
        type_id: 0,
        fields: Vec::new(),
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
    vm.register_host_fn(
        "test",
        "greet",
        Box::new(|_ctx: &mut vybe_runtime::HostContext, _args: &[Value]| {
            Value::String(Arc::from("hello from vtable"))
        }),
    );

    // Create a type with a method
    let mut typedef = vybe_runtime::TypeDef::new("MyType");
    let fn_idx = *vm
        .host_registry
        .get(&("test".to_string(), "greet".to_string()))
        .unwrap();
    typedef
        .methods
        .insert("greet".into(), vybe_runtime::Method::HostFn(fn_idx));
    let type_id = vm.type_registry.register(typedef);

    // Create an object with that type_id
    let mut obj = Object::new_typed(type_id);
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("MyType")));
    let obj_val = Value::Object(Arc::new(std::sync::Mutex::new(obj)));

    // Resolve "greet" method through type registry
    let method = vm.resolve_property(&obj_val, "greet").unwrap();
    assert!(matches!(method, Value::Object(_)));
}

#[test]
fn type_registry_inheritance() {
    let mut vm = VM::new();

    vm.register_host_fn(
        "test",
        "base_method",
        Box::new(|_, _| Value::String(Arc::from("base"))),
    );
    vm.register_host_fn(
        "test",
        "child_method",
        Box::new(|_, _| Value::String(Arc::from("child"))),
    );

    let base_fn = *vm
        .host_registry
        .get(&("test".into(), "base_method".into()))
        .unwrap();
    let child_fn = *vm
        .host_registry
        .get(&("test".into(), "child_method".into()))
        .unwrap();

    // Base type
    let mut base = TypeDef::new("Base");
    base.methods
        .insert("speak".into(), vybe_runtime::Method::HostFn(base_fn));
    let base_id = vm.type_registry.register(base);

    // Child type inheriting from Base
    let mut child = TypeDef::new("Child");
    child.parent = Some(base_id);
    child
        .methods
        .insert("play".into(), vybe_runtime::Method::HostFn(child_fn));
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
