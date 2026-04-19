/// Tests for shared-everything threads: shared GC object access.

use vybe_bytecode::{VM, Value, Chunk, Op};
use std::sync::Arc;

#[test]
fn shared_new_creates_typed_object() {
    let mut vm = VM::new();
    // Register a type
    let mut td = vybe_bytecode::TypeDef::new("SharedPoint");
    td.add_field("x");
    td.add_field("y");
    let tid = vm.type_registry.register(td);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    let type_id = chunk.add_constant(Value::I32(tid as i32));
    chunk.emit_op_u16(Op::CONST, type_id, 0);
    chunk.emit_op(Op::SHARED_NEW, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    match &result {
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            assert_eq!(o.type_id, tid);
            assert_eq!(o.fields.len(), 2); // x, y
        }
        other => panic!("expected Object, got {:?}", other),
    }
}

#[test]
fn shared_struct_get_set() {
    let mut vm = VM::new();
    let mut td = vybe_bytecode::TypeDef::new("Counter");
    td.add_field("count");
    let tid = vm.type_registry.register(td);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    // Create shared object
    let type_id = chunk.add_constant(Value::I32(tid as i32));
    chunk.emit_op_u16(Op::CONST, type_id, 0);
    chunk.emit_op(Op::SHARED_NEW, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);

    // Set field 0 to 42
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let val = chunk.add_constant(Value::I32(42));
    chunk.emit_op_u16(Op::CONST, val, 0);
    chunk.emit_op_u16(Op::SHARED_STRUCT_SET, 0, 0); // field 0

    // Read field 0
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    chunk.emit_op_u16(Op::SHARED_STRUCT_GET, 0, 0); // field 0

    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 42);
}

#[test]
fn shared_array_get_set() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    // Create array [10, 20, 30]
    let v10 = chunk.add_constant(Value::I32(10));
    let v20 = chunk.add_constant(Value::I32(20));
    let v30 = chunk.add_constant(Value::I32(30));
    chunk.emit_op_u16(Op::CONST, v10, 0);
    chunk.emit_op_u16(Op::CONST, v20, 0);
    chunk.emit_op_u16(Op::CONST, v30, 0);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);

    // shared_array_set: arr[1] = 99
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let idx = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, idx, 0);
    let v99 = chunk.add_constant(Value::I32(99));
    chunk.emit_op_u16(Op::CONST, v99, 0);
    chunk.emit_op(Op::SHARED_ARRAY_SET, 0);

    // shared_array_get: arr[1]
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let idx2 = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, idx2, 0);
    chunk.emit_op(Op::SHARED_ARRAY_GET, 0);

    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 99);
}

#[test]
fn shared_struct_cas_success() {
    let mut vm = VM::new();
    let mut td = vybe_bytecode::TypeDef::new("Atomic");
    td.add_field("value");
    let tid = vm.type_registry.register(td);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    // Create shared object, set field 0 to 10
    let type_id = chunk.add_constant(Value::I32(tid as i32));
    chunk.emit_op_u16(Op::CONST, type_id, 0);
    chunk.emit_op(Op::SHARED_NEW, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let v10 = chunk.add_constant(Value::I32(10));
    chunk.emit_op_u16(Op::CONST, v10, 0);
    chunk.emit_op_u16(Op::SHARED_STRUCT_SET, 0, 0);

    // CAS: expect 10, set 20 → should succeed, return old value 10
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let expected = chunk.add_constant(Value::I32(10));
    chunk.emit_op_u16(Op::CONST, expected, 0);
    let new_val = chunk.add_constant(Value::I32(20));
    chunk.emit_op_u16(Op::CONST, new_val, 0);
    chunk.emit_op_u16(Op::SHARED_STRUCT_CAS, 0, 0); // field 0

    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 10, "CAS should return old value");
}

#[test]
fn shared_struct_cas_failure() {
    let mut vm = VM::new();
    let mut td = vybe_bytecode::TypeDef::new("Atomic2");
    td.add_field("value");
    let tid = vm.type_registry.register(td);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    // Create shared object, set field 0 to 10
    let type_id = chunk.add_constant(Value::I32(tid as i32));
    chunk.emit_op_u16(Op::CONST, type_id, 0);
    chunk.emit_op(Op::SHARED_NEW, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let v10 = chunk.add_constant(Value::I32(10));
    chunk.emit_op_u16(Op::CONST, v10, 0);
    chunk.emit_op_u16(Op::SHARED_STRUCT_SET, 0, 0);

    // CAS: expect 99 (wrong), set 20 → should fail, return old value 10, field unchanged
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let expected = chunk.add_constant(Value::I32(99));
    chunk.emit_op_u16(Op::CONST, expected, 0);
    let new_val = chunk.add_constant(Value::I32(20));
    chunk.emit_op_u16(Op::CONST, new_val, 0);
    chunk.emit_op_u16(Op::SHARED_STRUCT_CAS, 0, 0);
    chunk.emit_op(Op::DROP, 0); // drop old value

    // Read field — should still be 10 (CAS failed)
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    chunk.emit_op_u16(Op::SHARED_STRUCT_GET, 0, 0);

    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 10, "field should be unchanged after failed CAS");
}
