/// Tests for weak references, finalizers, and GC post-MVP features.

use vybe_bytecode::{VM, Value, Chunk, Op};
use std::rc::Rc;
use std::cell::RefCell;
use vybe_bytecode::value::{Object, ObjectKind};

#[test]
fn make_weak_ref() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    // Create an object, make a weak ref
    let name = chunk.add_constant(Value::String(Rc::from("hello")));
    chunk.emit_op_u16(Op::r#const, name, 0);   // push string
    chunk.emit_op(Op::ref_make_weak, 0);        // make weak ref (strings → null)
    chunk.emit_op(Op::halt, 0);

    let result = vm.run(vec![chunk]).unwrap();
    // Strings aren't objects, so ref_make_weak returns null
    assert!(matches!(result, Value::Null));
}

#[test]
fn weak_ref_from_object() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 3;

    // Create an array object [1, 2, 3]
    let one = chunk.add_constant(Value::I32(1));
    let two = chunk.add_constant(Value::I32(2));
    let three = chunk.add_constant(Value::I32(3));
    chunk.emit_op_u16(Op::r#const, one, 0);
    chunk.emit_op_u16(Op::r#const, two, 0);
    chunk.emit_op_u16(Op::r#const, three, 0);
    chunk.emit_op_u16(Op::array_new, 3, 0);

    // Store in local 1 (keep strong ref)
    chunk.emit_op_u16(Op::local_set, 1, 0);

    // Get it back, make weak ref, store in local 2
    chunk.emit_op_u16(Op::local_get, 1, 0);
    chunk.emit_op(Op::ref_make_weak, 0);
    chunk.emit_op_u16(Op::local_set, 2, 0);

    // Deref the weak ref — should succeed (strong ref still in local 1)
    chunk.emit_op_u16(Op::local_get, 2, 0);
    chunk.emit_op(Op::ref_deref_weak, 0);
    chunk.emit_op(Op::halt, 0);

    let result = vm.run(vec![chunk]).unwrap();
    // Should return the array object
    assert!(matches!(result, Value::Object(_)));
}

#[test]
fn weak_ref_is_alive() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 3;

    // Create object, store in local 1
    let val = chunk.add_constant(Value::I32(42));
    chunk.emit_op_u16(Op::r#const, val, 0);
    chunk.emit_op_u16(Op::array_new, 1, 0);
    chunk.emit_op_u16(Op::local_set, 1, 0);

    // Make weak ref, store in local 2
    chunk.emit_op_u16(Op::local_get, 1, 0);
    chunk.emit_op(Op::ref_make_weak, 0);
    chunk.emit_op_u16(Op::local_set, 2, 0);

    // Check if alive — should be true
    chunk.emit_op_u16(Op::local_get, 2, 0);
    chunk.emit_op(Op::ref_is_alive, 0);
    chunk.emit_op(Op::halt, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert!(matches!(result, Value::Bool(true)));
}

#[test]
fn deref_strong_ref_passthrough() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;

    // Push a regular value, deref_weak should just pass through
    let val = chunk.add_constant(Value::I32(99));
    chunk.emit_op_u16(Op::r#const, val, 0);
    chunk.emit_op(Op::ref_deref_weak, 0);
    chunk.emit_op(Op::halt, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 99);
}

#[test]
fn register_finalizer() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    // Create an object
    let val = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::r#const, val, 0);
    chunk.emit_op_u16(Op::array_new, 1, 0);
    chunk.emit_op(Op::dup, 0);
    chunk.emit_op_u16(Op::local_set, 1, 0);

    // Register a finalizer (null callback for now — just test it doesn't crash)
    chunk.emit_op(Op::null, 0);
    chunk.emit_op(Op::ref_register_finalizer, 0);

    // Return something
    chunk.emit_op_u16(Op::local_get, 1, 0);
    chunk.emit_op(Op::halt, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert!(matches!(result, Value::Object(_)));

    // After run, collect_dead_finalizers should work
    let dead = vm.collect_dead_finalizers();
    // Object is still alive (held by result), so no dead finalizers
    assert!(dead.is_empty());
}

#[test]
fn is_alive_null_returns_false() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;

    chunk.emit_op(Op::null, 0);
    chunk.emit_op(Op::ref_is_alive, 0);
    chunk.emit_op(Op::halt, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert!(matches!(result, Value::Bool(false)));
}
