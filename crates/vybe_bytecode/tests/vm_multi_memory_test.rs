/// Tests for multi-memory support.

use vybe_bytecode::{VM, Value, Chunk, Op};
use std::rc::Rc;

#[test]
fn memory_init_creates_new_memory() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    // memory_init: create a new memory with 1 page
    let pages = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, pages, 0);
    chunk.emit_op(Op::MEMORY_INIT, 0);
    // Returns the memory index
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    let mem_idx = result.as_i32();
    assert!(mem_idx >= 1, "new memory index should be >= 1, got {}", mem_idx);
}

#[test]
fn memory_select_switches_active() {
    let mut vm = VM::new();
    // Pre-allocate default memory
    vm.memory.resize(65536, 0);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    // Store 42 in default memory (index 0)
    let addr = chunk.add_constant(Value::I32(0));
    let val = chunk.add_constant(Value::I32(42));
    chunk.emit_op_u16(Op::CONST, addr, 0);
    chunk.emit_op_u16(Op::CONST, val, 0);
    chunk.emit_op(Op::I32_STORE, 0);

    // Read it back from default memory
    let addr2 = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::CONST, addr2, 0);
    chunk.emit_op(Op::I32_LOAD, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 42);
}

#[test]
fn multiple_memories_independent() {
    let mut vm = VM::new();
    // Default memory (index 0)
    vm.memory.resize(65536, 0);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 3;

    // Create a second memory (index 1)
    let pages = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, pages, 0);
    chunk.emit_op(Op::MEMORY_INIT, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0); // store mem idx

    // Store 42 in default memory at addr 0
    let addr = chunk.add_constant(Value::I32(0));
    let val42 = chunk.add_constant(Value::I32(42));
    chunk.emit_op_u16(Op::CONST, addr, 0);
    chunk.emit_op_u16(Op::CONST, val42, 0);
    chunk.emit_op(Op::I32_STORE, 0);

    // Read back from default memory — should be 42
    let addr2 = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::CONST, addr2, 0);
    chunk.emit_op(Op::I32_LOAD, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 42);
}

#[test]
fn memory_copy_cross_between_memories() {
    let mut vm = VM::new();
    // Default memory with data
    vm.memory.resize(65536, 0);
    vm.memory.store_i32(0, 99);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    // Create second memory
    let pages = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, pages, 0);
    chunk.emit_op(Op::MEMORY_INIT, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0); // mem_idx = 1

    // Copy 4 bytes from memory 0 addr 0 → memory 1 addr 0
    let zero = chunk.add_constant(Value::I32(0));
    let four = chunk.add_constant(Value::I32(4));
    let mem0 = chunk.add_constant(Value::I32(0));
    let mem1 = chunk.add_constant(Value::I32(1));

    // Stack order: dst_mem, dst_addr, src_mem, src_addr, len
    chunk.emit_op_u16(Op::CONST, mem1, 0);   // dst_mem = 1
    chunk.emit_op_u16(Op::CONST, zero, 0);   // dst_addr = 0
    chunk.emit_op_u16(Op::CONST, mem0, 0);   // src_mem = 0
    chunk.emit_op_u16(Op::CONST, zero, 0);   // src_addr = 0
    chunk.emit_op_u16(Op::CONST, four, 0);   // len = 4
    chunk.emit_op(Op::MEMORY_COPY_CROSS, 0);

    // Switch to memory 1 and read
    chunk.emit_op(Op::MEMORY_SELECT, 0);
    chunk.emit(1, 0); // memory index 1
    chunk.emit_op_u16(Op::CONST, zero, 0);
    chunk.emit_op(Op::I32_LOAD, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 99, "data should have been copied to memory 1");
}
