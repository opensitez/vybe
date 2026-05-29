//! Tests for the bulk-memory-operations WASM proposal.
//! Spec: `proposals/bulk-memory-operations/`, opcodes 0xFC 0x08–0x0E.
//!
//! memory.init (0x08) is skipped: the VM repurposes that opcode for
//! multi-memory allocation and does not implement the spec data-segment copy.

use vybe_bytecode::{Chunk, Op, VM, Value};

fn run_with_memory(mem_size: usize, emit: impl FnOnce(&mut Chunk)) -> VM {
    let mut vm = VM::new();
    vm.memory.resize(mem_size, 0);
    let mut chunk = Chunk::new("<script>");
    emit(&mut chunk);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("VM execution failed");
    vm
}

fn push_i32(c: &mut Chunk, v: i32) {
    let idx = c.add_constant(Value::I32(v));
    c.emit_op_u16(Op::CONST, idx, 0);
}

fn read_byte(vm: &VM, addr: usize) -> u8 {
    let mut buf = [0u8; 1];
    vm.memory.read_bytes(addr, &mut buf);
    buf[0]
}

// ── memory.fill (0xFC 0x0B) ──────────────────────────────────────────

#[test]
fn memory_fill_writes_byte_to_range() {
    let vm = run_with_memory(256, |c| {
        // memory.fill dst=10 val=0xAB count=5
        push_i32(c, 10);
        push_i32(c, 0xAB);
        push_i32(c, 5);
        c.emit_op(Op::MEMORY_FILL, 0);
    });
    assert_eq!(read_byte(&vm, 9), 0x00); // before range untouched
    for addr in 10..15 {
        assert_eq!(read_byte(&vm, addr), 0xAB);
    }
    assert_eq!(read_byte(&vm, 15), 0x00); // after range untouched
}

#[test]
fn memory_fill_zero_count_is_noop() {
    let vm = run_with_memory(64, |c| {
        push_i32(c, 0);
        push_i32(c, 0xFF);
        push_i32(c, 0); // count = 0
        c.emit_op(Op::MEMORY_FILL, 0);
    });
    assert_eq!(read_byte(&vm, 0), 0x00);
}

#[test]
fn memory_fill_zero_byte_clears_range() {
    // Pre-fill then clear with 0
    let vm = run_with_memory(64, |c| {
        push_i32(c, 4);
        push_i32(c, 0xFF);
        push_i32(c, 4);
        c.emit_op(Op::MEMORY_FILL, 0);
        push_i32(c, 5);
        push_i32(c, 0x00);
        push_i32(c, 2);
        c.emit_op(Op::MEMORY_FILL, 0);
    });
    assert_eq!(read_byte(&vm, 4), 0xFF); // before cleared range
    assert_eq!(read_byte(&vm, 5), 0x00);
    assert_eq!(read_byte(&vm, 6), 0x00);
    assert_eq!(read_byte(&vm, 7), 0xFF); // after cleared range
}

// ── memory.copy (0xFC 0x0A) ──────────────────────────────────────────

#[test]
fn memory_copy_non_overlapping() {
    let vm = run_with_memory(256, |c| {
        // Fill src region [20..24] with 0x55
        push_i32(c, 20);
        push_i32(c, 0x55);
        push_i32(c, 4);
        c.emit_op(Op::MEMORY_FILL, 0);
        // Copy to dst [0..4]
        push_i32(c, 0);  // dst
        push_i32(c, 20); // src
        push_i32(c, 4);  // count
        c.emit_op(Op::MEMORY_COPY, 0);
    });
    for addr in 0..4 {
        assert_eq!(read_byte(&vm, addr), 0x55);
    }
    assert_eq!(read_byte(&vm, 4), 0x00); // untouched
}

#[test]
fn memory_copy_zero_count_is_noop() {
    let vm = run_with_memory(64, |c| {
        push_i32(c, 10);
        push_i32(c, 0xFF);
        push_i32(c, 4);
        c.emit_op(Op::MEMORY_FILL, 0);
        // Copy 0 bytes — destination stays untouched
        push_i32(c, 0);
        push_i32(c, 10);
        push_i32(c, 0);
        c.emit_op(Op::MEMORY_COPY, 0);
    });
    assert_eq!(read_byte(&vm, 0), 0x00);
}

#[test]
fn memory_copy_overlapping_forward() {
    // src=[0..4]=0x77, dst=[2..6] — overlap, must copy correctly
    let vm = run_with_memory(256, |c| {
        push_i32(c, 0);
        push_i32(c, 0x77);
        push_i32(c, 4);
        c.emit_op(Op::MEMORY_FILL, 0);
        push_i32(c, 2);  // dst
        push_i32(c, 0);  // src
        push_i32(c, 4);  // count
        c.emit_op(Op::MEMORY_COPY, 0);
    });
    assert_eq!(read_byte(&vm, 2), 0x77);
    assert_eq!(read_byte(&vm, 5), 0x77);
}

// ── data.drop (0xFC 0x09) ────────────────────────────────────────────

#[test]
fn data_drop_does_not_trap() {
    // data.drop is a stub in the VM — verify it executes without error
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op_u8(Op::DATA_DROP, 0, 0); // data segment index 0
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("data.drop should not trap");
}

// ── table.copy (0xFC 0x0E) ───────────────────────────────────────────

#[test]
fn table_copy_copies_entries() {
    let mut vm = VM::new();
    // Set up func_table with 8 slots: [0,1,2,3,null,null,null,null]
    vm.func_table = (0..4)
        .map(|i| Value::I32(i))
        .chain(std::iter::repeat(Value::Null).take(4))
        .collect();

    let mut chunk = Chunk::new("<script>");
    // table.copy dst_table=0; stack: dst=4, src=0, count=4
    push_i32(&mut chunk, 4); // dst
    push_i32(&mut chunk, 0); // src
    push_i32(&mut chunk, 4); // count
    chunk.emit_op_u8(Op::TABLE_COPY, 0, 0); // dst_table index = 0
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("table.copy should not trap");

    for i in 0..4 {
        assert_eq!(vm.func_table[4 + i].as_i32(), i as i32);
    }
}

#[test]
fn table_copy_overlapping_backward() {
    // Copy [0..4] to [1..5] — dst > src, must not clobber before reading
    let mut vm = VM::new();
    vm.func_table = (0..8).map(|i| Value::I32(i)).collect();

    let mut chunk = Chunk::new("<script>");
    push_i32(&mut chunk, 1); // dst
    push_i32(&mut chunk, 0); // src
    push_i32(&mut chunk, 4); // count
    chunk.emit_op_u8(Op::TABLE_COPY, 0, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("table.copy overlapping should not trap");

    assert_eq!(vm.func_table[1].as_i32(), 0);
    assert_eq!(vm.func_table[2].as_i32(), 1);
    assert_eq!(vm.func_table[3].as_i32(), 2);
    assert_eq!(vm.func_table[4].as_i32(), 3);
}

// ── elem.drop (0xFC 0x0D) ────────────────────────────────────────────

#[test]
fn elem_drop_does_not_trap() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op_u8(Op::ELEM_DROP, 0, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("elem.drop should not trap");
}

// ── table.init (0xFC 0x0C) ───────────────────────────────────────────

#[test]
fn table_init_stub_does_not_trap() {
    let mut vm = VM::new();
    vm.func_table.resize(8, Value::Null);
    let mut chunk = Chunk::new("<script>");
    push_i32(&mut chunk, 0); // dst
    push_i32(&mut chunk, 0); // src offset
    push_i32(&mut chunk, 0); // count
    chunk.emit_op_u8(Op::TABLE_INIT, 0, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("table.init should not trap");
}
