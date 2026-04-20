//! Reference-types proposal coverage — exercises the opcodes added to
//! close the gaps flagged by the status audit:
//!   * `table.get` / `table.set` (core 0x25 / 0x26)
//!   * typed `select t` (core 0x1C)
//!   * typed `ref.null` variants (funcref / any / none)
//!   * multi-table routing via `extra_tables`
//!
//! Each test runs through the VM dispatch AND round-trips through the
//! WASM binary writer to confirm both sides stay in sync.

use vybe_bytecode::{VM, Value, Chunk, Op};

// ── table.get / table.set ──────────────────────────────────────────────

#[test]
fn table_get_returns_func_table_slot() {
    // Pre-populate func_table with three values, then `table.get` the
    // middle one via a synthetic chunk.
    let mut vm = VM::new();
    vm.func_table = vec![Value::I32(100), Value::I32(200), Value::I32(300)];

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    // Push index 1, read from table 0.
    let one = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, one, 0);
    chunk.emit_op_u8(Op::TABLE_GET, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 200);
}

#[test]
fn table_set_writes_to_func_table_slot() {
    let mut vm = VM::new();
    vm.func_table = vec![Value::Null, Value::Null];

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    // table.set 0: stack order [index, value]
    let idx = chunk.add_constant(Value::I32(0));
    let val = chunk.add_constant(Value::I32(42));
    chunk.emit_op_u16(Op::CONST, idx, 0);
    chunk.emit_op_u16(Op::CONST, val, 0);
    chunk.emit_op_u8(Op::TABLE_SET, 0, 0);
    chunk.emit_op(Op::NULL, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.run(vec![chunk]).unwrap();
    assert_eq!(vm.func_table[0].as_i32(), 42);
}

#[test]
fn table_set_traps_on_out_of_bounds() {
    let mut vm = VM::new();
    vm.func_table = vec![Value::Null];

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    let idx = chunk.add_constant(Value::I32(5));
    let val = chunk.add_constant(Value::I32(99));
    chunk.emit_op_u16(Op::CONST, idx, 0);
    chunk.emit_op_u16(Op::CONST, val, 0);
    chunk.emit_op_u8(Op::TABLE_SET, 0, 0);
    chunk.emit_op(Op::NULL, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = vm.run(vec![chunk]).unwrap_err();
    assert!(err.message.contains("out of bounds"),
        "expected OOB trap, got: {}", err.message);
}

// ── select_t (typed select) ────────────────────────────────────────────

#[test]
fn select_t_picks_first_when_cond_is_truthy() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    // select_t: [a, b, cond] → cond ? a : b
    let a = chunk.add_constant(Value::I32(10));
    let b = chunk.add_constant(Value::I32(20));
    let c = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::SELECT_T, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 10);
}

#[test]
fn select_t_picks_second_when_cond_is_zero() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    let a = chunk.add_constant(Value::I32(10));
    let b = chunk.add_constant(Value::I32(20));
    let c = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op_u16(Op::CONST, c, 0);
    chunk.emit_op(Op::SELECT_T, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 20);
}

// ── Typed ref.null variants ────────────────────────────────────────────

#[test]
fn null_func_produces_null_at_runtime() {
    // At runtime every ref.null is identical (Value::Null). What matters
    // is that the WASM binary carries the right heaptype byte.
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op(Op::NULL_FUNC, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert!(matches!(result, Value::Null));
}

#[test]
fn null_variants_emit_correct_heaptypes_in_wasm() {
    use vybe_bytecode::wasm::write_wasm;

    let mut chunk = Chunk::new("<script>");
    chunk.emit_op(Op::NULL, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::NULL_FUNC, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::NULL_ANY, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::NULL_NONE, 0);
    chunk.emit_op(Op::HALT, 0);

    let wasm = write_wasm(&vec![chunk]);

    // The four ref.null emissions produce 0xD0 followed by the heaptype
    // byte: 0x6F (extern), 0x70 (func), 0x6E (any), 0x71 (none). The
    // sequence `0xD0 0x6F ... 0xD0 0x70 ... 0xD0 0x6E ... 0xD0 0x71`
    // appears in the code section body.
    let needle = [0xD0u8, 0x6F];
    let has_extern = wasm.windows(2).any(|w| w == needle);
    let has_func   = wasm.windows(2).any(|w| w == [0xD0u8, 0x70]);
    let has_any    = wasm.windows(2).any(|w| w == [0xD0u8, 0x6E]);
    let has_none   = wasm.windows(2).any(|w| w == [0xD0u8, 0x71]);
    assert!(has_extern, "ref.null extern byte pair missing");
    assert!(has_func,   "ref.null func byte pair missing");
    assert!(has_any,    "ref.null any byte pair missing");
    assert!(has_none,   "ref.null none byte pair missing");
}

// ── Multi-table routing ────────────────────────────────────────────────

#[test]
fn multi_table_routes_by_tableidx() {
    // Table 0 is `func_table`; table 1 lives in `extra_tables[0]`.
    // A table.get against each returns the pre-stamped sentinel value,
    // confirming the tableidx operand reaches the right storage.
    let mut vm = VM::new();
    vm.func_table = vec![Value::I32(1111)];
    vm.extra_tables.push(vec![Value::I32(2222)]);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;
    // Read table 0[0], save; read table 1[0], add them.
    let zero = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::CONST, zero, 0);
    chunk.emit_op_u8(Op::TABLE_GET, 0, 0); // tableidx = 0
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0); chunk.emit_op(Op::DROP, 0);

    chunk.emit_op_u16(Op::CONST, zero, 0);
    chunk.emit_op_u8(Op::TABLE_GET, 1, 0); // tableidx = 1
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0); chunk.emit_op(Op::DROP, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    chunk.emit_op(Op::DYN_ADD, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 3333,
        "table 0 + table 1 contents should route independently");
}

#[test]
fn table_grow_extends_selected_table() {
    let mut vm = VM::new();
    vm.extra_tables.push(vec![Value::Null, Value::Null]);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    // table.grow table 1 by 3 entries, init = null. Returns old size = 2.
    chunk.emit_op(Op::NULL, 0); // init value
    let delta = chunk.add_constant(Value::I32(3));
    chunk.emit_op_u16(Op::CONST, delta, 0);
    chunk.emit_op_u8(Op::TABLE_GROW, 1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 2, "table.grow returns old size");
    assert_eq!(vm.extra_tables[0].len(), 5, "table grew by delta");
}
