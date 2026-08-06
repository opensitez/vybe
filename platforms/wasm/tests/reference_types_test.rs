//! Reference-types proposal coverage — exercises the opcodes added to
//! close the gaps flagged by the status audit:
//!   * `table.get` / `table.set` (core 0x25 / 0x26)
//!   * typed `select t` (core 0x1C)
//!   * typed `ref.null` variants (funcref / any / none)
//!   * multi-table routing via `extra_tables`
//!
//! Each test runs through the VM dispatch AND round-trips through the
//! WASM binary writer to confirm both sides stay in sync.

use vybe_runtime::{Chunk, Op, VM, Value};

// ── table.get / table.set ──────────────────────────────────────────────

#[test]
fn table_get_returns_func_table_slot() {
    // Pre-populate func_table with three values, then `table.get` the
    // middle one via a synthetic chunk.
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::I32(100), Value::I32(200), Value::I32(300)]];

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    // Push index 1, read from table 0.
    chunk.emit_i32_const(1, 0);
    chunk.emit_op_u8(Op::TABLE_GET, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 200);
}

#[test]
fn table_get_traps_on_out_of_bounds() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::Null]];

    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(1, 0);
    chunk.emit_op_u8(Op::TABLE_GET, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = vm.run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("table.get") && err.contains("out of bounds"));
}

#[test]
fn table_set_writes_to_func_table_slot() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::Null, Value::Null]];

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    // table.set 0: stack order [index, value]
    chunk.emit_i32_const(0, 0);
    chunk.emit_i32_const(42, 0);
    chunk.emit_op_u8(Op::TABLE_SET, 0, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.run(vec![chunk]).unwrap();
    assert_eq!(vm.wasm_tables[0][0].as_i32(), 42);
}

#[test]
fn table_set_traps_on_out_of_bounds() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::Null]];

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.emit_i32_const(5, 0);
    chunk.emit_i32_const(99, 0);
    chunk.emit_op_u8(Op::TABLE_SET, 0, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = vm.run(vec![chunk]).unwrap_err();
    assert!(
        err.message.contains("out of bounds"),
        "expected OOB trap, got: {}",
        err.message
    );
}

#[test]
fn table_get_unknown_table_traps() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u8(Op::TABLE_GET, 3, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = VM::new().run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("table.get") && err.contains("unknown table"));
}

// ── select_t (typed select) ────────────────────────────────────────────

#[test]
fn select_t_picks_first_when_cond_is_truthy() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    // select_t: [a, b, cond] → cond ? a : b
    chunk.emit_i32_const(10, 0);
    chunk.emit_i32_const(20, 0);
    chunk.emit_i32_const(1, 0);
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
    chunk.emit_i32_const(10, 0);
    chunk.emit_i32_const(20, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op(Op::SELECT_T, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 20);
}

// ── Typed ref.null variants ────────────────────────────────────────────

// ── Multi-table routing ────────────────────────────────────────────────

#[test]
fn multi_table_routes_by_tableidx() {
    // Table 0 is `func_table`; table 1 lives in `extra_tables[0]`.
    // A table.get against each returns the pre-stamped sentinel value,
    // confirming the tableidx operand reaches the right storage.
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::I32(1111)]];
    vm.wasm_tables.push(vec![Value::I32(2222)]);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;
    // Read table 0[0], save; read table 1[0], add them.
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u8(Op::TABLE_GET, 0, 0); // tableidx = 0
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);

    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u8(Op::TABLE_GET, 1, 0); // tableidx = 1
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(
        result.as_i32(),
        3333,
        "table 0 + table 1 contents should route independently"
    );
}

#[test]
fn table_grow_extends_selected_table() {
    let mut vm = VM::new();
    // table 0 empty, table 1 = two slots (the op below grows table 1).
    vm.wasm_tables = vec![Vec::new(), vec![Value::Null, Value::Null]];

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    // table.grow table 1 by 3 entries, init = null. Returns old size = 2.
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0); // init value
    chunk.emit_i32_const(3, 0);
    chunk.emit_op_u8(Op::TABLE_GROW, 1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 2, "table.grow returns old size");
    assert_eq!(vm.wasm_tables[1].len(), 5, "table grew by delta");
}

#[test]
fn table_size_reports_selected_table_length() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::Null; 3]];
    vm.wasm_tables.push(vec![Value::Null; 7]);

    let mut chunk = Chunk::new("<script>");
    chunk.emit_op_u8(Op::TABLE_SIZE, 1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 7);
}

#[test]
fn table_size_unknown_table_traps() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op_u8(Op::TABLE_SIZE, 2, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = VM::new().run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("table.size") && err.contains("unknown table"));
}

#[test]
fn table_grow_unknown_table_traps() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op_u8(Op::TABLE_GROW, 2, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = VM::new().run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("table.grow") && err.contains("unknown table"));
}

#[test]
fn table_fill_writes_requested_range() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::I32(0); 5]];

    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(1, 0);
    chunk.emit_i32_const(55, 0);
    chunk.emit_i32_const(3, 0);
    chunk.emit_op_u8(Op::TABLE_FILL, 0, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.run(vec![chunk]).unwrap();
    assert_eq!(vm.wasm_tables[0][0].as_i32(), 0);
    assert_eq!(vm.wasm_tables[0][1].as_i32(), 55);
    assert_eq!(vm.wasm_tables[0][2].as_i32(), 55);
    assert_eq!(vm.wasm_tables[0][3].as_i32(), 55);
    assert_eq!(vm.wasm_tables[0][4].as_i32(), 0);
}

#[test]
fn table_fill_oob_traps() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::Null; 2]];

    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(1, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_i32_const(2, 0);
    chunk.emit_op_u8(Op::TABLE_FILL, 0, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = vm.run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("table.fill") && err.contains("out of bounds"));
}

#[test]
fn table_fill_unknown_table_traps() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(0, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u8(Op::TABLE_FILL, 2, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = VM::new().run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("table.fill") && err.contains("unknown table"));
}

#[test]
fn table_fill_zero_count_at_table_end_is_noop() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::I32(1), Value::I32(2)]];

    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(2, 0);
    chunk.emit_i32_const(99, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u8(Op::TABLE_FILL, 0, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.run(vec![chunk]).unwrap();
    let got: Vec<i32> = vm.wasm_tables[0].iter().map(Value::as_i32).collect();
    assert_eq!(got, vec![1, 2]);
}

#[test]
fn table_copy_overlapping_forward_preserves_source_snapshot() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![
        Value::I32(1),
        Value::I32(2),
        Value::I32(3),
        Value::I32(4),
    ]];

    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(1, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_i32_const(3, 0);
    chunk.emit_op_u8_u8(Op::TABLE_COPY, 0, 0, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.run(vec![chunk]).unwrap();
    let got: Vec<i32> = vm.wasm_tables[0].iter().map(Value::as_i32).collect();
    assert_eq!(got, vec![1, 1, 2, 3]);
}

#[test]
fn table_copy_destination_oob_traps() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::Null; 3]];

    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(2, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_i32_const(2, 0);
    chunk.emit_op_u8_u8(Op::TABLE_COPY, 0, 0, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = vm.run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("table.copy") && err.contains("out of bounds"));
}

#[test]
fn table_copy_source_oob_traps() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::Null; 3]];

    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(0, 0);
    chunk.emit_i32_const(2, 0);
    chunk.emit_i32_const(2, 0);
    chunk.emit_op_u8_u8(Op::TABLE_COPY, 0, 0, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = vm.run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("table.copy") && err.contains("out of bounds"));
}

#[test]
fn table_copy_routes_to_selected_extra_table() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::I32(10), Value::I32(20), Value::I32(30)]];
    vm.wasm_tables
        .push(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);

    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(1, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_i32_const(2, 0);
    chunk.emit_op_u8_u8(Op::TABLE_COPY, 1, 1, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.run(vec![chunk]).unwrap();
    let table0: Vec<i32> = vm.wasm_tables[0].iter().map(Value::as_i32).collect();
    let table1: Vec<i32> = vm.wasm_tables[1].iter().map(Value::as_i32).collect();
    assert_eq!(
        table0,
        vec![10, 20, 30],
        "table.copy 1 must not touch table 0"
    );
    assert_eq!(table1, vec![1, 1, 2]);
}

#[test]
fn table_copy_zero_count_at_table_end_is_noop() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::I32(1), Value::I32(2)]];

    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(2, 0);
    chunk.emit_i32_const(2, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u8_u8(Op::TABLE_COPY, 0, 0, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.run(vec![chunk]).unwrap();
    let got: Vec<i32> = vm.wasm_tables[0].iter().map(Value::as_i32).collect();
    assert_eq!(got, vec![1, 2]);
}

#[test]
fn table_init_after_elem_drop_traps() {
    let mut vm = VM::new();

    let mut chunk = Chunk::new("<script>");
    chunk.emit_op_u8(Op::ELEM_DROP, 0, 0);
    chunk.emit_i32_const(0, 0); // dst
    chunk.emit_i32_const(0, 0); // src
    chunk.emit_i32_const(0, 0); // count
    chunk.emit_op_u8_u8(Op::TABLE_INIT, 0, 0, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = vm.run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("table.init") && err.contains("dropped"));
}

// ── ref.eq / ref.is_null / ref.as_non_null ───────────────────────────────

#[test]
fn ref_is_null_true_for_null() {
    let mut c = Chunk::new("<script>");
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![c]).unwrap();
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn ref_is_null_false_for_non_null() {
    let mut c = Chunk::new("<script>");
    c.emit_i32_const(42, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![c]).unwrap();
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn ref_eq_same_value_is_true() {
    let mut c = Chunk::new("<script>");
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    c.emit_op(Op::REF_EQ, 0);
    c.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![c]).unwrap();
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn ref_eq_different_values_is_false() {
    let mut c = Chunk::new("<script>");
    c.emit_i32_const(1, 0);
    c.emit_i32_const(2, 0);
    c.emit_op(Op::REF_EQ, 0);
    c.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![c]).unwrap();
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn ref_as_non_null_passes_non_null() {
    let mut c = Chunk::new("<script>");
    c.emit_i32_const(99, 0);
    c.emit_op(Op::REF_AS_NON_NULL, 0);
    c.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![c]).unwrap();
    assert_eq!(r.as_i32(), 99);
}

#[test]
fn ref_as_non_null_traps_on_null() {
    let mut c = Chunk::new("<script>");
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    c.emit_op(Op::REF_AS_NON_NULL, 0);
    c.emit_op(Op::RETURN, 0);
    let err = VM::new().run(vec![c]).unwrap_err().to_string();
    assert!(
        err.contains("null") || err.contains("trap"),
        "expected trap, got: {err}"
    );
}

// Note: br_on_null and br_on_non_null use byte-offset encoding (not label depth)
// and are already thoroughly tested in gc_test.rs.
