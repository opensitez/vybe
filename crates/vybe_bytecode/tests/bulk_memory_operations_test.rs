//! Tests for the bulk-memory-operations WASM proposal.
//! Spec: `proposals/bulk-memory-operations/`, opcodes 0xFC 0x08–0x0E.
//!
//! Covers VM semantics for memory/table copy, fill, init, and drop.

use vybe_bytecode::value::ObjectKind;
use vybe_bytecode::{Chunk, Op, VM, Value};

fn write_leb_u32(out: &mut Vec<u8>, mut value: u32) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        out.push(byte);
        if value == 0 {
            break;
        }
    }
}

fn push_section(out: &mut Vec<u8>, id: u8, payload: &[u8]) {
    out.push(id);
    write_leb_u32(out, payload.len() as u32);
    out.extend_from_slice(payload);
}

fn standard_module_with_sections(sections: &[(u8, Vec<u8>)]) -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);
    for (id, payload) in sections {
        push_section(&mut bytes, *id, payload);
    }
    bytes
}

fn code_section_for_body(body_ops: &[u8]) -> Vec<u8> {
    let mut body = Vec::new();
    body.push(0x00); // local declarations
    body.extend_from_slice(body_ops);
    body.push(0x0b);

    let mut code = Vec::new();
    code.push(0x01); // function count
    write_leb_u32(&mut code, body.len() as u32);
    code.extend_from_slice(&body);
    code
}

fn run_with_memory(mem_size: usize, emit: impl FnOnce(&mut Chunk)) -> VM {
    let mut vm = VM::new();
    vm.memory.resize(mem_size, 0);
    let mut chunk = Chunk::new("<script>");
    emit(&mut chunk);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("VM execution failed");
    vm
}

fn run_with_memory_err(mem_size: usize, emit: impl FnOnce(&mut Chunk)) -> String {
    let mut vm = VM::new();
    vm.memory.resize(mem_size, 0);
    let mut chunk = Chunk::new("<script>");
    emit(&mut chunk);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).unwrap_err().to_string()
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
fn memory_fill_zero_count_at_memory_end_is_noop() {
    let vm = run_with_memory(64, |c| {
        push_i32(c, 64);
        push_i32(c, 0xFF);
        push_i32(c, 0);
        c.emit_op(Op::MEMORY_FILL, 0);
    });
    assert_eq!(read_byte(&vm, 63), 0x00);
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

#[test]
fn memory_fill_oob_traps() {
    let err = run_with_memory_err(4, |c| {
        push_i32(c, 2);
        push_i32(c, 0xAA);
        push_i32(c, 3);
        c.emit_op(Op::MEMORY_FILL, 0);
    });
    assert!(err.contains("out of bounds") || err.contains("trap"));
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
        push_i32(c, 0); // dst
        push_i32(c, 20); // src
        push_i32(c, 4); // count
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
fn memory_copy_zero_count_at_memory_end_is_noop() {
    let vm = run_with_memory(64, |c| {
        push_i32(c, 0);
        push_i32(c, 0xAA);
        push_i32(c, 1);
        c.emit_op(Op::MEMORY_FILL, 0);
        push_i32(c, 64);
        push_i32(c, 64);
        push_i32(c, 0);
        c.emit_op(Op::MEMORY_COPY, 0);
    });
    assert_eq!(read_byte(&vm, 0), 0xAA);
}

#[test]
fn memory_copy_overlapping_forward() {
    // src=[0..4]=0x77, dst=[2..6] — overlap, must copy correctly
    let vm = run_with_memory(256, |c| {
        push_i32(c, 0);
        push_i32(c, 0x77);
        push_i32(c, 4);
        c.emit_op(Op::MEMORY_FILL, 0);
        push_i32(c, 2); // dst
        push_i32(c, 0); // src
        push_i32(c, 4); // count
        c.emit_op(Op::MEMORY_COPY, 0);
    });
    assert_eq!(read_byte(&vm, 2), 0x77);
    assert_eq!(read_byte(&vm, 5), 0x77);
}

#[test]
fn memory_copy_source_oob_traps() {
    let err = run_with_memory_err(4, |c| {
        push_i32(c, 0);
        push_i32(c, 2);
        push_i32(c, 3);
        c.emit_op(Op::MEMORY_COPY, 0);
    });
    assert!(err.contains("out of bounds") || err.contains("trap"));
}

#[test]
fn memory_copy_destination_oob_traps() {
    let err = run_with_memory_err(4, |c| {
        push_i32(c, 2);
        push_i32(c, 0);
        push_i32(c, 3);
        c.emit_op(Op::MEMORY_COPY, 0);
    });
    assert!(err.contains("out of bounds") || err.contains("trap"));
}

// ── data.drop (0xFC 0x09) ────────────────────────────────────────────

#[test]
fn data_drop_does_not_trap() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op_u8(Op::DATA_DROP, 0, 0); // data segment index 0
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("data.drop should not trap");
}

#[test]
fn memory_init_zero_count_without_drop_is_noop() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    push_i32(&mut chunk, 0); // dst
    push_i32(&mut chunk, 0); // src offset
    push_i32(&mut chunk, 0); // count
    chunk.emit_op_u8(Op::MEMORY_INIT, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.run(vec![chunk])
        .expect("zero-count memory.init should not trap before data.drop");
}

#[test]
fn decoded_standard_memory_init_copies_passive_data_segment() {
    let wasm = standard_module_with_sections(&[
        (1, vec![0x01, 0x60, 0x00, 0x01, 0x7f]),
        (3, vec![0x01, 0x00]),
        (5, vec![0x01, 0x00, 0x01]),
        (12, vec![0x01]),
        (
            10,
            code_section_for_body(&[
                0x41, 0x00, // i32.const 0: dst
                0x41, 0x00, // i32.const 0: src offset
                0x41, 0x04, // i32.const 4: byte count
                0xfc, 0x08, 0x00, 0x00, // memory.init dataidx=0 memidx=0
                0x41, 0x00, // i32.const 0
                0x28, 0x02, 0x00, // i32.load align=2 offset=0
            ]),
        ),
        (11, vec![0x01, 0x01, 0x04, 0x09, 0x08, 0x07, 0x06]),
    ]);
    let mut chunks = vybe_bytecode::wasm::read_wasm(&wasm).expect("standard wasm should decode");
    assert_eq!(chunks[0].data_segments, vec![vec![9, 8, 7, 6]]);

    let function = chunks.remove(1);
    let result = VM::new()
        .run(vec![function])
        .expect("decoded memory.init should execute");
    assert_eq!(result.as_i32(), 0x0607_0809);
}

#[test]
fn decoded_standard_active_data_segment_initializes_memory() {
    let wasm = standard_module_with_sections(&[
        (1, vec![0x01, 0x60, 0x00, 0x01, 0x7f]),
        (3, vec![0x01, 0x00]),
        (5, vec![0x01, 0x00, 0x01]),
        (
            10,
            code_section_for_body(&[
                0x41, 0x04, // i32.const 4
                0x28, 0x02, 0x00, // i32.load align=2 offset=0
            ]),
        ),
        (
            11,
            vec![
                0x01, // data segment count
                0x00, // active memory 0
                0x41, 0x04, 0x0b, // offset = 4
                0x04, // byte count
                0x01, 0x02, 0x03, 0x04,
            ],
        ),
    ]);
    let mut chunks = vybe_bytecode::wasm::read_wasm(&wasm).expect("standard wasm should decode");
    assert_eq!(chunks[0].active_data_segments.len(), 1);

    let function = chunks.remove(1);
    let result = VM::new()
        .run(vec![function])
        .expect("active data segment should instantiate");
    assert_eq!(result.as_i32(), 0x0403_0201);
}

#[test]
fn memory_init_after_data_drop_traps() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op_u8(Op::DATA_DROP, 0, 0);
    push_i32(&mut chunk, 0); // dst
    push_i32(&mut chunk, 0); // src offset
    push_i32(&mut chunk, 0); // count
    chunk.emit_op_u8(Op::MEMORY_INIT, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = vm
        .run(vec![chunk])
        .expect_err("memory.init after data.drop must trap");
    assert!(
        err.to_string().contains("data segment dropped"),
        "unexpected trap: {err}"
    );
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
    chunk.emit_op_u8_u8(Op::TABLE_COPY, 0, 0, 0); // dst_table=0, src_table=0
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
    chunk.emit_op_u8_u8(Op::TABLE_COPY, 0, 0, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk])
        .expect("table.copy overlapping should not trap");

    assert_eq!(vm.func_table[1].as_i32(), 0);
    assert_eq!(vm.func_table[2].as_i32(), 1);
    assert_eq!(vm.func_table[3].as_i32(), 2);
    assert_eq!(vm.func_table[4].as_i32(), 3);
}

#[test]
fn decoded_standard_table_copy_preserves_distinct_source_and_destination_tables() {
    let wasm = standard_module_with_sections(&[
        (1, vec![0x01, 0x60, 0x00, 0x01, 0x6f]),
        (3, vec![0x01, 0x00]),
        (4, vec![0x02, 0x70, 0x00, 0x01, 0x70, 0x00, 0x01]),
        (
            9,
            vec![
                0x01, // segment count
                0x00, // active, table 0, elemkind funcidx
                0x41, 0x00, 0x0b, // offset = 0
                0x01, // element count
                0x00, // function index 0
            ],
        ),
        (
            10,
            code_section_for_body(&[
                0x41, 0x00, // dst index
                0x41, 0x00, // src index
                0x41, 0x01, // count
                0xfc, 0x0e, 0x01, 0x00, // table.copy dst=1 src=0
                0x41, 0x00, // table.get index
                0x25, 0x01, // table.get 1
            ]),
        ),
    ]);
    let mut chunks = vybe_bytecode::wasm::read_wasm(&wasm).expect("standard wasm should decode");
    let function = chunks.remove(1);
    let result = VM::new()
        .run(vec![function])
        .expect("decoded table.copy should execute across tables");
    assert!(
        matches!(&result, Value::Object(obj) if matches!(obj.lock().unwrap().kind, ObjectKind::Function(_)))
    );
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
    vm.set_elem_segment(0, vec![Value::I32(10), Value::I32(11), Value::I32(12)]);
    let mut chunk = Chunk::new("<script>");
    push_i32(&mut chunk, 2); // dst
    push_i32(&mut chunk, 1); // src offset
    push_i32(&mut chunk, 2); // count
    chunk.emit_op_u8_u8(Op::TABLE_INIT, 0, 0, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("table.init should not trap");

    assert_eq!(vm.func_table[2].as_i32(), 11);
    assert_eq!(vm.func_table[3].as_i32(), 12);
}

#[test]
fn decoded_standard_table_init_copies_passive_element_segment() {
    let wasm = standard_module_with_sections(&[
        (1, vec![0x01, 0x60, 0x00, 0x01, 0x6f]),
        (3, vec![0x01, 0x00]),
        (4, vec![0x01, 0x70, 0x00, 0x01]),
        (
            9,
            vec![
                0x01, // segment count
                0x05, // passive expressions
                0x70, // funcref
                0x01, // element count
                0xd2, 0x00, 0x0b, // ref.func 0; end
            ],
        ),
        (
            10,
            code_section_for_body(&[
                0x41, 0x00, // i32.const 0: table dst
                0x41, 0x00, // i32.const 0: element src
                0x41, 0x01, // i32.const 1: count
                0xfc, 0x0c, 0x00, 0x00, // table.init elemidx=0 tableidx=0
                0x41, 0x00, // i32.const 0
                0x25, 0x00, // table.get 0
            ]),
        ),
    ]);
    let mut chunks = vybe_bytecode::wasm::read_wasm(&wasm).expect("standard wasm should decode");
    assert_eq!(chunks[0].table_min_sizes, vec![1]);
    assert_eq!(chunks[0].elem_segments, vec![vec![Value::I32(0)]]);

    let function = chunks.remove(1);
    let mut vm = VM::new();
    let result = vm
        .run(vec![function])
        .expect("decoded table.init should execute");

    assert!(
        matches!(&result, Value::Object(obj) if matches!(obj.lock().unwrap().kind, ObjectKind::Function(_)))
    );
    assert!(
        matches!(&vm.func_table[0], Value::Object(obj) if matches!(obj.lock().unwrap().kind, ObjectKind::Function(_)))
    );
}

#[test]
fn decoded_standard_table_init_uses_encoded_table_index() {
    let wasm = standard_module_with_sections(&[
        (1, vec![0x01, 0x60, 0x00, 0x01, 0x6f]),
        (3, vec![0x01, 0x00]),
        (4, vec![0x02, 0x70, 0x00, 0x01, 0x70, 0x00, 0x01]),
        (
            9,
            vec![
                0x01, // segment count
                0x05, // passive expressions
                0x70, // funcref
                0x01, // element count
                0xd2, 0x00, 0x0b, // ref.func 0; end
            ],
        ),
        (
            10,
            code_section_for_body(&[
                0x41, 0x00, // table dst
                0x41, 0x00, // element src
                0x41, 0x01, // count
                0xfc, 0x0c, 0x00, 0x01, // table.init elemidx=0 tableidx=1
                0x41, 0x00, // table.get index
                0x25, 0x01, // table.get 1
            ]),
        ),
    ]);
    let mut chunks = vybe_bytecode::wasm::read_wasm(&wasm).expect("standard wasm should decode");
    let function = chunks.remove(1);
    let result = VM::new()
        .run(vec![function])
        .expect("decoded table.init should target encoded table");
    assert!(
        matches!(&result, Value::Object(obj) if matches!(obj.lock().unwrap().kind, ObjectKind::Function(_)))
    );
}

#[test]
fn decoded_standard_active_element_segment_initializes_table() {
    let wasm = standard_module_with_sections(&[
        (1, vec![0x01, 0x60, 0x00, 0x01, 0x6f]),
        (3, vec![0x01, 0x00]),
        (4, vec![0x01, 0x70, 0x00, 0x01]),
        (
            9,
            vec![
                0x01, // segment count
                0x00, // active, table 0, elemkind funcidx
                0x41, 0x00, 0x0b, // offset = 0
                0x01, // element count
                0x00, // function index 0
            ],
        ),
        (
            10,
            code_section_for_body(&[
                0x41, 0x00, // i32.const 0
                0x25, 0x00, // table.get 0
            ]),
        ),
    ]);
    let mut chunks = vybe_bytecode::wasm::read_wasm(&wasm).expect("standard wasm should decode");
    assert_eq!(chunks[0].active_elem_segments.len(), 1);

    let function = chunks.remove(1);
    let result = VM::new()
        .run(vec![function])
        .expect("active element segment should instantiate");
    assert!(
        matches!(&result, Value::Object(obj) if matches!(obj.lock().unwrap().kind, ObjectKind::Function(_)))
    );
}

#[test]
fn table_init_source_oob_traps() {
    let mut vm = VM::new();
    vm.func_table.resize(8, Value::Null);
    vm.set_elem_segment(0, vec![Value::I32(10)]);
    let mut chunk = Chunk::new("<script>");
    push_i32(&mut chunk, 0); // dst
    push_i32(&mut chunk, 0); // src offset
    push_i32(&mut chunk, 2); // count exceeds segment
    chunk.emit_op_u8_u8(Op::TABLE_INIT, 0, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = vm
        .run(vec![chunk])
        .expect_err("table.init source OOB must trap");
    assert!(err.to_string().contains("source out of bounds"));
}

#[test]
fn table_init_destination_oob_traps() {
    let mut vm = VM::new();
    vm.func_table.resize(1, Value::Null);
    vm.set_elem_segment(0, vec![Value::I32(10), Value::I32(11)]);
    let mut chunk = Chunk::new("<script>");
    push_i32(&mut chunk, 0); // dst
    push_i32(&mut chunk, 0); // src offset
    push_i32(&mut chunk, 2); // count exceeds table
    chunk.emit_op_u8_u8(Op::TABLE_INIT, 0, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = vm
        .run(vec![chunk])
        .expect_err("table.init destination OOB must trap");
    assert!(err.to_string().contains("destination out of bounds"));
}

#[test]
fn table_init_after_elem_drop_traps() {
    let mut vm = VM::new();
    vm.func_table.resize(8, Value::Null);
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op_u8(Op::ELEM_DROP, 0, 0);
    push_i32(&mut chunk, 0); // dst
    push_i32(&mut chunk, 0); // src offset
    push_i32(&mut chunk, 0); // count
    chunk.emit_op_u8_u8(Op::TABLE_INIT, 0, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = vm
        .run(vec![chunk])
        .expect_err("table.init after elem.drop must trap");
    assert!(
        err.to_string().contains("element segment dropped"),
        "unexpected trap: {err}"
    );
}
