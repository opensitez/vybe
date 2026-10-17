//! Tests for the bulk-memory-operations WASM proposal.
//! Spec: `proposals/bulk-memory-operations/`, opcodes 0xFC 0x08–0x0E.
//!
//! Covers VM semantics for memory/table copy, fill, init, and drop.

use vybe_runtime::value::ObjectKind;
use vybe_runtime::{Chunk, Op, VM, Value};

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
    c.emit_i32_const(v, 0);
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
        c.emit_op_idx(Op::MEMORY_FILL, 0u32, 0);
    });
    assert_eq!(read_byte(&vm, 9), 0x00); // before range untouched
    for addr in 10..15 {
        assert_eq!(read_byte(&vm, addr), 0xAB);
    }
    assert_eq!(read_byte(&vm, 15), 0x00); // after range untouched
}

#[test]
fn memory_fill_and_init_boxed_locals_emit_valid_wasm() {
    for init in [false, true] {
        for memory_idx in [0u32, 1u32] {
            let mut chunk = Chunk::new("<boxed-bulk-memory>");
            chunk.memory_min_pages.extend([1, 1]);
            if init {
                chunk.data_segments.push(vec![1, 2, 3, 4]);
            }
            let dst = chunk.alloc_scratch(1);
            let middle = chunk.alloc_scratch(1);
            let len = chunk.alloc_scratch(1);
            for (slot, value) in [(dst, 16), (middle, 0), (len, 4)] {
                chunk.emit_i32_const(value, 0);
                chunk.emit_op_u16(Op::LOCAL_SET, slot, 0);
            }
            for slot in [dst, middle, len] {
                chunk.emit_op_u16(Op::LOCAL_GET, slot, 0);
            }
            if init {
                chunk.emit_op_idx_idx(Op::MEMORY_INIT, 0u32, memory_idx, 0);
            } else {
                chunk.emit_op_idx(Op::MEMORY_FILL, memory_idx, 0);
            }
            chunk.emit_i32_const(0, 0);
            chunk.emit_op(Op::RETURN, 0);

            let binary = vybe_platform_wasm::write_wasm(&[chunk]);
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&binary)
                .expect("boxed bulk-memory operands must lower to valid standard Wasm");
        }
    }
}

#[test]
fn memory64_grow_and_size_emit_valid_wasm() {
    let mut chunk = Chunk::new("<memory64-grow-size>");
    chunk.memory_min_pages.extend([1, 1]);
    chunk.memory_is_64 = vec![false, true];
    chunk.emit_i64_const(1, 0);
    chunk.emit_op_idx(Op::MEMORY_GROW, 1u32, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op_idx(Op::MEMORY_SIZE, 1u32, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::RETURN, 0);

    let binary = vybe_platform_wasm::write_wasm(&[chunk]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&binary)
        .expect("memory64 size/grow must use i64 operands and results");
}

#[test]
fn memory32_grow_and_size_boxed_values_emit_valid_wasm() {
    for memory_idx in [0u32, 1u32] {
        let mut chunk = Chunk::new("<memory32-grow-size>");
        chunk.memory_min_pages.extend([1, 1]);
        chunk.emit_i32_const(1, 0);
        chunk.emit_op_idx(Op::MEMORY_GROW, memory_idx, 0);
        chunk.emit_op(Op::DROP, 0);
        chunk.emit_op_idx(Op::MEMORY_SIZE, memory_idx, 0);
        chunk.emit_op(Op::DROP, 0);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);

        let binary = vybe_platform_wasm::write_wasm(&[chunk]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .expect("memory32 size/grow must retain boxed-value lowering");
    }
}

#[test]
fn memory64_bulk_operations_emit_valid_wasm() {
    let validate = |mut chunk: Chunk| {
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);
        let binary = vybe_platform_wasm::write_wasm(&[chunk]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .expect("memory64 bulk operands must have their standard Wasm types");
    };
    let new_chunk = || {
        let mut chunk = Chunk::new("<memory64-bulk>");
        chunk.memory_min_pages.extend([1, 1]);
        chunk.memory_is_64 = vec![false, true];
        chunk
    };

    let mut fill = new_chunk();
    fill.emit_i64_const(8, 0);
    fill.emit_i32_const(42, 0);
    fill.emit_i64_const(4, 0);
    fill.emit_op_idx(Op::MEMORY_FILL, 1u32, 0);
    validate(fill);

    let mut init = new_chunk();
    init.data_segments.push(vec![1, 2, 3, 4]);
    init.emit_i64_const(8, 0);
    init.emit_i32_const(0, 0);
    init.emit_i32_const(4, 0);
    init.emit_op_idx_idx(Op::MEMORY_INIT, 0u32, 1u32, 0);
    validate(init);

    for (dst, src) in [(0u32, 1u32), (1, 0), (1, 1)] {
        let mut copy = new_chunk();
        if dst == 1 {
            copy.emit_i64_const(8, 0);
        } else {
            copy.emit_i32_const(8, 0);
        }
        if src == 1 {
            copy.emit_i64_const(4, 0);
        } else {
            copy.emit_i32_const(4, 0);
        }
        if dst == 1 && src == 1 {
            copy.emit_i64_const(4, 0);
        } else {
            copy.emit_i32_const(4, 0);
        }
        copy.emit_op_idx_idx(Op::MEMORY_COPY, dst, src, 0);
        validate(copy);
    }
}

#[test]
fn memory_load_store_address_widths_emit_valid_wasm() {
    let emit_memarg = |chunk: &mut Chunk, align: u32, memory_idx: u32| {
        let marker = 0x80 | align | if memory_idx == 1 { 0x140 } else { 0 };
        chunk.emit_leb_u32(marker, 0);
        chunk.emit_leb_u32(7, 0);
        if memory_idx != 0 {
            chunk.emit_leb_u32(memory_idx, 0);
        }
    };
    for memory_idx in [0u32, 1u32] {
        let mut chunk = Chunk::new("<memory-address-width>");
        chunk.memory_min_pages.extend([1, 1]);
        chunk.memory_is_64 = vec![false, true];
        if memory_idx == 1 {
            chunk.emit_i64_const(8, 0);
        } else {
            chunk.emit_i32_const(8, 0);
        }
        chunk.emit_i32_const(42, 0);
        chunk.emit_op(Op::I32_STORE, 0);
        emit_memarg(&mut chunk, 2, memory_idx);
        if memory_idx == 1 {
            chunk.emit_i64_const(8, 0);
        } else {
            chunk.emit_i32_const(8, 0);
        }
        chunk.emit_op(Op::I32_LOAD, 0);
        emit_memarg(&mut chunk, 2, memory_idx);
        chunk.emit_op(Op::DROP, 0);

        if memory_idx == 1 {
            chunk.emit_i64_const(16, 0);
        } else {
            chunk.emit_i32_const(16, 0);
        }
        chunk.emit_i64_const(123, 0);
        chunk.emit_op(Op::I64_STORE, 0);
        emit_memarg(&mut chunk, 3, memory_idx);
        if memory_idx == 1 {
            chunk.emit_i64_const(16, 0);
        } else {
            chunk.emit_i32_const(16, 0);
        }
        chunk.emit_op(Op::I64_LOAD, 0);
        emit_memarg(&mut chunk, 3, memory_idx);
        chunk.emit_op(Op::DROP, 0);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);

        let store_at = chunk
            .code
            .windows(4)
            .position(|bytes| bytes == Op::I32_STORE.encode())
            .expect("i32.store opcode");
        let mut memarg_at = store_at + 4;
        let marker = vybe_runtime::opcode::read_leb_u32(&chunk.code, &mut memarg_at);
        assert_eq!(marker & 0x40 != 0, memory_idx == 1);
        let binary = vybe_platform_wasm::write_wasm(&[chunk]);
        let mut emitted_memargs = Vec::new();
        for payload in wasmparser::Parser::new(0).parse_all(&binary) {
            if let wasmparser::Payload::CodeSectionEntry(body) = payload.expect("code section") {
                let mut ops = body.get_operators_reader().expect("operators");
                while !ops.eof() {
                    match ops.read().expect("operator") {
                        wasmparser::Operator::I32Store { memarg }
                        | wasmparser::Operator::I32Load { memarg }
                        | wasmparser::Operator::I64Store { memarg }
                        | wasmparser::Operator::I64Load { memarg } => {
                            emitted_memargs.push((memarg.memory, memarg.offset));
                        }
                        _ => {}
                    }
                }
            }
        }
        assert_eq!(emitted_memargs, vec![(memory_idx, 7); 4]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .unwrap_or_else(|error| panic!("memory {memory_idx} load/store validation: {error}"));
    }
}

#[test]
fn memory_fill_zero_count_is_noop() {
    let vm = run_with_memory(64, |c| {
        push_i32(c, 0);
        push_i32(c, 0xFF);
        push_i32(c, 0); // count = 0
        c.emit_op_idx(Op::MEMORY_FILL, 0u32, 0);
    });
    assert_eq!(read_byte(&vm, 0), 0x00);
}

#[test]
fn memory_fill_zero_count_at_memory_end_is_noop() {
    let vm = run_with_memory(64, |c| {
        push_i32(c, 64);
        push_i32(c, 0xFF);
        push_i32(c, 0);
        c.emit_op_idx(Op::MEMORY_FILL, 0u32, 0);
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
        c.emit_op_idx(Op::MEMORY_FILL, 0u32, 0);
        push_i32(c, 5);
        push_i32(c, 0x00);
        push_i32(c, 2);
        c.emit_op_idx(Op::MEMORY_FILL, 0u32, 0);
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
        c.emit_op_idx(Op::MEMORY_FILL, 0u32, 0);
    });
    assert!(err.contains("out of bounds") || err.contains("trap"));
}

// ── memory.copy (0xFC 0x0A) ──────────────────────────────────────────

#[test]
fn memory_copy_boxed_locals_emit_valid_wasm() {
    for (dst_memory, src_memory) in [(0u32, 0u32), (1, 0), (0, 1), (1, 1)] {
        let mut chunk = Chunk::new("<boxed-memory-copy>");
        chunk.memory_min_pages.extend([1, 1]);
        let dst = chunk.alloc_scratch(1);
        let src = chunk.alloc_scratch(1);
        let len = chunk.alloc_scratch(1);
        for (slot, value) in [(dst, 16), (src, 0), (len, 4)] {
            chunk.emit_i32_const(value, 0);
            chunk.emit_op_u16(Op::LOCAL_SET, slot, 0);
        }
        for slot in [dst, src, len] {
            chunk.emit_op_u16(Op::LOCAL_GET, slot, 0);
        }
        chunk.emit_op_idx_idx(Op::MEMORY_COPY, dst_memory, src_memory, 0);
        chunk.emit_i32_const(0, 0);
        chunk.emit_op(Op::RETURN, 0);

        let binary = vybe_platform_wasm::write_wasm(&[chunk]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .expect("boxed memory.copy operands must lower to valid standard Wasm");
    }
}

#[test]
fn writer_preserves_multiple_memory_limits_and_index_types() {
    let mut chunk = Chunk::new("<multi-memory-limits>");
    chunk.memory_min_pages = vec![2, 3];
    chunk.memory_max_pages = vec![Some(4), Some(5)];
    chunk.memory_is_64 = vec![false, true];
    chunk.emit_i32_const(0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let binary = vybe_platform_wasm::write_wasm(&[chunk]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&binary)
        .expect("mixed memory32/memory64 declarations must validate");
    let memories = wasmparser::Parser::new(0)
        .parse_all(&binary)
        .find_map(|payload| match payload.expect("valid Wasm section") {
            wasmparser::Payload::MemorySection(reader) => Some(
                reader
                    .into_iter()
                    .collect::<Result<Vec<_>, _>>()
                    .expect("memory declarations"),
            ),
            _ => None,
        })
        .expect("memory section");
    assert_eq!(memories.len(), 2);
    assert_eq!(
        (
            memories[0].initial,
            memories[0].maximum,
            memories[0].memory64
        ),
        (2, Some(4), false)
    );
    assert_eq!(
        (
            memories[1].initial,
            memories[1].maximum,
            memories[1].memory64
        ),
        (3, Some(5), true)
    );
}

#[test]
fn writer_uses_first_chunk_memory_declarations() {
    let mut helper = Chunk::new("helper");
    helper.memory_min_pages = vec![2, 3];
    helper.emit_i32_const(0, 0);
    helper.emit_op(Op::RETURN, 0);

    let mut script = Chunk::new("<script>");
    script.memory_min_pages = vec![9];
    script.emit_i32_const(0, 0);
    script.emit_op(Op::RETURN, 0);

    let binary = vybe_platform_wasm::write_wasm(&[helper, script]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&binary)
        .expect("first-chunk memory declarations must validate");
    let memories = wasmparser::Parser::new(0)
        .parse_all(&binary)
        .find_map(|payload| match payload.expect("valid Wasm section") {
            wasmparser::Payload::MemorySection(reader) => Some(
                reader
                    .into_iter()
                    .collect::<Result<Vec<_>, _>>()
                    .expect("memory declarations"),
            ),
            _ => None,
        })
        .expect("memory section");
    assert_eq!(
        memories
            .iter()
            .map(|memory| memory.initial)
            .collect::<Vec<_>>(),
        vec![2, 3]
    );
}

#[test]
fn memory_copy_non_overlapping() {
    let vm = run_with_memory(256, |c| {
        // Fill src region [20..24] with 0x55
        push_i32(c, 20);
        push_i32(c, 0x55);
        push_i32(c, 4);
        c.emit_op_idx(Op::MEMORY_FILL, 0u32, 0);
        // Copy to dst [0..4]
        push_i32(c, 0); // dst
        push_i32(c, 20); // src
        push_i32(c, 4); // count
        c.emit_op_idx_idx(Op::MEMORY_COPY, 0u32, 0u32, 0);
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
        c.emit_op_idx(Op::MEMORY_FILL, 0u32, 0);
        // Copy 0 bytes — destination stays untouched
        push_i32(c, 0);
        push_i32(c, 10);
        push_i32(c, 0);
        c.emit_op_idx_idx(Op::MEMORY_COPY, 0u32, 0u32, 0);
    });
    assert_eq!(read_byte(&vm, 0), 0x00);
}

#[test]
fn memory_copy_zero_count_at_memory_end_is_noop() {
    let vm = run_with_memory(64, |c| {
        push_i32(c, 0);
        push_i32(c, 0xAA);
        push_i32(c, 1);
        c.emit_op_idx(Op::MEMORY_FILL, 0u32, 0);
        push_i32(c, 64);
        push_i32(c, 64);
        push_i32(c, 0);
        c.emit_op_idx_idx(Op::MEMORY_COPY, 0u32, 0u32, 0);
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
        c.emit_op_idx(Op::MEMORY_FILL, 0u32, 0);
        push_i32(c, 2); // dst
        push_i32(c, 0); // src
        push_i32(c, 4); // count
        c.emit_op_idx_idx(Op::MEMORY_COPY, 0u32, 0u32, 0);
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
        c.emit_op_idx_idx(Op::MEMORY_COPY, 0u32, 0u32, 0);
    });
    assert!(err.contains("out of bounds") || err.contains("trap"));
}

#[test]
fn memory_copy_destination_oob_traps() {
    let err = run_with_memory_err(4, |c| {
        push_i32(c, 2);
        push_i32(c, 0);
        push_i32(c, 3);
        c.emit_op_idx_idx(Op::MEMORY_COPY, 0u32, 0u32, 0);
    });
    assert!(err.contains("out of bounds") || err.contains("trap"));
}

// ── data.drop (0xFC 0x09) ────────────────────────────────────────────

#[test]
fn data_drop_does_not_trap() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op_idx(Op::DATA_DROP, 0u32, 0); // data segment index 0
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
    chunk.emit_op_idx_idx(Op::MEMORY_INIT, 0u32, 0u32, 0);
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
    let mut chunks = vybe_platform_wasm::read_wasm(&wasm).expect("standard wasm should decode");
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
    let mut chunks = vybe_platform_wasm::read_wasm(&wasm).expect("standard wasm should decode");
    assert_eq!(chunks[0].active_data_segments.len(), 1);

    let function = chunks.remove(1);
    let result = VM::new()
        .run(vec![function])
        .expect("active data segment should instantiate");
    assert_eq!(result.as_i32(), 0x0403_0201);
}

#[test]
fn memory_init_after_data_drop_zero_length_ok_nonzero_traps() {
    // Spec (bulk-memory Overview §data.drop): a dropped segment shrinks to
    // zero length and "only a zero-length access at offset zero will not
    // trap". So count=0/src=0 succeeds; any non-zero access is an ordinary
    // out-of-bounds trap (not a distinct "dropped" error class).
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op_idx(Op::DATA_DROP, 0u32, 0);
    push_i32(&mut chunk, 0); // dst
    push_i32(&mut chunk, 0); // src offset
    push_i32(&mut chunk, 0); // count
    chunk.emit_op_idx_idx(Op::MEMORY_INIT, 0u32, 0u32, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk])
        .expect("zero-length memory.init at offset 0 on a dropped segment must not trap");

    let mut vm = VM::new();
    vm.memory.resize(1024, 0);
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op_idx(Op::DATA_DROP, 0u32, 0);
    push_i32(&mut chunk, 0); // dst
    push_i32(&mut chunk, 0); // src offset
    push_i32(&mut chunk, 1); // count — reads past the zero-length segment
    chunk.emit_op_idx_idx(Op::MEMORY_INIT, 0u32, 0u32, 0);
    chunk.emit_op(Op::RETURN, 0);
    let err = vm
        .run(vec![chunk])
        .expect_err("non-zero memory.init on a dropped segment must trap");
    assert!(
        err.to_string().contains("out of bounds"),
        "unexpected trap: {err}"
    );
}

// ── table.copy (0xFC 0x0E) ───────────────────────────────────────────

#[test]
fn table_copy_copies_entries() {
    let mut vm = VM::new();
    // Set up func_table with 8 slots: [0,1,2,3,null,null,null,null]
    vm.wasm_tables = vec![
        (0..4)
            .map(|i| Value::I32(i))
            .chain(std::iter::repeat(Value::Null).take(4))
            .collect(),
    ];

    let mut chunk = Chunk::new("<script>");
    // table.copy dst_table=0; stack: dst=4, src=0, count=4
    push_i32(&mut chunk, 4); // dst
    push_i32(&mut chunk, 0); // src
    push_i32(&mut chunk, 4); // count
    chunk.emit_op_idx_idx(Op::TABLE_COPY, 0u32, 0u32, 0); // dst_table=0, src_table=0
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("table.copy should not trap");

    for i in 0..4 {
        assert_eq!(vm.wasm_tables[0][4 + i].as_i32(), i as i32);
    }
}

#[test]
fn table_copy_overlapping_backward() {
    // Copy [0..4] to [1..5] — dst > src, must not clobber before reading
    let mut vm = VM::new();
    vm.wasm_tables = vec![(0..8).map(|i| Value::I32(i)).collect()];

    let mut chunk = Chunk::new("<script>");
    push_i32(&mut chunk, 1); // dst
    push_i32(&mut chunk, 0); // src
    push_i32(&mut chunk, 4); // count
    chunk.emit_op_idx_idx(Op::TABLE_COPY, 0u32, 0u32, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk])
        .expect("table.copy overlapping should not trap");

    assert_eq!(vm.wasm_tables[0][1].as_i32(), 0);
    assert_eq!(vm.wasm_tables[0][2].as_i32(), 1);
    assert_eq!(vm.wasm_tables[0][3].as_i32(), 2);
    assert_eq!(vm.wasm_tables[0][4].as_i32(), 3);
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
    let mut chunks = vybe_platform_wasm::read_wasm(&wasm).expect("standard wasm should decode");
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
    chunk.emit_op_idx(Op::ELEM_DROP, 0u32, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("elem.drop should not trap");
}

// ── table.init (0xFC 0x0C) ───────────────────────────────────────────

#[test]
fn table_init_stub_does_not_trap() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::Null; 8]];
    vm.set_elem_segment(0, vec![Value::I32(10), Value::I32(11), Value::I32(12)]);
    let mut chunk = Chunk::new("<script>");
    push_i32(&mut chunk, 2); // dst
    push_i32(&mut chunk, 1); // src offset
    push_i32(&mut chunk, 2); // count
    chunk.emit_op_idx_idx(Op::TABLE_INIT, 0u32, 0u32, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("table.init should not trap");

    assert_eq!(vm.wasm_tables[0][2].as_i32(), 11);
    assert_eq!(vm.wasm_tables[0][3].as_i32(), 12);
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
    let mut chunks = vybe_platform_wasm::read_wasm(&wasm).expect("standard wasm should decode");
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
        matches!(&vm.wasm_tables[0][0], Value::Object(obj) if matches!(obj.lock().unwrap().kind, ObjectKind::Function(_)))
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
    let mut chunks = vybe_platform_wasm::read_wasm(&wasm).expect("standard wasm should decode");
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
    let mut chunks = vybe_platform_wasm::read_wasm(&wasm).expect("standard wasm should decode");
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
    vm.wasm_tables = vec![vec![Value::Null; 8]];
    vm.set_elem_segment(0, vec![Value::I32(10)]);
    let mut chunk = Chunk::new("<script>");
    push_i32(&mut chunk, 0); // dst
    push_i32(&mut chunk, 0); // src offset
    push_i32(&mut chunk, 2); // count exceeds segment
    chunk.emit_op_idx_idx(Op::TABLE_INIT, 0u32, 0u32, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = vm
        .run(vec![chunk])
        .expect_err("table.init source OOB must trap");
    assert!(
        err.to_string()
            .contains("out of bounds table access (table.init source)")
    );
}

#[test]
fn table_init_destination_oob_traps() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::Null; 1]];
    vm.set_elem_segment(0, vec![Value::I32(10), Value::I32(11)]);
    let mut chunk = Chunk::new("<script>");
    push_i32(&mut chunk, 0); // dst
    push_i32(&mut chunk, 0); // src offset
    push_i32(&mut chunk, 2); // count exceeds table
    chunk.emit_op_idx_idx(Op::TABLE_INIT, 0u32, 0u32, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = vm
        .run(vec![chunk])
        .expect_err("table.init destination OOB must trap");
    assert!(
        err.to_string()
            .contains("out of bounds table access (table.init destination)")
    );
}

#[test]
fn table_init_after_elem_drop_traps() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::Null; 8]];
    vm.set_elem_segment(0, vec![Value::I32(10)]);
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op_idx(Op::ELEM_DROP, 0u32, 0);
    push_i32(&mut chunk, 0); // dst
    push_i32(&mut chunk, 0); // src offset
    push_i32(&mut chunk, 1); // count
    chunk.emit_op_idx_idx(Op::TABLE_INIT, 0u32, 0u32, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = vm
        .run(vec![chunk])
        .expect_err("table.init after elem.drop must trap");
    assert!(
        err.to_string()
            .contains("out of bounds table access (table.init source)"),
        "unexpected trap: {err}"
    );
}
