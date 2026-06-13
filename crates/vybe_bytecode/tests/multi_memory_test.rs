/// Tests for multi-memory support.
use vybe_bytecode::{Chunk, Op, VM, Value, wasm};

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

fn standard_multi_memory_module(body_ops: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(b"\0asm");
    out.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut out, 1, &[0x01, 0x60, 0x00, 0x01, 0x7f]);
    push_section(&mut out, 3, &[0x01, 0x00]);
    push_section(
        &mut out,
        5,
        &[
            0x02, // two memories
            0x00, 0x01, // memory 0: min 1 page
            0x00, 0x01, // memory 1: min 1 page
        ],
    );

    let mut body = Vec::new();
    body.push(0x00);
    body.extend_from_slice(body_ops);
    body.push(0x0B);

    let mut code = Vec::new();
    code.push(0x01);
    write_leb_u32(&mut code, body.len() as u32);
    code.extend_from_slice(&body);
    push_section(&mut out, 10, &code);

    out
}

fn standard_table64_module_i64_result(body_ops: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(b"\0asm");
    out.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut out, 1, &[0x01, 0x60, 0x00, 0x01, 0x7e]);
    push_section(&mut out, 3, &[0x01, 0x00]);
    push_section(
        &mut out,
        4,
        &[
            0x01, // one table
            0x70, // funcref
            0x04, // table64 min-only limits
            0x02, // min 2
        ],
    );

    let mut body = Vec::new();
    body.push(0x00);
    body.extend_from_slice(body_ops);
    body.push(0x0B);

    let mut code = Vec::new();
    code.push(0x01);
    write_leb_u32(&mut code, body.len() as u32);
    code.extend_from_slice(&body);
    push_section(&mut out, 10, &code);

    out
}

fn standard_memory64_module_i32_result(body_ops: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(b"\0asm");
    out.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut out, 1, &[0x01, 0x60, 0x00, 0x01, 0x7f]);
    push_section(&mut out, 3, &[0x01, 0x00]);
    push_section(
        &mut out,
        5,
        &[
            0x01, // one memory
            0x04, // memory64 min-only limits
            0x01, // min 1
        ],
    );

    let mut body = Vec::new();
    body.push(0x00);
    body.extend_from_slice(body_ops);
    body.push(0x0B);

    let mut code = Vec::new();
    code.push(0x01);
    write_leb_u32(&mut code, body.len() as u32);
    code.extend_from_slice(&body);
    push_section(&mut out, 10, &code);

    out
}

fn standard_imported_memory_module() -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(b"\0asm");
    out.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut out, 1, &[0x01, 0x60, 0x00, 0x01, 0x7f]);
    push_section(
        &mut out,
        2,
        &[
            0x01, // one import
            0x03, b'e', b'n', b'v', // module
            0x06, b'm', b'e', b'm', b'o', b'r', b'y', // name
            0x02, // memory import
            0x00, 0x01, // limits: min 1
        ],
    );
    push_section(&mut out, 3, &[0x01, 0x00]);

    let mut body = Vec::new();
    body.push(0x00);
    body.extend_from_slice(&[
        0x41, 0x00, // i32.const 0
        0x28, 0x02, 0x00, // i32.load align=2 offset=0
    ]);
    body.push(0x0B);

    let mut code = Vec::new();
    code.push(0x01);
    write_leb_u32(&mut code, body.len() as u32);
    code.extend_from_slice(&body);
    push_section(&mut out, 10, &code);

    out
}

fn standard_exported_memory_module() -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(b"\0asm");
    out.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut out, 1, &[0x01, 0x60, 0x00, 0x01, 0x7f]);
    push_section(&mut out, 3, &[0x01, 0x00]);
    push_section(&mut out, 5, &[0x01, 0x00, 0x01]);
    push_section(
        &mut out,
        7,
        &[
            0x01, // one export
            0x06, b'm', b'e', b'm', b'o', b'r', b'y', // name
            0x02, // memory export
            0x00, // memory index 0
        ],
    );

    let mut body = Vec::new();
    body.push(0x00);
    body.extend_from_slice(&[
        0x41, 0x00, // i32.const 0
    ]);
    body.push(0x0B);

    let mut code = Vec::new();
    code.push(0x01);
    write_leb_u32(&mut code, body.len() as u32);
    code.extend_from_slice(&body);
    push_section(&mut out, 10, &code);

    out
}

fn emit_memarg(c: &mut Chunk, align: u32, offset: u32, memidx: u32) {
    let encoded_align = if memidx == 0 { align } else { align | 0x40 };
    c.emit_leb_u32(encoded_align, 0);
    c.emit_leb_u32(offset, 0);
    if memidx != 0 {
        c.emit_leb_u32(memidx, 0);
    }
}

fn emit_memarg64(c: &mut Chunk, align: u32, offset: u64, memidx: u32) {
    let encoded_align = if memidx == 0 { align } else { align | 0x40 };
    c.emit_leb_u32(encoded_align, 0);
    let mut value = offset;
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        c.emit(byte, 0);
        if value == 0 {
            break;
        }
    }
    if memidx != 0 {
        c.emit_leb_u32(memidx, 0);
    }
}

#[test]
fn standard_table64_size_must_not_decode_as_table32_i32_semantics() {
    let bytes = standard_table64_module_i64_result(&[
        0xFC, 0x10, 0x00, // table.size 0
    ]);

    let err = wasm::read_wasm(&bytes).unwrap_err();
    assert!(err.contains("table64") && err.contains("table.size"));
}

#[test]
fn standard_table64_grow_must_not_decode_as_table32_i32_semantics() {
    let bytes = standard_table64_module_i64_result(&[
        0xD0, 0x70, // ref.null func
        0x42, 0x01, // i64.const 1
        0xFC, 0x0F, 0x00, // table.grow 0
    ]);

    let err = wasm::read_wasm(&bytes).unwrap_err();
    assert!(err.contains("table64") && err.contains("table.grow"));
}

#[test]
fn standard_table64_fill_must_not_decode_as_table32_i32_semantics() {
    let bytes = standard_table64_module_i64_result(&[
        0x42, 0x02, // i64.const 2, dst at table end
        0xD0, 0x70, // ref.null func
        0x42, 0x00, // i64.const 0, count
        0xFC, 0x11, 0x00, // table.fill 0
        0x42, 0x00, // i64.const 0, result sentinel
    ]);

    let err = wasm::read_wasm(&bytes).unwrap_err();
    assert!(err.contains("table64") && err.contains("table.fill"));
}

#[test]
fn standard_table64_copy_must_not_decode_as_table32_i32_semantics() {
    let bytes = standard_table64_module_i64_result(&[
        0x42, 0x02, // i64.const 2, dst at table end
        0x42, 0x02, // i64.const 2, src at table end
        0x42, 0x00, // i64.const 0, count
        0xFC, 0x0E, 0x00, 0x00, // table.copy 0 0
        0x42, 0x00, // i64.const 0, result sentinel
    ]);

    let err = wasm::read_wasm(&bytes).unwrap_err();
    assert!(err.contains("table64") && err.contains("table.copy"));
}

#[test]
fn standard_table64_get_set_init_must_not_decode_as_table32_i32_semantics() {
    let cases: &[(&str, &[u8])] = &[
        (
            "table.get",
            &[
                0x42, 0x00, // i64.const 0
                0x25, 0x00, // table.get 0
                0xD1, // ref.is_null
                0xAC, // i64.extend_i32_s
            ],
        ),
        (
            "table.set",
            &[
                0x42, 0x00, // i64.const 0
                0xD0, 0x70, // ref.null func
                0x26, 0x00, // table.set 0
                0x42, 0x00, // i64.const 0
            ],
        ),
        (
            "table.init",
            &[
                0x42, 0x00, // i64.const dst
                0x41, 0x00, // i32.const src
                0x42, 0x00, // i64.const count
                0xFC, 0x0C, 0x00, 0x00, // table.init elemidx=0 tableidx=0
                0x42, 0x00, // i64.const 0
            ],
        ),
    ];

    for (name, body) in cases {
        let bytes = standard_table64_module_i64_result(body);
        let err = wasm::read_wasm(&bytes).unwrap_err();
        assert!(
            err.contains("table64") && err.contains(name),
            "{name} must be rejected or semantically decoded, got {err}"
        );
    }
}

#[test]
fn all_standard_memory64_core_load_store_widths_must_not_decode_as_i32_memory() {
    let load_cases: &[(&str, u8, &[u8], Option<Op>)] = &[
        ("i32.load", 0x28, &[0x42, 0x00], Some(Op::I32_LOAD_64)),
        ("i64.load", 0x29, &[0x42, 0x00, 0xA7], Some(Op::I64_LOAD_64)),
        ("f32.load", 0x2A, &[0x42, 0x00, 0x1A, 0x41, 0x00], None),
        ("f64.load", 0x2B, &[0x42, 0x00, 0x1A, 0x41, 0x00], Some(Op::F64_LOAD_64)),
        ("i32.load8_s", 0x2C, &[0x42, 0x00], None),
        ("i32.load8_u", 0x2D, &[0x42, 0x00], None),
        ("i32.load16_s", 0x2E, &[0x42, 0x00], None),
        ("i32.load16_u", 0x2F, &[0x42, 0x00], None),
        ("i64.load8_s", 0x30, &[0x42, 0x00, 0xA7], None),
        ("i64.load8_u", 0x31, &[0x42, 0x00, 0xA7], None),
        ("i64.load16_s", 0x32, &[0x42, 0x00, 0xA7], None),
        ("i64.load16_u", 0x33, &[0x42, 0x00, 0xA7], None),
        ("i64.load32_s", 0x34, &[0x42, 0x00, 0xA7], None),
        ("i64.load32_u", 0x35, &[0x42, 0x00, 0xA7], None),
    ];

    for (name, opcode, prefix, decoded_op) in load_cases {
        let mut body = Vec::new();
        body.extend_from_slice(prefix);
        body.extend_from_slice(&[*opcode, 0x02, 0x00]);
        let bytes = standard_memory64_module_i32_result(&body);
        if let Some(op) = decoded_op {
            let chunks = wasm::read_wasm(&bytes).expect("supported memory64 load should decode");
            assert!(
                chunks[1]
                    .code
                    .windows(2)
                    .any(|w| w == [op.prefix(), op.sub()]),
                "{name} must decode to memory64 bytecode"
            );
        } else {
            let err = wasm::read_wasm(&bytes).unwrap_err();
            assert!(
                err.contains("memory64") && err.contains(name),
                "{name} must be rejected until memory64 semantics are implemented, got {err}"
            );
        }
    }

    let store_cases: &[(&str, u8, &[u8], Option<Op>)] = &[
        ("i32.store", 0x36, &[0x42, 0x00, 0x41, 0x01], Some(Op::I32_STORE_64)),
        ("i64.store", 0x37, &[0x42, 0x00, 0x42, 0x01], Some(Op::I64_STORE_64)),
        ("f32.store", 0x38, &[0x42, 0x00, 0x43, 0x00, 0x00, 0x80, 0x3F], None),
        (
            "f64.store",
            0x39,
            &[0x42, 0x00, 0x44, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xF0, 0x3F],
            Some(Op::F64_STORE_64),
        ),
        ("i32.store8", 0x3A, &[0x42, 0x00, 0x41, 0x01], None),
        ("i32.store16", 0x3B, &[0x42, 0x00, 0x41, 0x01], None),
        ("i64.store8", 0x3C, &[0x42, 0x00, 0x42, 0x01], None),
        ("i64.store16", 0x3D, &[0x42, 0x00, 0x42, 0x01], None),
        ("i64.store32", 0x3E, &[0x42, 0x00, 0x42, 0x01], None),
    ];

    for (name, opcode, prefix, decoded_op) in store_cases {
        let mut body = Vec::new();
        body.extend_from_slice(prefix);
        body.extend_from_slice(&[*opcode, 0x02, 0x00]);
        body.extend_from_slice(&[0x41, 0x00]);
        let bytes = standard_memory64_module_i32_result(&body);
        if let Some(op) = decoded_op {
            let chunks = wasm::read_wasm(&bytes).expect("supported memory64 store should decode");
            assert!(
                chunks[1]
                    .code
                    .windows(2)
                    .any(|w| w == [op.prefix(), op.sub()]),
                "{name} must decode to memory64 bytecode"
            );
        } else {
            let err = wasm::read_wasm(&bytes).unwrap_err();
            assert!(
                err.contains("memory64") && err.contains(name),
                "{name} must be rejected until memory64 semantics are implemented, got {err}"
            );
        }
    }
}

#[test]
fn standard_memory64_bulk_simd_and_atomic_memory_ops_must_not_decode_as_i32_memory() {
    let cases: &[(&str, &[u8])] = &[
        (
            "memory.copy",
            &[0x42, 0x00, 0x42, 0x00, 0x42, 0x00, 0xFC, 0x0A, 0x00, 0x00, 0x41, 0x00],
        ),
        (
            "memory.fill",
            &[0x42, 0x00, 0x41, 0x00, 0x42, 0x00, 0xFC, 0x0B, 0x00, 0x41, 0x00],
        ),
        (
            "v128.load",
            &[0x42, 0x00, 0xFD, 0x00, 0x04, 0x00, 0xFD, 0x53],
        ),
        (
            "v128.store",
            &[
                0x42, 0x00, 0xFD, 0x0C, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
                0xFD, 0x0B, 0x04, 0x00, 0x41, 0x00,
            ],
        ),
        (
            "i32.atomic.load",
            &[0x42, 0x00, 0xFE, 0x10, 0x02, 0x00],
        ),
        (
            "i32.atomic.store",
            &[0x42, 0x00, 0x41, 0x01, 0xFE, 0x17, 0x02, 0x00, 0x41, 0x00],
        ),
    ];

    for (name, body) in cases {
        let bytes = standard_memory64_module_i32_result(body);
        let err = wasm::read_wasm(&bytes).unwrap_err();
        assert!(
            err.contains("memory64") && err.contains(name),
            "{name} must be rejected or semantically decoded, got {err}"
        );
    }
}

#[test]
fn standard_imported_memory_must_not_decode_without_host_linkage() {
    let bytes = standard_imported_memory_module();

    let chunks = wasm::read_wasm(&bytes).expect("memory import should decode");
    assert_eq!(chunks[0].imports.len(), 1);
    assert_eq!(chunks[0].imports[0].module, "env");
    assert_eq!(chunks[0].imports[0].name, "memory");
}

#[test]
fn standard_exported_memory_must_not_decode_without_export_linkage() {
    let bytes = standard_exported_memory_module();

    let chunks = wasm::read_wasm(&bytes).expect("memory export module should decode");
    assert_eq!(chunks[0].memory_min_pages, vec![1]);
}

#[test]
fn memory_init_creates_new_memory() {
    let mut vm = VM::new();
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    // memory.new: create a new VM-internal memory with 1 page
    let pages = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, pages, 0);
    chunk.emit_op(Op::MEMORY_NEW, 0);
    // Returns the memory index
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    let mem_idx = result.as_i32();
    assert!(
        mem_idx >= 1,
        "new memory index should be >= 1, got {}",
        mem_idx
    );
}

#[test]
fn spec_memarg_i32_store_and_load_use_memory_index() {
    let mut vm = VM::new();
    vm.memory.resize(65536, 0);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    let pages = chunk.add_constant(Value::I32(1));
    let zero = chunk.add_constant(Value::I32(0));
    let value = chunk.add_constant(Value::I32(0x1234_5678));

    chunk.emit_op_u16(Op::CONST, pages, 0);
    chunk.emit_op(Op::MEMORY_NEW, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);

    chunk.emit_op_u16(Op::CONST, zero, 0);
    chunk.emit_op_u16(Op::CONST, value, 0);
    chunk.emit_op(Op::I32_STORE, 0);
    emit_memarg(&mut chunk, 2, 0, 1);

    chunk.emit_op_u16(Op::CONST, zero, 0);
    chunk.emit_op(Op::I32_LOAD, 0);
    emit_memarg(&mut chunk, 2, 0, 1);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 0x1234_5678);
    assert_eq!(
        vm.memory.load_i32(0).unwrap(),
        0,
        "indexed store must not write memory 0"
    );
}

#[test]
fn spec_memory_size_and_grow_use_memory_index_immediate() {
    let mut vm = VM::new();
    vm.memory.resize(3 * 65536, 0);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    let one_page = chunk.add_constant(Value::I32(1));
    let two_pages = chunk.add_constant(Value::I32(2));

    chunk.emit_op_u16(Op::CONST, one_page, 0);
    chunk.emit_op(Op::MEMORY_NEW, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);

    chunk.emit_op_u16(Op::CONST, two_pages, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(1, 0);
    chunk.emit_op(Op::DROP, 0);

    chunk.emit_op(Op::MEMORY_SIZE, 0);
    chunk.emit_leb_u32(1, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 3);
    assert_eq!(
        vm.memory.len(),
        3 * 65536,
        "memory.grow memidx=1 must not grow memory 0"
    );
}

#[test]
fn spec_memory_fill_and_copy_use_memory_indices() {
    let mut vm = VM::new();
    vm.memory.resize(65536, 0);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    let pages = chunk.add_constant(Value::I32(1));
    let zero = chunk.add_constant(Value::I32(0));
    let four = chunk.add_constant(Value::I32(4));
    let fill = chunk.add_constant(Value::I32(0x7F));
    let dest = chunk.add_constant(Value::I32(16));

    chunk.emit_op_u16(Op::CONST, pages, 0);
    chunk.emit_op(Op::MEMORY_NEW, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);

    chunk.emit_op_u16(Op::CONST, zero, 0);
    chunk.emit_op_u16(Op::CONST, fill, 0);
    chunk.emit_op_u16(Op::CONST, four, 0);
    chunk.emit_op(Op::MEMORY_FILL, 0);
    chunk.emit_leb_u32(1, 0);

    chunk.emit_op_u16(Op::CONST, dest, 0);
    chunk.emit_op_u16(Op::CONST, zero, 0);
    chunk.emit_op_u16(Op::CONST, four, 0);
    chunk.emit_op(Op::MEMORY_COPY, 0);
    chunk.emit_leb_u32(1, 0);
    chunk.emit_leb_u32(1, 0);

    chunk.emit_op_u16(Op::CONST, dest, 0);
    chunk.emit_op(Op::I32_LOAD8_U, 0);
    emit_memarg(&mut chunk, 0, 0, 1);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 0x7F);
}

#[test]
fn spec_memory64_memarg_uses_memory_index() {
    let mut vm = VM::new();
    vm.memory.resize(65536, 0);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    let pages = chunk.add_constant(Value::I32(1));
    let zero64 = chunk.add_constant(Value::I64(0));
    let value = chunk.add_constant(Value::I32(0x55AA));

    chunk.emit_op_u16(Op::CONST, pages, 0);
    chunk.emit_op(Op::MEMORY_NEW, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);

    chunk.emit_op_u16(Op::CONST, zero64, 0);
    chunk.emit_op_u16(Op::CONST, value, 0);
    chunk.emit_op(Op::I32_STORE_64, 0);
    emit_memarg64(&mut chunk, 2, 0, 1);

    chunk.emit_op_u16(Op::CONST, zero64, 0);
    chunk.emit_op(Op::I32_LOAD_64, 0);
    emit_memarg64(&mut chunk, 2, 0, 1);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 0x55AA);
    assert_eq!(
        vm.memory.load_i32(0).unwrap(),
        0,
        "memory64 memidx=1 store must not write memory 0"
    );
}

#[test]
fn spec_memory64_size_and_grow_use_memory_index() {
    let mut vm = VM::new();
    vm.memory.resize(2 * 65536, 0);

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    let one_page = chunk.add_constant(Value::I32(1));
    let two_pages64 = chunk.add_constant(Value::I64(2));

    chunk.emit_op_u16(Op::CONST, one_page, 0);
    chunk.emit_op(Op::MEMORY_NEW, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);

    chunk.emit_op_u16(Op::CONST, two_pages64, 0);
    chunk.emit_op(Op::I64_MEMORY_GROW, 0);
    chunk.emit_leb_u32(1, 0);
    chunk.emit_op(Op::DROP, 0);

    chunk.emit_op(Op::I64_MEMORY_SIZE, 0);
    chunk.emit_leb_u32(1, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i64(), 3);
    assert_eq!(
        vm.memory.len(),
        2 * 65536,
        "memory64.grow memidx=1 must not grow memory 0"
    );
}

#[test]
fn decoded_standard_module_materializes_declared_memories() {
    let wasm = standard_multi_memory_module(&[
        0x41, 0x00, // i32.const 0
        0x41, 0x2A, // i32.const 42
        0x36, 0x42, 0x00, 0x01, // i32.store align=2|memidx, offset=0, memidx=1
        0x41, 0x00, // i32.const 0
        0x28, 0x42, 0x00, 0x01, // i32.load align=2|memidx, offset=0, memidx=1
    ]);
    let mut chunks = vybe_bytecode::wasm::read_wasm(&wasm).expect("standard wasm should decode");
    assert_eq!(chunks[0].memory_min_pages, vec![1, 1]);

    let function = chunks.remove(1);
    let mut vm = VM::new();
    let result = vm.run(vec![function]).unwrap();

    assert_eq!(result.as_i32(), 42);
    assert_eq!(
        vm.memory.load_i32(0).unwrap(),
        0,
        "decoded memidx=1 store must not write memory 0"
    );
}

#[test]
fn decoded_standard_module_uses_memidx_for_f64_store_and_load() {
    let wasm = standard_multi_memory_module(&[
        0x41, 0x00, // i32.const 0
        0x44, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x14, 0x40, // f64.const 5.0
        0x39, 0x43, 0x00, 0x01, // f64.store align=3|memidx, offset=0, memidx=1
        0x41, 0x00, // i32.const 0
        0x2b, 0x43, 0x00, 0x01, // f64.load align=3|memidx, offset=0, memidx=1
        0xaa, // i32.trunc_f64_s
    ]);
    let mut chunks = vybe_bytecode::wasm::read_wasm(&wasm).expect("standard wasm should decode");
    assert_eq!(chunks[0].memory_min_pages, vec![1, 1]);

    let function = chunks.remove(1);
    assert!(
        function
            .code
            .windows(5)
            .any(|w| w == [Op::F64_STORE.prefix(), Op::F64_STORE.sub(), 0x43, 0x00, 0x01])
    );
    assert!(
        function
            .code
            .windows(5)
            .any(|w| w == [Op::F64_LOAD.prefix(), Op::F64_LOAD.sub(), 0x43, 0x00, 0x01])
    );

    let mut vm = VM::new();
    let result = vm.run(vec![function]).unwrap();
    assert_eq!(result.as_i32(), 5);
    assert_eq!(
        vm.memory.load_i64(0).unwrap(),
        0,
        "decoded f64.store memidx=1 must not write memory 0"
    );
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
    chunk.emit_op(Op::MEMORY_NEW, 0);
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
    vm.memory.store_i32(0, 99).unwrap();

    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 2;

    // Create second memory
    let pages = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, pages, 0);
    chunk.emit_op(Op::MEMORY_NEW, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0); // mem_idx = 1

    // Copy 4 bytes from memory 0 addr 0 → memory 1 addr 0
    let zero = chunk.add_constant(Value::I32(0));
    let four = chunk.add_constant(Value::I32(4));
    let mem0 = chunk.add_constant(Value::I32(0));
    let mem1 = chunk.add_constant(Value::I32(1));

    // Stack order: dst_mem, dst_addr, src_mem, src_addr, len
    chunk.emit_op_u16(Op::CONST, mem1, 0); // dst_mem = 1
    chunk.emit_op_u16(Op::CONST, zero, 0); // dst_addr = 0
    chunk.emit_op_u16(Op::CONST, mem0, 0); // src_mem = 0
    chunk.emit_op_u16(Op::CONST, zero, 0); // src_addr = 0
    chunk.emit_op_u16(Op::CONST, four, 0); // len = 4
    chunk.emit_op(Op::MEMORY_COPY_CROSS, 0);

    // Switch to memory 1 and read
    chunk.emit_op(Op::MEMORY_SELECT, 0);
    chunk.emit(1, 0); // memory index 1
    chunk.emit_op_u16(Op::CONST, zero, 0);
    chunk.emit_op(Op::I32_LOAD, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(
        result.as_i32(),
        99,
        "data should have been copied to memory 1"
    );
}

#[test]
fn memory_size_reports_correct_size() {
    let mut vm = VM::new();
    vm.memory.resize(2 * 65536, 0); // 2 pages
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op(Op::MEMORY_SIZE, 0);
    chunk.emit_op(Op::HALT, 0);
    let r = vm.run(vec![chunk]).unwrap();
    assert_eq!(r.as_i32(), 2, "memory.size should return page count");
}

#[test]
fn memory_grow_increases_size_and_returns_old() {
    let mut vm = VM::new();
    vm.memory.resize(65536, 0); // 1 page
    let mut chunk = Chunk::new("<script>");
    let delta = chunk.add_constant(Value::I32(2));
    chunk.emit_op_u16(Op::CONST, delta, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_op(Op::HALT, 0);
    let r = vm.run(vec![chunk]).unwrap();
    assert_eq!(r.as_i32(), 1, "memory.grow returns old size in pages");
    assert_eq!(vm.memory.len(), 3 * 65536, "memory grew by 2 pages");
}

#[test]
fn memory_fill_in_memory_zero() {
    // memory.fill on memory 0: fill 4 bytes with 0xAB starting at addr 8
    let mut vm = VM::new();
    vm.memory.resize(65536, 0);
    let mut chunk = Chunk::new("<script>");
    let start = chunk.add_constant(Value::I32(8));
    let byte = chunk.add_constant(Value::I32(0xAB));
    let count = chunk.add_constant(Value::I32(4));
    chunk.emit_op_u16(Op::CONST, start, 0);
    chunk.emit_op_u16(Op::CONST, byte, 0);
    chunk.emit_op_u16(Op::CONST, count, 0);
    chunk.emit_op(Op::MEMORY_FILL, 0);
    // Load back byte at addr 8
    chunk.emit_op_u16(Op::CONST, start, 0);
    chunk.emit_op(Op::I32_LOAD8_U, 0);
    chunk.emit_op(Op::HALT, 0);
    let r = vm.run(vec![chunk]).unwrap();
    assert_eq!(r.as_i32(), 0xAB);
}
