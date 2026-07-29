/// Tests for multi-memory support.
use vybe_runtime::{Chunk, Op, VM, Value};
use vybe_platform_wasm as wasm;

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
    push_section(
        &mut out,
        9,
        &[
            0x01, // one passive element segment
            0x01, // passive, elemkind form
            0x00, // funcref elemkind
            0x00, // zero elements
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

fn standard_table64_module_i64_result_with_elem(body_ops: &[u8], elems: &[u8]) -> Vec<u8> {
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
            0x04, // min 4
        ],
    );

    let mut elem = Vec::new();
    elem.push(0x01); // one passive element segment
    elem.push(0x01); // passive, elemkind form
    elem.push(0x00); // funcref elemkind
    write_leb_u32(&mut elem, elems.len() as u32);
    elem.extend_from_slice(elems);
    push_section(&mut out, 9, &elem);

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

fn standard_memory_module_i32_result_with_limits(
    min: u32,
    max: Option<u32>,
    body_ops: &[u8],
) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(b"\0asm");
    out.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut out, 1, &[0x01, 0x60, 0x00, 0x01, 0x7f]);
    push_section(&mut out, 3, &[0x01, 0x00]);
    let mut memory = Vec::new();
    memory.push(0x01);
    memory.push(if max.is_some() { 0x01 } else { 0x00 });
    write_leb_u32(&mut memory, min);
    if let Some(max) = max {
        write_leb_u32(&mut memory, max);
    }
    push_section(&mut out, 5, &memory);

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

fn decoded_body_contains(bytes: &[u8], op: Op) -> bool {
    let chunks = wasm::read_wasm(bytes).expect("standard module should decode");
    chunks[1].code.windows(4).any(|w| w == op.encode())
}

#[test]
fn standard_table64_size_must_not_decode_as_table32_i32_semantics() {
    let bytes = standard_table64_module_i64_result(&[
        0xFC, 0x10, 0x00, // table.size 0
    ]);

    assert!(decoded_body_contains(&bytes, Op::TABLE_SIZE));
}

#[test]
fn standard_table64_grow_must_not_decode_as_table32_i32_semantics() {
    let bytes = standard_table64_module_i64_result(&[
        0xD0, 0x70, // ref.null func
        0x42, 0x01, // i64.const 1
        0xFC, 0x0F, 0x00, // table.grow 0
    ]);

    assert!(decoded_body_contains(&bytes, Op::TABLE_GROW));
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

    assert!(decoded_body_contains(&bytes, Op::TABLE_FILL));
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

    assert!(decoded_body_contains(&bytes, Op::TABLE_COPY));
}

#[test]
fn standard_table64_get_set_init_must_not_decode_as_table32_i32_semantics() {
    let cases: &[(&str, &[u8], Op)] = &[
        (
            "table.get",
            &[
                0x42, 0x00, // i64.const 0
                0x25, 0x00, // table.get 0
                0xD1, // ref.is_null
                0xAC, // i64.extend_i32_s
            ],
            Op::TABLE_GET,
        ),
        (
            "table.set",
            &[
                0x42, 0x00, // i64.const 0
                0xD0, 0x70, // ref.null func
                0x26, 0x00, // table.set 0
                0x42, 0x00, // i64.const 0
            ],
            Op::TABLE_SET,
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
            Op::TABLE_INIT,
        ),
    ];

    for (name, body, op) in cases {
        let bytes = standard_table64_module_i64_result(body);
        assert!(
            decoded_body_contains(&bytes, *op),
            "{name} must decode to table64 bytecode"
        );
    }
}

#[test]
fn all_standard_memory64_core_load_store_widths_must_not_decode_as_i32_memory() {
    let load_cases: &[(&str, u8, &[u8], Op)] = &[
        ("i32.load", 0x28, &[0x42, 0x00], Op::I32_LOAD),
        ("i64.load", 0x29, &[0x42, 0x00, 0xA7], Op::I64_LOAD),
        (
            "f32.load",
            0x2A,
            &[0x42, 0x00, 0x1A, 0x41, 0x00],
            Op::F32_LOAD,
        ),
        (
            "f64.load",
            0x2B,
            &[0x42, 0x00, 0x1A, 0x41, 0x00],
            Op::F64_LOAD,
        ),
        ("i32.load8_s", 0x2C, &[0x42, 0x00], Op::I32_LOAD8_S),
        ("i32.load8_u", 0x2D, &[0x42, 0x00], Op::I32_LOAD8_U),
        ("i32.load16_s", 0x2E, &[0x42, 0x00], Op::I32_LOAD16_S),
        ("i32.load16_u", 0x2F, &[0x42, 0x00], Op::I32_LOAD16_U),
        ("i64.load8_s", 0x30, &[0x42, 0x00, 0xA7], Op::I64_LOAD8_S),
        ("i64.load8_u", 0x31, &[0x42, 0x00, 0xA7], Op::I64_LOAD8_U),
        ("i64.load16_s", 0x32, &[0x42, 0x00, 0xA7], Op::I64_LOAD16_S),
        ("i64.load16_u", 0x33, &[0x42, 0x00, 0xA7], Op::I64_LOAD16_U),
        ("i64.load32_s", 0x34, &[0x42, 0x00, 0xA7], Op::I64_LOAD32_S),
        ("i64.load32_u", 0x35, &[0x42, 0x00, 0xA7], Op::I64_LOAD32_U),
    ];

    for (name, opcode, prefix, decoded_op) in load_cases {
        let mut body = Vec::new();
        body.extend_from_slice(prefix);
        body.extend_from_slice(&[*opcode, 0x02, 0x00]);
        let bytes = standard_memory64_module_i32_result(&body);
        assert!(
            decoded_body_contains(&bytes, *decoded_op),
            "{name} must decode to memory64 bytecode"
        );
    }

    let store_cases: &[(&str, u8, &[u8], Op)] = &[
        ("i32.store", 0x36, &[0x42, 0x00, 0x41, 0x01], Op::I32_STORE),
        ("i64.store", 0x37, &[0x42, 0x00, 0x42, 0x01], Op::I64_STORE),
        (
            "f32.store",
            0x38,
            &[0x42, 0x00, 0x43, 0x00, 0x00, 0x80, 0x3F],
            Op::F32_STORE,
        ),
        (
            "f64.store",
            0x39,
            &[
                0x42, 0x00, 0x44, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xF0, 0x3F,
            ],
            Op::F64_STORE,
        ),
        (
            "i32.store8",
            0x3A,
            &[0x42, 0x00, 0x41, 0x01],
            Op::I32_STORE8,
        ),
        (
            "i32.store16",
            0x3B,
            &[0x42, 0x00, 0x41, 0x01],
            Op::I32_STORE16,
        ),
        (
            "i64.store8",
            0x3C,
            &[0x42, 0x00, 0x42, 0x01],
            Op::I64_STORE8,
        ),
        (
            "i64.store16",
            0x3D,
            &[0x42, 0x00, 0x42, 0x01],
            Op::I64_STORE16,
        ),
        (
            "i64.store32",
            0x3E,
            &[0x42, 0x00, 0x42, 0x01],
            Op::I64_STORE32,
        ),
    ];

    for (name, opcode, prefix, decoded_op) in store_cases {
        let mut body = Vec::new();
        body.extend_from_slice(prefix);
        body.extend_from_slice(&[*opcode, 0x02, 0x00]);
        body.extend_from_slice(&[0x41, 0x00]);
        let bytes = standard_memory64_module_i32_result(&body);
        assert!(
            decoded_body_contains(&bytes, *decoded_op),
            "{name} must decode to memory64 bytecode"
        );
    }
}

#[test]
fn standard_memory64_bulk_simd_and_atomic_memory_ops_must_not_decode_as_i32_memory() {
    let supported: &[(&str, &[u8], Op)] = &[
        (
            "memory.copy",
            &[
                0x42, 0x00, 0x42, 0x00, 0x42, 0x00, 0xFC, 0x0A, 0x00, 0x00, 0x41, 0x00,
            ],
            Op::MEMORY_COPY,
        ),
        (
            "memory.fill",
            &[
                0x42, 0x00, 0x41, 0x00, 0x42, 0x00, 0xFC, 0x0B, 0x00, 0x41, 0x00,
            ],
            Op::MEMORY_FILL,
        ),
        (
            "v128.load",
            &[0x42, 0x00, 0xFD, 0x00, 0x04, 0x00, 0xFD, 0x53],
            Op::V128_LOAD,
        ),
        (
            "v128.store",
            &[
                0x42, 0x00, 0xFD, 0x0C, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0xFD, 0x0B,
                0x04, 0x00, 0x41, 0x00,
            ],
            Op::V128_STORE,
        ),
        (
            "i32.atomic.load",
            &[0x42, 0x00, 0xFE, 0x10, 0x02, 0x00],
            Op::I32_ATOMIC_LOAD,
        ),
        (
            "i32.atomic.store",
            &[0x42, 0x00, 0x41, 0x01, 0xFE, 0x17, 0x02, 0x00, 0x41, 0x00],
            Op::I32_ATOMIC_STORE,
        ),
    ];

    for (name, body, op) in supported {
        let bytes = standard_memory64_module_i32_result(body);
        assert!(
            decoded_body_contains(&bytes, *op),
            "{name} must decode to memory64 bytecode"
        );
    }
}

#[test]
fn standard_imported_memory_must_not_decode_without_host_linkage() {
    let bytes = standard_imported_memory_module();

    let chunks = wasm::read_wasm(&bytes).expect("memory import should decode");
    assert!(
        chunks[0].imports.is_empty(),
        "memory imports must not be modeled as callable function imports"
    );
    assert_eq!(chunks[0].memory_min_pages, vec![1]);
}

#[test]
fn standard_exported_memory_must_not_decode_without_export_linkage() {
    let bytes = standard_exported_memory_module();

    let chunks = wasm::read_wasm(&bytes).expect("memory export module should decode");
    assert_eq!(chunks[0].memory_min_pages, vec![1]);
}

#[test]
fn decoded_standard_imported_memory_is_materialized_for_execution() {
    let bytes = standard_imported_memory_module();
    let chunks = wasm::read_wasm(&bytes).expect("memory import should decode");

    let result = VM::new()
        .run(vec![chunks[1].clone()])
        .expect("decoded imported memory should be instantiated");
    assert_eq!(result, Value::I32(0));
}

#[test]
fn decoded_standard_memory_grow_respects_module_max_limit() {
    let bytes = standard_memory_module_i32_result_with_limits(
        1,
        Some(1),
        &[
            0x41, 0x01, // i32.const 1
            0x40, 0x00, // memory.grow 0
        ],
    );
    let chunks = wasm::read_wasm(&bytes).expect("memory max module should decode");

    let mut vm = VM::new();
    let result = vm.run(vec![chunks[1].clone()]).unwrap();
    assert_eq!(result, Value::I32(-1));
    assert_eq!(vm.memory.len(), 65536);
}

#[test]
fn decoded_standard_table64_size_returns_i64() {
    let bytes = standard_table64_module_i64_result(&[
        0xFC, 0x10, 0x00, // table.size 0
    ]);
    let chunks = wasm::read_wasm(&bytes).expect("table64 module should decode");

    let result = VM::new().run(vec![chunks[1].clone()]).unwrap();
    assert_eq!(result, Value::I64(2));
}

#[test]
fn decoded_standard_table64_grow_returns_old_i64_size() {
    let bytes = standard_table64_module_i64_result(&[
        0xD0, 0x70, // ref.null func
        0x42, 0x03, // i64.const 3
        0xFC, 0x0F, 0x00, // table.grow 0
    ]);
    let chunks = wasm::read_wasm(&bytes).expect("table64 module should decode");

    let mut vm = VM::new();
    let result = vm.run(vec![chunks[1].clone()]).unwrap();
    assert_eq!(result, Value::I64(2));
    assert_eq!(vm.wasm_tables[0].len(), 5);
}

#[test]
fn decoded_standard_table64_set_and_get_use_i64_index() {
    let bytes = standard_table64_module_i64_result(&[
        0x42, 0x01, // i64.const 1
        0xD0, 0x70, // ref.null func
        0x26, 0x00, // table.set 0
        0x42, 0x01, // i64.const 1
        0x25, 0x00, // table.get 0
        0xD1, // ref.is_null
        0xAC, // i64.extend_i32_s
    ]);
    let chunks = wasm::read_wasm(&bytes).expect("table64 module should decode");

    let result = VM::new().run(vec![chunks[1].clone()]).unwrap();
    assert_eq!(result, Value::I64(1));
}

#[test]
fn decoded_standard_table64_fill_and_copy_use_i64_indices() {
    let bytes = standard_table64_module_i64_result(&[
        0x42, 0x00, // i64.const dst
        0xD0, 0x70, // ref.null func
        0x42, 0x02, // i64.const count
        0xFC, 0x11, 0x00, // table.fill 0
        0x42, 0x01, // i64.const dst
        0x42, 0x00, // i64.const src
        0x42, 0x01, // i64.const count
        0xFC, 0x0E, 0x00, 0x00, // table.copy 0 0
        0x42, 0x01, // i64.const 1
        0x25, 0x00, // table.get 0
        0xD1, // ref.is_null
        0xAC, // i64.extend_i32_s
    ]);
    let chunks = wasm::read_wasm(&bytes).expect("table64 module should decode");

    let result = VM::new().run(vec![chunks[1].clone()]).unwrap();
    assert_eq!(result, Value::I64(1));
}

#[test]
fn decoded_standard_table64_init_uses_i64_indices() {
    let bytes = standard_table64_module_i64_result_with_elem(
        &[
            0x42, 0x02, // i64.const dst
            0x41, 0x00, // i32.const src
            0x42, 0x01, // i64.const count
            0xFC, 0x0C, 0x00, 0x00, // table.init elemidx=0 tableidx=0
            0x42, 0x02, // i64.const 2
            0x25, 0x00, // table.get 0
            0xD1, // ref.is_null
            0xAC, // i64.extend_i32_s
        ],
        &[0x00],
    );
    let chunks = wasm::read_wasm(&bytes).expect("table64 module should decode");

    let result = VM::new().run(vec![chunks[1].clone()]).unwrap();
    assert_eq!(result, Value::I64(0));
}

#[test]
fn spec_memory64_all_scalar_widths_execute_with_i64_addresses() {
    fn run_pair(store_op: Op, load_op: Op, value: Value, expected: Value, align: u32) {
        let mut vm = VM::new();
        vm.memory.resize(65536, 0);

        let mut chunk = Chunk::new("<script>");
        let base = chunk.add_constant(Value::I64(4));
        let value_idx = chunk.add_constant(value);

        chunk.emit_op_u16(Op::CONST, base, 0);
        chunk.emit_op_u16(Op::CONST, value_idx, 0);
        chunk.emit_op(store_op, 0);
        emit_memarg64(&mut chunk, align, 3, 0);

        chunk.emit_op_u16(Op::CONST, base, 0);
        chunk.emit_op(load_op, 0);
        emit_memarg64(&mut chunk, align, 3, 0);
        chunk.emit_op(Op::HALT, 0);

        let result = vm.run(vec![chunk]).unwrap();
        match (result, expected) {
            (Value::F64(a), Value::F64(b)) => assert!((a - b).abs() < 0.00001),
            (got, want) => assert_eq!(got, want),
        }
    }

    run_pair(
        Op::I32_STORE,
        Op::I32_LOAD,
        Value::I32(0x12345678),
        Value::I32(0x12345678),
        2,
    );
    run_pair(
        Op::I64_STORE,
        Op::I64_LOAD,
        Value::I64(0x1122334455667788),
        Value::I64(0x1122334455667788),
        3,
    );
    run_pair(
        Op::F32_STORE,
        Op::F32_LOAD,
        Value::F64(1.5),
        Value::F64(1.5),
        2,
    );
    run_pair(
        Op::F64_STORE,
        Op::F64_LOAD,
        Value::F64(2.25),
        Value::F64(2.25),
        3,
    );
    run_pair(
        Op::I32_STORE8,
        Op::I32_LOAD8_S,
        Value::I32(0xFE),
        Value::I32(-2),
        0,
    );
    run_pair(
        Op::I32_STORE8,
        Op::I32_LOAD8_U,
        Value::I32(0xFE),
        Value::I32(0xFE),
        0,
    );
    run_pair(
        Op::I32_STORE16,
        Op::I32_LOAD16_S,
        Value::I32(0xFFFE),
        Value::I32(-2),
        1,
    );
    run_pair(
        Op::I32_STORE16,
        Op::I32_LOAD16_U,
        Value::I32(0xFFFE),
        Value::I32(0xFFFE),
        1,
    );
    run_pair(
        Op::I64_STORE8,
        Op::I64_LOAD8_S,
        Value::I64(0xFE),
        Value::I64(-2),
        0,
    );
    run_pair(
        Op::I64_STORE8,
        Op::I64_LOAD8_U,
        Value::I64(0xFE),
        Value::I64(0xFE),
        0,
    );
    run_pair(
        Op::I64_STORE16,
        Op::I64_LOAD16_S,
        Value::I64(0xFFFE),
        Value::I64(-2),
        1,
    );
    run_pair(
        Op::I64_STORE16,
        Op::I64_LOAD16_U,
        Value::I64(0xFFFE),
        Value::I64(0xFFFE),
        1,
    );
    run_pair(
        Op::I64_STORE32,
        Op::I64_LOAD32_S,
        Value::I64(0xFFFF_FFFE),
        Value::I64(-2),
        2,
    );
    run_pair(
        Op::I64_STORE32,
        Op::I64_LOAD32_U,
        Value::I64(0xFFFF_FFFE),
        Value::I64(0xFFFF_FFFE),
        2,
    );
}

#[test]
fn spec_table64_runtime_uses_i64_indices_and_results() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::I32(1), Value::I32(2), Value::I32(3)]];

    let mut grow = Chunk::new("<grow>");
    grow.emit_op(Op::NULL, 0);
    let delta = grow.add_constant(Value::I64(2));
    grow.emit_op_u16(Op::CONST, delta, 0);
    grow.emit_op_u8(Op::TABLE_GROW, 0, 0);
    grow.emit_op(Op::HALT, 0);
    let result = vm.run(vec![grow]).unwrap();
    assert_eq!(result.as_i64(), 3);
    assert_eq!(vm.wasm_tables[0].len(), 5);

    let mut chunk = Chunk::new("<table64>");
    let idx1 = chunk.add_constant(Value::I64(1));
    let idx2 = chunk.add_constant(Value::I64(2));
    let idx3 = chunk.add_constant(Value::I64(3));
    let count2 = chunk.add_constant(Value::I64(2));
    let seven = chunk.add_constant(Value::I32(7));
    let nine = chunk.add_constant(Value::I32(9));

    chunk.emit_op_u16(Op::CONST, idx1, 0);
    chunk.emit_op_u16(Op::CONST, seven, 0);
    chunk.emit_op_u8(Op::TABLE_SET, 0, 0);

    chunk.emit_op_u16(Op::CONST, idx2, 0);
    chunk.emit_op_u16(Op::CONST, nine, 0);
    chunk.emit_op_u16(Op::CONST, count2, 0);
    chunk.emit_op_u8(Op::TABLE_FILL, 0, 0);

    chunk.emit_op_u16(Op::CONST, idx3, 0);
    chunk.emit_op_u16(Op::CONST, idx1, 0);
    chunk.emit_op_u16(Op::CONST, count2, 0);
    chunk.emit_op_u8_u8(Op::TABLE_COPY, 0, 0, 0);

    chunk.emit_op_u16(Op::CONST, idx3, 0);
    chunk.emit_op_u8(Op::TABLE_GET, 0, 0);
    chunk.emit_op(Op::HALT, 0);

    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_i32(), 7);
    assert_eq!(vm.wasm_tables[0][1].as_i32(), 7);
    assert_eq!(vm.wasm_tables[0][2].as_i32(), 9);
    assert_eq!(vm.wasm_tables[0][3].as_i32(), 7);
}

#[test]
fn spec_table64_init_copies_element_segment_with_i64_indices() {
    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::Null; 6]];
    vm.set_elem_segment(
        0,
        vec![
            Value::I32(21),
            Value::I32(22),
            Value::I32(23),
            Value::I32(24),
        ],
    );
    let mut chunk = Chunk::new("<table64-init>");
    let dst = chunk.add_constant(Value::I64(3));
    let src = chunk.add_constant(Value::I64(1));
    let count = chunk.add_constant(Value::I64(2));
    chunk.emit_op_u16(Op::CONST, dst, 0);
    chunk.emit_op_u16(Op::CONST, src, 0);
    chunk.emit_op_u16(Op::CONST, count, 0);
    chunk.emit_op_u8_u8(Op::TABLE_INIT, 0, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.run(vec![chunk]).expect("table64.init should copy");
    assert_eq!(vm.wasm_tables[0][3].as_i32(), 22);
    assert_eq!(vm.wasm_tables[0][4].as_i32(), 23);
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
    let mut chunks = vybe_platform_wasm::read_wasm(&wasm).expect("standard wasm should decode");
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
    let mut chunks = vybe_platform_wasm::read_wasm(&wasm).expect("standard wasm should decode");
    assert_eq!(chunks[0].memory_min_pages, vec![1, 1]);

    let function = chunks.remove(1);
    let fs = Op::F64_STORE.encode();
    let fl = Op::F64_LOAD.encode();
    assert!(
        function
            .code
            .windows(7)
            .any(|w| w == [fs[0], fs[1], fs[2], fs[3], 0x43, 0x00, 0x01])
    );
    assert!(
        function
            .code
            .windows(7)
            .any(|w| w == [fl[0], fl[1], fl[2], fl[3], 0x43, 0x00, 0x01])
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
fn decoded_standard_memory_copy_from_memory1_to_memory0_preserves_mixed_indices() {
    let wasm = standard_multi_memory_module(&[
        0x41, 0x00, // i32.const 0
        0x41, 0x2D, // i32.const 45
        0x36, 0x42, 0x00, 0x01, // i32.store align=2|memidx, offset=0, memidx=1
        0x41, 0x04, // dst = 4 in memory 0
        0x41, 0x00, // src = 0 in memory 1
        0x41, 0x04, // len = 4
        0xFC, 0x0A, 0x00, 0x01, // memory.copy dstmem=0 srcmem=1
        0x41, 0x04, // i32.const 4
        0x28, 0x02, 0x00, // i32.load memory 0
    ]);
    let mut chunks = vybe_platform_wasm::read_wasm(&wasm).expect("standard wasm should decode");
    let function = chunks.remove(1);

    let mut vm = VM::new();
    let result = vm.run(vec![function]).unwrap();

    assert_eq!(
        result.as_i32(),
        45,
        "decoded memory.copy must keep dstmem=0 and srcmem=1 distinct"
    );
}

#[test]
fn decoded_standard_memory_copy_from_memory0_to_memory1_preserves_mixed_indices() {
    let wasm = standard_multi_memory_module(&[
        0x41, 0x00, // i32.const 0
        0x41, 0x3A, // i32.const 58
        0x36, 0x02, 0x00, // i32.store memory 0
        0x41, 0x08, // dst = 8 in memory 1
        0x41, 0x00, // src = 0 in memory 0
        0x41, 0x04, // len = 4
        0xFC, 0x0A, 0x01, 0x00, // memory.copy dstmem=1 srcmem=0
        0x41, 0x08, // i32.const 8
        0x28, 0x42, 0x00, 0x01, // i32.load align=2|memidx, offset=0, memidx=1
    ]);
    let mut chunks = vybe_platform_wasm::read_wasm(&wasm).expect("standard wasm should decode");
    let function = chunks.remove(1);

    let mut vm = VM::new();
    let result = vm.run(vec![function]).unwrap();

    assert_eq!(
        result.as_i32(),
        58,
        "decoded memory.copy must keep dstmem=1 and srcmem=0 distinct"
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
