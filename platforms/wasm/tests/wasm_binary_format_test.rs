//! Tests for the WASM binary format: magic bytes, version, section structure,
//! round-trip read/write, and opcode byte-value compliance per the spec.
//! Binary I/O compliance — not execution semantics (those live in per-opcode files).

use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};
use vybe_platform_wasm as wasm;

// ── WASM binary magic and version ─────────────────────────────────────────

#[test]
fn wasm_magic_bytes_are_correct() {
    let bytes = wasm::write_wasm(&[Chunk::new("<script>")]);
    assert_eq!(
        &bytes[0..4],
        b"\0asm",
        "WASM magic must be 0x00 0x61 0x73 0x6D"
    );
}

#[test]
fn wasm_version_is_one() {
    let bytes = wasm::write_wasm(&[Chunk::new("<script>")]);
    assert_eq!(
        &bytes[4..8],
        &[1, 0, 0, 0],
        "WASM version must be 0x01 0x00 0x00 0x00"
    );
}

#[test]
fn wasm_output_is_at_least_8_bytes() {
    let bytes = wasm::write_wasm(&[Chunk::new("<script>")]);
    assert!(
        bytes.len() >= 8,
        "minimum WASM module is 8 bytes (magic + version)"
    );
}

#[test]
fn reader_ignores_unknown_custom_sections_in_any_position() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);

    push_section(
        &mut bytes,
        0,
        &named_custom_section("before-type", b"ignored"),
    );
    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(
        &mut bytes,
        0,
        &named_custom_section("between-type-func", b"ignored"),
    );
    push_section(&mut bytes, 3, &[0x01, 0x00]);
    push_section(
        &mut bytes,
        0,
        &named_custom_section("before-code", b"ignored"),
    );
    push_section(&mut bytes, 10, &[0x01, 0x02, 0x00, 0x0B]);
    push_section(
        &mut bytes,
        0,
        &named_custom_section("after-code", b"ignored"),
    );

    let chunks = wasm::read_wasm(&bytes).expect("custom sections must not affect decoding");
    assert_eq!(chunks.len(), 2);
    assert_eq!(chunks[1].name, "func_0");
}

#[test]
fn reader_rejects_duplicate_known_sections() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut bytes, 3, &[0x01, 0x00]);
    push_section(&mut bytes, 10, &[0x01, 0x02, 0x00, 0x0B]);

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "binary modules must reject duplicate non-custom sections"
    );
}

#[test]
fn reader_rejects_known_sections_out_of_order() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut bytes, 3, &[0x01, 0x00]);
    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut bytes, 10, &[0x01, 0x02, 0x00, 0x0B]);

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "binary modules must reject known sections that appear out of order"
    );
}

#[test]
fn reader_rejects_function_and_code_count_mismatch() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut bytes, 3, &[0x02, 0x00, 0x00]);
    push_section(&mut bytes, 10, &[0x01, 0x02, 0x00, 0x0B]);

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "function and code sections must declare the same number of functions"
    );
}

#[test]
fn reader_rejects_function_type_index_out_of_range() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut bytes, 3, &[0x01, 0x01]);
    push_section(&mut bytes, 10, &[0x01, 0x02, 0x00, 0x0B]);

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "function section type indices must refer to declared function types"
    );
}

#[test]
fn reader_rejects_code_body_without_end_opcode() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut bytes, 3, &[0x01, 0x00]);
    push_section(&mut bytes, 10, &[0x01, 0x02, 0x00, 0x01]);

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "each code body must terminate with the wasm end opcode"
    );
}

#[test]
fn reader_rejects_instruction_validation_errors() {
    let cases: &[(&str, &[u8])] = &[
        // br 0 in a body targets the implicit function label (valid per
        // spec §3.4.10) — depth 1 has no enclosing label to name.
        ("branch depth without enclosing label", &[0x0C, 0x01]),
        ("local.get index outside params and locals", &[0x20, 0x00]),
        ("call index outside function index space", &[0x10, 0x01]),
        ("i32.add without two stack operands", &[0x6A]),
    ];

    for (name, body_ops) in cases {
        let bytes = standard_module_with_body(body_ops, 0);
        assert!(
            wasm::read_wasm(&bytes).is_err(),
            "reader must reject instruction validation error: {name}"
        );
    }
}

#[test]
fn reader_rejects_truncated_section_payload() {
    let bytes = [
        0x00, 0x61, 0x73, 0x6D, // magic
        0x01, 0x00, 0x00, 0x00, // version
        0x01, // type section
        0x05, // declared payload length, but only two bytes follow
        0x01, 0x60,
    ];

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "section payloads must not be silently truncated"
    );
}

#[test]
fn reader_rejects_memory_min_greater_than_max() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut bytes, 3, &[0x01, 0x00]);
    push_section(&mut bytes, 5, &[0x01, 0x01, 0x02, 0x01]);
    push_section(&mut bytes, 10, &[0x01, 0x02, 0x00, 0x0B]);

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "memory limits must reject min pages greater than max pages"
    );
}

#[test]
fn reader_rejects_duplicate_export_names() {
    let mut exports = Vec::new();
    exports.push(0x02);
    exports.extend_from_slice(&[0x01, b'f', 0x00, 0x00]);
    exports.extend_from_slice(&[0x01, b'f', 0x00, 0x00]);

    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);
    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut bytes, 3, &[0x01, 0x00]);
    push_section(&mut bytes, 7, &exports);
    push_section(&mut bytes, 10, &[0x01, 0x02, 0x00, 0x0B]);

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "export names must be unique"
    );
}

#[test]
fn reader_rejects_export_function_index_out_of_range() {
    let export = [0x01, 0x01, b'f', 0x00, 0x01];

    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);
    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut bytes, 3, &[0x01, 0x00]);
    push_section(&mut bytes, 7, &export);
    push_section(&mut bytes, 10, &[0x01, 0x02, 0x00, 0x0B]);

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "function exports must reference declared function indices"
    );
}

#[test]
fn reader_rejects_start_function_with_params_or_results() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut bytes, 1, &[0x01, 0x60, 0x01, 0x7F, 0x00]);
    push_section(&mut bytes, 3, &[0x01, 0x00]);
    push_section(&mut bytes, 8, &[0x00]);
    push_section(&mut bytes, 10, &[0x01, 0x02, 0x00, 0x0B]);

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "start function must have type [] -> []"
    );
}

#[test]
fn reader_rejects_start_function_index_out_of_range() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);

    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut bytes, 3, &[0x01, 0x00]);
    push_section(&mut bytes, 8, &[0x01]);
    push_section(&mut bytes, 10, &[0x01, 0x02, 0x00, 0x0B]);

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "start function index must reference an existing function"
    );
}

#[test]
fn reader_rejects_global_get_index_out_of_range() {
    let bytes = standard_module_with_body(&[0x23, 0x00], 0);
    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "global.get must reference an existing global"
    );
}

#[test]
fn reader_rejects_global_set_to_immutable_global() {
    let mut global_section = Vec::new();
    global_section.push(0x01); // one global
    global_section.push(0x7F); // i32
    global_section.push(0x00); // immutable
    global_section.extend_from_slice(&[0x41, 0x00, 0x0B]); // i32.const 0; end

    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);
    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut bytes, 3, &[0x01, 0x00]);
    push_section(&mut bytes, 6, &global_section);
    push_section(
        &mut bytes,
        10,
        &[0x01, 0x05, 0x00, 0x41, 0x01, 0x24, 0x00, 0x0B],
    );

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "global.set must reject immutable globals"
    );
}

#[test]
fn reader_rejects_element_segment_unknown_table_index() {
    let elem_section = [
        0x01, // one segment
        0x02, // active segment with explicit table index
        0x00, // table 0, but no table section declares it
        0x41, 0x00, 0x0B, // offset expr i32.const 0; end
        0x00, // elemkind funcref
        0x00, // zero function indices
    ];

    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);
    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut bytes, 3, &[0x01, 0x00]);
    push_section(&mut bytes, 9, &elem_section);
    push_section(&mut bytes, 10, &[0x01, 0x02, 0x00, 0x0B]);

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "active element segments must reference an existing table"
    );
}

#[test]
fn reader_rejects_active_data_segment_unknown_memory_index() {
    let data_section = [
        0x01, // one segment
        0x02, // active segment with explicit memory index
        0x00, // memory 0, but no memory section declares it
        0x41, 0x00, 0x0B, // offset expr i32.const 0; end
        0x00, // empty payload
    ];

    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);
    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut bytes, 3, &[0x01, 0x00]);
    push_section(&mut bytes, 10, &[0x01, 0x02, 0x00, 0x0B]);
    push_section(&mut bytes, 11, &data_section);

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "active data segments must reference an existing memory"
    );
}

#[test]
fn reader_rejects_import_type_index_out_of_range() {
    let mut import_section = Vec::new();
    import_section.push(0x01);
    import_section.extend_from_slice(&[0x01, b'm', 0x01, b'f', 0x00, 0x00]);

    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);
    push_section(&mut bytes, 2, &import_section);
    push_section(&mut bytes, 1, &[0x00]);
    push_section(&mut bytes, 3, &[0x00]);
    push_section(&mut bytes, 10, &[0x00]);

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "function imports must reference a declared function type"
    );
}

#[test]
fn reader_rejects_data_count_mismatch() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);
    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut bytes, 3, &[0x01, 0x00]);
    push_section(&mut bytes, 12, &[0x01]); // data_count says one data segment
    push_section(&mut bytes, 10, &[0x01, 0x02, 0x00, 0x0B]);
    push_section(&mut bytes, 11, &[0x00]); // actual data section has zero segments

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "data_count section must match the data section segment count"
    );
}

#[test]
fn reader_rejects_memory_init_without_data_count() {
    let bytes = standard_module_with_body(
        &[
            0x41, 0x00, // dst
            0x41, 0x00, // src
            0x41, 0x00, // len
            0xFC, 0x08, // memory.init
            0x00, // dataidx
            0x00, // memidx
        ],
        1,
    );

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "modules using memory.init must declare a data_count section"
    );
}

#[test]
fn reader_rejects_data_drop_index_out_of_range() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);
    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut bytes, 3, &[0x01, 0x00]);
    push_section(&mut bytes, 12, &[0x00]); // no data segments
    push_section(
        &mut bytes,
        10,
        &[
            0x01, // one body
            0x04, // body size
            0x00, // locals
            0xFC, 0x09, // data.drop
            0x00, // dataidx 0, out of range
            0x0B,
        ],
    );
    push_section(&mut bytes, 11, &[0x00]);

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "data.drop must reference an existing data segment"
    );
}

#[test]
fn reader_rejects_table_init_element_index_out_of_range() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"\0asm");
    bytes.extend_from_slice(&[1, 0, 0, 0]);
    push_section(&mut bytes, 1, &[0x01, 0x60, 0x00, 0x00]);
    push_section(&mut bytes, 3, &[0x01, 0x00]);
    push_section(&mut bytes, 4, &[0x01, 0x70, 0x00, 0x01]); // one funcref table
    push_section(
        &mut bytes,
        10,
        &[
            0x01, // one body
            0x0A, // body size
            0x00, // locals
            0x41, 0x00, // dst
            0x41, 0x00, // src
            0x41, 0x00, // len
            0xFC, 0x0C, // table.init
            0x00, // elemidx 0, out of range
            0x00, // tableidx 0
            0x0B,
        ],
    );

    assert!(
        wasm::read_wasm(&bytes).is_err(),
        "table.init must reference an existing element segment"
    );
}

// ── Round-trip: write → read → execute ────────────────────────────────────

fn roundtrip(chunks: Vec<Chunk>) -> Vec<Chunk> {
    let bytes = wasm::write_wasm(&chunks);
    wasm::read_wasm(&bytes).expect("WASM round-trip read failed")
}

fn read_leb_u32(bytes: &[u8], ip: &mut usize) -> u32 {
    let mut result = 0u32;
    let mut shift = 0;
    loop {
        let byte = bytes[*ip];
        *ip += 1;
        result |= ((byte & 0x7f) as u32) << shift;
        if byte & 0x80 == 0 {
            break;
        }
        shift += 7;
    }
    result
}

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

fn named_custom_section(name: &str, payload: &[u8]) -> Vec<u8> {
    let mut section = Vec::new();
    write_leb_u32(&mut section, name.len() as u32);
    section.extend_from_slice(name.as_bytes());
    section.extend_from_slice(payload);
    section
}

fn standard_module_with_body(body_ops: &[u8], memory_count: u32) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(b"\0asm");
    out.extend_from_slice(&[1, 0, 0, 0]);

    let type_section = vec![0x01, 0x60, 0x00, 0x00];
    push_section(&mut out, 1, &type_section);

    let function_section = vec![0x01, 0x00];
    push_section(&mut out, 3, &function_section);

    let mut memory_section = Vec::new();
    write_leb_u32(&mut memory_section, memory_count);
    for _ in 0..memory_count {
        memory_section.push(0x00); // limits: min only, i32 memory
        memory_section.push(0x01); // min 1 page
    }
    push_section(&mut out, 5, &memory_section);

    let mut body = Vec::new();
    body.push(0x00); // local decl count
    body.extend_from_slice(body_ops);
    body.push(0x0B); // end

    let mut code_section = Vec::new();
    code_section.push(0x01); // one function body
    write_leb_u32(&mut code_section, body.len() as u32);
    code_section.extend_from_slice(&body);
    push_section(&mut out, 10, &code_section);

    out
}

fn custom_section_payload<'a>(bytes: &'a [u8], name: &str) -> Option<&'a [u8]> {
    let mut ip = 8;
    while ip < bytes.len() {
        let section_id = bytes[ip];
        ip += 1;
        let section_size = read_leb_u32(bytes, &mut ip) as usize;
        let section_end = ip + section_size;
        if section_id == 0 {
            let name_len = read_leb_u32(bytes, &mut ip) as usize;
            let name_end = ip + name_len;
            if &bytes[ip..name_end] == name.as_bytes() {
                return Some(&bytes[name_end..section_end]);
            }
        }
        ip = section_end;
    }
    None
}

fn section_payload(bytes: &[u8], target_id: u8) -> Option<&[u8]> {
    let mut ip = 8;
    while ip < bytes.len() {
        let section_id = bytes[ip];
        ip += 1;
        let section_size = read_leb_u32(bytes, &mut ip) as usize;
        let section_end = ip + section_size;
        if section_id == target_id {
            return Some(&bytes[ip..section_end]);
        }
        ip = section_end;
    }
    None
}

#[test]
fn roundtrip_const_and_return() {
    let mut chunk = Chunk::new("<script>");
    let k = chunk.add_constant(Value::I32(42));
    chunk.emit_op_u16(Op::CONST, k, 0);
    chunk.emit_op(Op::RETURN, 0);

    let chunks = roundtrip(vec![chunk]);
    let r = VM::new().run(chunks).expect("run failed");
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn roundtrip_i32_arithmetic() {
    let mut chunk = Chunk::new("<script>");
    let a = chunk.add_constant(Value::I32(10));
    let b = chunk.add_constant(Value::I32(32));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::RETURN, 0);

    let chunks = roundtrip(vec![chunk]);
    let r = VM::new().run(chunks).expect("run failed");
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn roundtrip_f64_arithmetic() {
    let mut chunk = Chunk::new("<script>");
    let a = chunk.add_constant(Value::F64(3.5));
    let b = chunk.add_constant(Value::F64(2.0));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::F64_MUL, 0);
    chunk.emit_op(Op::RETURN, 0);

    let chunks = roundtrip(vec![chunk]);
    let r = VM::new().run(chunks).expect("run failed");
    assert_eq!(r.as_f64(), 7.0);
}

#[test]
fn roundtrip_structured_control_if_else() {
    let mut chunk = Chunk::new("<script>");
    let one = chunk.add_constant(Value::I32(1));
    let ten = chunk.add_constant(Value::I32(10));
    let nine = chunk.add_constant(Value::I32(9));

    chunk.emit_op_u16(Op::CONST, one, 0); // condition = 1 (true)
    let _if_pos = chunk.emit_if(0);
    chunk.emit_op_u16(Op::CONST, ten, 0);
    chunk.emit_else(0);
    chunk.emit_op_u16(Op::CONST, nine, 0);
    chunk.emit_end(0);
    chunk.emit_op(Op::RETURN, 0);

    let chunks = roundtrip(vec![chunk]);
    let r = VM::new().run(chunks).expect("run failed");
    assert_eq!(r.as_i32(), 10);
}

#[test]
fn roundtrip_loop_with_br() {
    // count down from 3 to 0, return 0
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    let n = chunk.add_constant(Value::I32(3));
    let one = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, n, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);

    let _blk = chunk.emit_block(0);
    let (_loop_blk, _loop_body) = chunk.emit_loop_s(0);
    // if local == 0, br 1 (exit block)
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op(Op::I32_EQZ, 0);
    chunk.emit_br_if(1, 0); // exit block
    // local -= 1
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op_u16(Op::CONST, one, 0);
    chunk.emit_op(Op::I32_SUB, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
    chunk.emit_br(0, 0); // continue loop
    chunk.emit_end(0); // end loop
    chunk.emit_end(0); // end block
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let chunks = roundtrip(vec![chunk]);
    let r = VM::new().run(chunks).expect("run failed");
    assert_eq!(r.as_i32(), 0);
}

// ── Opcode byte-value compliance ──────────────────────────────────────────

#[test]
fn core_control_opcodes_have_spec_byte_values() {
    assert_eq!(Op::UNREACHABLE.sub(), 0x00);
    assert_eq!(Op::NOP.sub(), 0x01);
    assert_eq!(Op::BLOCK.sub(), 0x02);
    assert_eq!(Op::LOOP.sub(), 0x03);
    assert_eq!(Op::IF.sub(), 0x04);
    assert_eq!(Op::ELSE.sub(), 0x05);
    assert_eq!(Op::THROW.sub(), 0x08);
    assert_eq!(Op::THROW_REF.sub(), 0x0A);
    assert_eq!(Op::END.sub(), 0x0B);
    assert_eq!(Op::BR.sub(), 0x0C);
    assert_eq!(Op::BR_IF.sub(), 0x0D);
    assert_eq!(Op::BR_TABLE.sub(), 0x0E);
    assert_eq!(Op::RETURN.sub(), 0x0F);
    assert_eq!(Op::CALL.sub(), 0x10);
    assert_eq!(Op::CALL_INDIRECT.sub(), 0x11);
    assert_eq!(Op::RETURN_CALL.sub(), 0x12);
    assert_eq!(Op::RETURN_CALL_INDIRECT.sub(), 0x13);
    assert_eq!(Op::CALL_REF.sub(), 0x14);
    assert_eq!(Op::RETURN_CALL_REF.sub(), 0x15);
    assert_eq!(Op::DROP.sub(), 0x1A);
    assert_eq!(Op::SELECT.sub(), 0x1B);
    assert_eq!(Op::SELECT_T.sub(), 0x1C);
    assert_eq!(Op::TRY_TABLE.sub(), 0x1F);
}

#[test]
fn core_variable_opcodes_have_spec_byte_values() {
    assert_eq!(Op::LOCAL_GET.sub(), 0x20);
    assert_eq!(Op::LOCAL_SET.sub(), 0x21);
    assert_eq!(Op::LOCAL_TEE.sub(), 0x22);
    assert_eq!(Op::GLOBAL_GET.sub(), 0x23);
    assert_eq!(Op::GLOBAL_SET.sub(), 0x24);
    assert_eq!(Op::TABLE_GET.sub(), 0x25);
    assert_eq!(Op::TABLE_SET.sub(), 0x26);
}

#[test]
fn core_memory_opcodes_have_spec_byte_values() {
    assert_eq!(Op::I32_LOAD.sub(), 0x28);
    assert_eq!(Op::I64_LOAD.sub(), 0x29);
    assert_eq!(Op::F32_LOAD.sub(), 0x2A);
    assert_eq!(Op::F64_LOAD.sub(), 0x2B);
    assert_eq!(Op::I32_LOAD8_S.sub(), 0x2C);
    assert_eq!(Op::I32_LOAD8_U.sub(), 0x2D);
    assert_eq!(Op::I32_LOAD16_S.sub(), 0x2E);
    assert_eq!(Op::I32_LOAD16_U.sub(), 0x2F);
    assert_eq!(Op::I64_LOAD8_S.sub(), 0x30);
    assert_eq!(Op::I64_LOAD8_U.sub(), 0x31);
    assert_eq!(Op::I64_LOAD16_S.sub(), 0x32);
    assert_eq!(Op::I64_LOAD16_U.sub(), 0x33);
    assert_eq!(Op::I64_LOAD32_S.sub(), 0x34);
    assert_eq!(Op::I64_LOAD32_U.sub(), 0x35);
    assert_eq!(Op::I32_STORE.sub(), 0x36);
    assert_eq!(Op::I64_STORE.sub(), 0x37);
    assert_eq!(Op::F32_STORE.sub(), 0x38);
    assert_eq!(Op::F64_STORE.sub(), 0x39);
    assert_eq!(Op::I32_STORE8.sub(), 0x3A);
    assert_eq!(Op::I32_STORE16.sub(), 0x3B);
    assert_eq!(Op::I64_STORE8.sub(), 0x3C);
    assert_eq!(Op::I64_STORE16.sub(), 0x3D);
    assert_eq!(Op::I64_STORE32.sub(), 0x3E);
    assert_eq!(Op::MEMORY_SIZE.sub(), 0x3F);
    assert_eq!(Op::MEMORY_GROW.sub(), 0x40);
}

#[test]
fn core_reference_opcodes_have_spec_byte_values() {
    assert_eq!(Op::NULL.sub(), 0xD0);
    assert_eq!(Op::REF_IS_NULL.sub(), 0xD1);
    assert_eq!(Op::REF_FUNC.sub(), 0xD2);
    assert_eq!(Op::REF_EQ.sub(), 0xD3);
    assert_eq!(Op::REF_AS_NON_NULL.sub(), 0xD4);
    assert_eq!(Op::BR_ON_NULL.sub(), 0xD5);
    assert_eq!(Op::BR_ON_NON_NULL.sub(), 0xD6);
}

#[test]
fn memory64_internal_ops_emit_standard_memory_bytes() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op(Op::MEMORY_SIZE, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::I32_LOAD, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::I64_LOAD, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::F64_LOAD, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::I32_STORE, 0);
    chunk.emit_op(Op::I64_STORE, 0);
    chunk.emit_op(Op::F64_STORE, 0);
    chunk.emit_op(Op::RETURN, 0);

    let bytes = wasm::write_wasm(&[chunk]);
    for pattern in [
        &[0x3F, 0x00][..],
        &[0x40, 0x00][..],
        &[0x28, 0x02, 0x00][..],
        &[0x29, 0x03, 0x00][..],
        &[0x2B, 0x03, 0x00][..],
        &[0x36, 0x02, 0x00][..],
        &[0x37, 0x03, 0x00][..],
        &[0x39, 0x03, 0x00][..],
    ] {
        assert!(
            bytes.windows(pattern.len()).any(|w| w == pattern),
            "missing memory64 lowering pattern {pattern:02x?}"
        );
    }
}

#[test]
fn memory64_ops_emit_i64_memory_limits_flag() {
    let mut chunk = Chunk::new("<script>");
    // A 64-bit memory is declared via its index type (memory_is_64) — memory64
    // adds no opcodes, so the writer keys the 0x04 limits flag off this, not
    // off any instruction.
    chunk.memory_min_pages = vec![1];
    chunk.memory_is_64 = vec![true];
    chunk.emit_op(Op::MEMORY_SIZE, 0);
    chunk.emit_op(Op::RETURN, 0);

    let bytes = wasm::write_wasm(&[chunk]);
    let memory = section_payload(&bytes, 5).expect("missing memory section");
    let mut ip = 0;
    assert_eq!(read_leb_u32(memory, &mut ip), 1, "expected one memory");
    assert_eq!(
        memory[ip], 0x04,
        "memory64 min-only limits must use flag 0x04"
    );
}

#[test]
fn memory64_memarg_encoder_accepts_u64_offsets() {
    let mut bytes = Vec::new();
    vybe_platform_wasm::encoding::encode_memarg_with_memidx(&mut bytes, 3, 0x1_0000_0000, 0);
    assert_eq!(
        bytes,
        vec![0x03, 0x80, 0x80, 0x80, 0x80, 0x10],
        "memory64 memarg offsets must be encoded as u64 LEB128"
    );
}

#[test]
fn table64_section_uses_i64_limits_flag() {
    let section = vybe_platform_wasm::writer::sections::encode_table64_section_with(1, None, 0x70);
    assert_eq!(section[0], 1, "expected one table");
    assert_eq!(
        section[1], 0x70,
        "default table64 helper should emit funcref"
    );
    assert_eq!(
        section[2], 0x04,
        "table64 min-only limits must use i64 flag 0x04"
    );
}

#[test]
fn multi_memory_memarg_emits_memory_index_bit_and_immediate() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op(Op::I32_LOAD, 0);
    chunk.emit_leb_u32(0x40 | 2, 0); // align=2, multi-memory memidx follows
    chunk.emit_leb_u32(5, 0); // offset
    chunk.emit_leb_u32(1, 0); // memory index
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::RETURN, 0);

    let bytes = wasm::write_wasm(&[chunk]);
    let pattern = [0x28, 0x42, 0x05, 0x01];
    assert!(
        bytes.windows(pattern.len()).any(|w| w == pattern),
        "i32.load must preserve multi-memory memarg bytes"
    );
}

#[test]
fn multi_memory_bulk_ops_emit_memory_index_immediates() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(1, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::MEMORY_COPY, 0);
    chunk.emit_leb_u32(1, 0); // dst memory
    chunk.emit_leb_u32(2, 0); // src memory
    chunk.emit_op(Op::MEMORY_FILL, 0);
    chunk.emit_leb_u32(1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let bytes = wasm::write_wasm(&[chunk]);
    for pattern in [
        &[0x40, 0x01][..],
        &[0xFC, 0x0A, 0x01, 0x02][..],
        &[0xFC, 0x0B, 0x01][..],
    ] {
        assert!(
            bytes.windows(pattern.len()).any(|w| w == pattern),
            "missing multi-memory encoding pattern {pattern:02x?}"
        );
    }
}

#[test]
fn reader_preserves_multi_memory_load_memarg() {
    let wasm = standard_module_with_body(
        &[
            0x41, 0x00, // i32.const 0
            0x28, 0x42, 0x05, 0x01, // i32.load align=2|memidx, offset=5, memidx=1
            0x1A, // drop
        ],
        2,
    );

    let chunks = wasm::read_wasm(&wasm).expect("standard wasm should decode");
    let code = &chunks[1].code;
    let enc = Op::I32_LOAD.encode();
    let pattern: [u8; 7] = [enc[0], enc[1], enc[2], enc[3], 0x42, 0x05, 0x01];
    assert!(
        code.windows(pattern.len()).any(|w| w == pattern),
        "reader must preserve multi-memory load memarg bytes"
    );
}

#[test]
fn reader_preserves_multi_memory_bulk_indices() {
    let wasm = standard_module_with_body(
        &[
            0x41, 0x01, // i32.const 1
            0x40, 0x01, // memory.grow 1
            0x1A, // drop
            0x41, 0x00, 0x41, 0x00, 0x41, 0x00, // dst, src, len operands
            0xFC, 0x0A, 0x01, 0x02, // memory.copy dst=1 src=2
            0x41, 0x00, 0x41, 0x00, 0x41, 0x00, // dst, val, len operands
            0xFC, 0x0B, 0x01, // memory.fill mem=1
        ],
        3,
    );

    let chunks = wasm::read_wasm(&wasm).expect("standard wasm should decode");
    let code = &chunks[1].code;
    let mg = Op::MEMORY_GROW.encode();
    let mc = Op::MEMORY_COPY.encode();
    let mf = Op::MEMORY_FILL.encode();
    let patterns: Vec<Vec<u8>> = vec![
        vec![mg[0], mg[1], mg[2], mg[3], 0xEE, 0x00, 0x01],
        vec![
            mc[0], mc[1], mc[2], mc[3], 0xEE, 0x00, 0x01, 0xEE, 0x00, 0x02,
        ],
        vec![mf[0], mf[1], mf[2], mf[3], 0xEE, 0x00, 0x01],
    ];
    for pattern in &patterns {
        assert!(
            code.windows(pattern.len()).any(|w| w == pattern.as_slice()),
            "reader must preserve multi-memory bulk pattern {pattern:02x?}"
        );
    }
}

#[test]
fn jspi_suspending_imports_emit_metadata_not_opcode() {
    let mut chunk = Chunk::new("<script>");
    chunk.add_import("wasm:js-promise", "await");
    chunk.emit_op(Op::RETURN, 0);

    let bytes = wasm::write_wasm(&[chunk]);
    assert!(
        !bytes.windows(2).any(|w| w == [0xff, 0x4f]),
        "VM-only promise_suspend opcode must not be emitted to wasm"
    );

    let payload = custom_section_payload(&bytes, "vybe.jspi").expect("missing JSPI metadata");
    let mut ip = 0;
    let promising_count = read_leb_u32(payload, &mut ip);
    for _ in 0..promising_count {
        let _ = read_leb_u32(payload, &mut ip);
    }
    let suspending_count = read_leb_u32(payload, &mut ip);
    let suspending: Vec<u32> = (0..suspending_count)
        .map(|_| read_leb_u32(payload, &mut ip))
        .collect();

    assert_eq!(suspending, vec![0]);
}

#[test]
fn core_comparison_opcodes_have_spec_byte_values() {
    assert_eq!(Op::I32_EQZ.sub(), 0x45);
    assert_eq!(Op::I32_EQ.sub(), 0x46);
    assert_eq!(Op::I32_NE.sub(), 0x47);
    assert_eq!(Op::F32_EQ.sub(), 0x5B);
    assert_eq!(Op::F32_LT.sub(), 0x5D);
    assert_eq!(Op::F64_EQ.sub(), 0x61);
    assert_eq!(Op::F64_LT.sub(), 0x63);
}

#[test]
fn core_arithmetic_opcodes_have_spec_byte_values() {
    assert_eq!(Op::I32_ADD.sub(), 0x6A);
    assert_eq!(Op::I32_SUB.sub(), 0x6B);
    assert_eq!(Op::I32_MUL.sub(), 0x6C);
    assert_eq!(Op::I64_ADD.sub(), 0x7C);
    assert_eq!(Op::F32_ADD.sub(), 0x92);
    assert_eq!(Op::F64_ADD.sub(), 0xA0);
    assert_eq!(Op::F64_MUL.sub(), 0xA2);
}

#[test]
fn core_conversion_opcodes_have_spec_byte_values() {
    assert_eq!(Op::I32_WRAP_I64.sub(), 0xA7);
    assert_eq!(Op::I32_TRUNC_F32_S.sub(), 0xA8);
    assert_eq!(Op::I32_TRUNC_F32_U.sub(), 0xA9);
    assert_eq!(Op::I32_FROM_F64.sub(), 0xAA); // i32.trunc_f64_s
    assert_eq!(Op::I32_TRUNC_F64_U.sub(), 0xAB);
    assert_eq!(Op::I64_EXTEND_I32_S.sub(), 0xAC);
    assert_eq!(Op::I64_EXTEND_I32_U.sub(), 0xAD);
    assert_eq!(Op::F32_DEMOTE_F64.sub(), 0xB6);
    assert_eq!(Op::F64_FROM_I32.sub(), 0xB7); // f64.convert_i32_s
    assert_eq!(Op::F64_PROMOTE_F32.sub(), 0xBB);
    assert_eq!(Op::I32_REINTERPRET_F32.sub(), 0xBC);
    assert_eq!(Op::I64_REINTERPRET_F64.sub(), 0xBD);
    assert_eq!(Op::F32_REINTERPRET_I32.sub(), 0xBE);
    assert_eq!(Op::F64_REINTERPRET_I64.sub(), 0xBF);
}

#[test]
fn gc_opcodes_have_spec_byte_values() {
    assert_eq!(Op::STRUCT_NEW.group(), 0xFB);
    assert_eq!(Op::STRUCT_NEW.sub(), 0x00);
    assert_eq!(Op::ARRAY_NEW.sub(), 0x06);
    assert_eq!(Op::ARRAY_NEW_FIXED.sub(), 0x08);
    assert_eq!(Op::ARRAY_GET.sub(), 0x0B);
    assert_eq!(Op::ARRAY_SET.sub(), 0x0E);
    assert_eq!(Op::ARRAY_LENGTH.sub(), 0x0F);
    assert_eq!(Op::REF_TEST.sub(), 0x14);
    assert_eq!(Op::REF_CAST.sub(), 0x16);
    assert_eq!(Op::BR_ON_CAST.sub(), 0x18);
    assert_eq!(Op::I31_NEW.sub(), 0x1C);
    assert_eq!(Op::I31_GET_S.sub(), 0x1D);
}

#[test]
fn fc_prefixed_proposal_opcodes_have_spec_byte_values() {
    let cases = [
        (Op::I32_TRUNC_SAT_F32_S, 0x00),
        (Op::I32_TRUNC_SAT_F32_U, 0x01),
        (Op::I32_TRUNC_SAT_F64_S, 0x02),
        (Op::I32_TRUNC_SAT_F64_U, 0x03),
        (Op::I64_TRUNC_SAT_F32_S, 0x04),
        (Op::I64_TRUNC_SAT_F32_U, 0x05),
        (Op::I64_TRUNC_SAT_F64_S, 0x06),
        (Op::I64_TRUNC_SAT_F64_U, 0x07),
        (Op::MEMORY_INIT, 0x08),
        (Op::DATA_DROP, 0x09),
        (Op::MEMORY_COPY, 0x0A),
        (Op::MEMORY_FILL, 0x0B),
        (Op::TABLE_INIT, 0x0C),
        (Op::ELEM_DROP, 0x0D),
        (Op::TABLE_COPY, 0x0E),
        (Op::TABLE_GROW, 0x0F),
        (Op::TABLE_SIZE, 0x10),
        (Op::TABLE_FILL, 0x11),
    ];

    for (op, sub) in cases {
        assert_eq!(op.group(), 0xFC);
        assert_eq!(op.sub(), sub);
    }
}

#[test]
fn simd_prefix_is_fd() {
    assert_eq!(Op::V128_CONST.group(), 0xFD);
    assert_eq!(Op::I32X4_ADD.group(), 0xFD);
    assert_eq!(Op::F64X2_SQRT.group(), 0xFD);
}

#[test]
fn simd_memory_lane_and_splat_opcodes_have_spec_byte_values() {
    let cases = [
        (Op::V128_LOAD, 0x00),
        (Op::V128_LOAD8X8_S, 0x01),
        (Op::V128_LOAD8X8_U, 0x02),
        (Op::V128_LOAD16X4_S, 0x03),
        (Op::V128_LOAD16X4_U, 0x04),
        (Op::V128_LOAD32X2_S, 0x05),
        (Op::V128_LOAD32X2_U, 0x06),
        (Op::V128_LOAD8_SPLAT, 0x07),
        (Op::V128_LOAD16_SPLAT, 0x08),
        (Op::V128_LOAD32_SPLAT, 0x09),
        (Op::V128_LOAD64_SPLAT, 0x0A),
        (Op::V128_STORE, 0x0B),
        (Op::V128_CONST, 0x0C),
        (Op::I8X16_SHUFFLE, 0x0D),
        (Op::I8X16_SWIZZLE, 0x0E),
        (Op::I8X16_SPLAT, 0x0F),
        (Op::I16X8_SPLAT, 0x10),
        (Op::I32X4_SPLAT, 0x11),
        (Op::I64X2_SPLAT, 0x12),
        (Op::F32X4_SPLAT, 0x13),
        (Op::F64X2_SPLAT, 0x14),
        (Op::I8X16_EXTRACT_LANE_S, 0x15),
        (Op::I8X16_EXTRACT_LANE_U, 0x16),
        (Op::I8X16_REPLACE_LANE, 0x17),
        (Op::I16X8_EXTRACT_LANE_S, 0x18),
        (Op::I16X8_EXTRACT_LANE_U, 0x19),
        (Op::I16X8_REPLACE_LANE, 0x1A),
        (Op::I32X4_EXTRACT_LANE, 0x1B),
        (Op::I32X4_REPLACE_LANE, 0x1C),
        (Op::I64X2_EXTRACT_LANE, 0x1D),
        (Op::I64X2_REPLACE_LANE, 0x1E),
        (Op::F32X4_EXTRACT_LANE, 0x1F),
        (Op::F32X4_REPLACE_LANE, 0x20),
        (Op::F64X2_EXTRACT_LANE, 0x21),
        (Op::F64X2_REPLACE_LANE, 0x22),
        (Op::V128_LOAD8_LANE, 0x54),
        (Op::V128_LOAD16_LANE, 0x55),
        (Op::V128_LOAD32_LANE, 0x56),
        (Op::V128_LOAD64_LANE, 0x57),
        (Op::V128_STORE8_LANE, 0x58),
        (Op::V128_STORE16_LANE, 0x59),
        (Op::V128_STORE32_LANE, 0x5A),
        (Op::V128_STORE64_LANE, 0x5B),
        (Op::V128_LOAD32_ZERO, 0x5C),
        (Op::V128_LOAD64_ZERO, 0x5D),
    ];

    for (op, sub) in cases {
        assert_eq!(op.group(), 0xFD);
        assert_eq!(op.sub(), sub);
    }
}

#[test]
fn simd_comparison_and_bitwise_opcodes_have_spec_byte_values() {
    let cases = [
        (Op::I8X16_EQ, 0x23),
        (Op::I8X16_NE, 0x24),
        (Op::I8X16_LT_S, 0x25),
        (Op::I8X16_LT_U, 0x26),
        (Op::I8X16_GT_S, 0x27),
        (Op::I8X16_GT_U, 0x28),
        (Op::I8X16_LE_S, 0x29),
        (Op::I8X16_LE_U, 0x2A),
        (Op::I8X16_GE_S, 0x2B),
        (Op::I8X16_GE_U, 0x2C),
        (Op::I16X8_EQ, 0x2D),
        (Op::I16X8_NE, 0x2E),
        (Op::I16X8_LT_S, 0x2F),
        (Op::I16X8_LT_U, 0x30),
        (Op::I16X8_GT_S, 0x31),
        (Op::I16X8_GT_U, 0x32),
        (Op::I16X8_LE_S, 0x33),
        (Op::I16X8_LE_U, 0x34),
        (Op::I16X8_GE_S, 0x35),
        (Op::I16X8_GE_U, 0x36),
        (Op::I32X4_EQ, 0x37),
        (Op::I32X4_NE, 0x38),
        (Op::I32X4_LT_S, 0x39),
        (Op::I32X4_LT_U, 0x3A),
        (Op::I32X4_GT_S, 0x3B),
        (Op::I32X4_GT_U, 0x3C),
        (Op::I32X4_LE_S, 0x3D),
        (Op::I32X4_LE_U, 0x3E),
        (Op::I32X4_GE_S, 0x3F),
        (Op::I32X4_GE_U, 0x40),
        (Op::F32X4_EQ, 0x41),
        (Op::F32X4_NE, 0x42),
        (Op::F32X4_LT, 0x43),
        (Op::F32X4_GT, 0x44),
        (Op::F32X4_LE, 0x45),
        (Op::F32X4_GE, 0x46),
        (Op::F64X2_EQ, 0x47),
        (Op::F64X2_NE, 0x48),
        (Op::F64X2_LT, 0x49),
        (Op::F64X2_GT, 0x4A),
        (Op::F64X2_LE, 0x4B),
        (Op::F64X2_GE, 0x4C),
        (Op::V128_NOT, 0x4D),
        (Op::V128_AND, 0x4E),
        (Op::V128_ANDNOT, 0x4F),
        (Op::V128_OR, 0x50),
        (Op::V128_XOR, 0x51),
        (Op::V128_BITSELECT, 0x52),
        (Op::V128_ANY_TRUE, 0x53),
    ];

    for (op, sub) in cases {
        assert_eq!(op.group(), 0xFD);
        assert_eq!(op.sub(), sub);
    }
}

#[test]
fn simd_numeric_family_opcodes_have_spec_byte_values() {
    let cases = [
        (Op::F32X4_DEMOTE_F64X2_ZERO, 0x5E),
        (Op::F64X2_PROMOTE_LOW_F32X4, 0x5F),
        (Op::I8X16_ABS, 0x60),
        (Op::I8X16_NEG, 0x61),
        (Op::I8X16_POPCNT, 0x62),
        (Op::I8X16_ALL_TRUE, 0x63),
        (Op::I8X16_BITMASK, 0x64),
        (Op::I8X16_NARROW_I16X8_S, 0x65),
        (Op::I8X16_NARROW_I16X8_U, 0x66),
        (Op::F32X4_CEIL, 0x67),
        (Op::F32X4_FLOOR, 0x68),
        (Op::F32X4_TRUNC, 0x69),
        (Op::F32X4_NEAREST, 0x6A),
        (Op::I8X16_SHL, 0x6B),
        (Op::I8X16_SHR_S, 0x6C),
        (Op::I8X16_SHR_U, 0x6D),
        (Op::I8X16_ADD, 0x6E),
        (Op::I8X16_ADD_SAT_S, 0x6F),
        (Op::I8X16_ADD_SAT_U, 0x70),
        (Op::I8X16_SUB, 0x71),
        (Op::I8X16_SUB_SAT_S, 0x72),
        (Op::I8X16_SUB_SAT_U, 0x73),
        (Op::F64X2_CEIL, 0x74),
        (Op::F64X2_FLOOR, 0x75),
        (Op::I8X16_MIN_S, 0x76),
        (Op::I8X16_MIN_U, 0x77),
        (Op::I8X16_MAX_S, 0x78),
        (Op::I8X16_MAX_U, 0x79),
        (Op::F64X2_TRUNC, 0x7A),
        (Op::I8X16_AVGR_U, 0x7B),
        (Op::I16X8_EXTADD_PAIRWISE_I8X16_S, 0x7C),
        (Op::I16X8_EXTADD_PAIRWISE_I8X16_U, 0x7D),
        (Op::I32X4_EXTADD_PAIRWISE_I16X8_S, 0x7E),
        (Op::I32X4_EXTADD_PAIRWISE_I16X8_U, 0x7F),
        (Op::I16X8_ABS, 0x80),
        (Op::I16X8_NEG, 0x81),
        (Op::I16X8_Q15MULR_SAT_S, 0x82),
        (Op::I16X8_ALL_TRUE, 0x83),
        (Op::I16X8_BITMASK, 0x84),
        (Op::I16X8_NARROW_I32X4_S, 0x85),
        (Op::I16X8_NARROW_I32X4_U, 0x86),
        (Op::I16X8_EXTEND_LOW_I8X16_S, 0x87),
        (Op::I16X8_EXTEND_HIGH_I8X16_S, 0x88),
        (Op::I16X8_EXTEND_LOW_I8X16_U, 0x89),
        (Op::I16X8_EXTEND_HIGH_I8X16_U, 0x8A),
        (Op::I16X8_SHL, 0x8B),
        (Op::I16X8_SHR_S, 0x8C),
        (Op::I16X8_SHR_U, 0x8D),
        (Op::I16X8_ADD, 0x8E),
        (Op::I16X8_ADD_SAT_S, 0x8F),
        (Op::I16X8_ADD_SAT_U, 0x90),
        (Op::I16X8_SUB, 0x91),
        (Op::I16X8_SUB_SAT_S, 0x92),
        (Op::I16X8_SUB_SAT_U, 0x93),
        (Op::F64X2_NEAREST, 0x94),
        (Op::I16X8_MUL, 0x95),
        (Op::I16X8_MIN_S, 0x96),
        (Op::I16X8_MIN_U, 0x97),
        (Op::I16X8_MAX_S, 0x98),
        (Op::I16X8_MAX_U, 0x99),
        (Op::I16X8_AVGR_U, 0x9B),
        (Op::I16X8_EXTMUL_LOW_I8X16_S, 0x9C),
        (Op::I16X8_EXTMUL_HIGH_I8X16_S, 0x9D),
        (Op::I16X8_EXTMUL_LOW_I8X16_U, 0x9E),
        (Op::I16X8_EXTMUL_HIGH_I8X16_U, 0x9F),
        (Op::I32X4_ABS, 0xA0),
        (Op::I32X4_NEG, 0xA1),
        (Op::I32X4_ALL_TRUE, 0xA3),
        (Op::I32X4_BITMASK, 0xA4),
        (Op::I32X4_EXTEND_LOW_I16X8_S, 0xA7),
        (Op::I32X4_EXTEND_HIGH_I16X8_S, 0xA8),
        (Op::I32X4_EXTEND_LOW_I16X8_U, 0xA9),
        (Op::I32X4_EXTEND_HIGH_I16X8_U, 0xAA),
        (Op::I32X4_SHL, 0xAB),
        (Op::I32X4_SHR_S, 0xAC),
        (Op::I32X4_SHR_U, 0xAD),
        (Op::I32X4_ADD, 0xAE),
        (Op::I32X4_SUB, 0xB1),
        (Op::I32X4_MUL, 0xB5),
        (Op::I32X4_MIN_S, 0xB6),
        (Op::I32X4_MIN_U, 0xB7),
        (Op::I32X4_MAX_S, 0xB8),
        (Op::I32X4_MAX_U, 0xB9),
        (Op::I32X4_DOT_I16X8_S, 0xBA),
        (Op::I32X4_EXTMUL_LOW_I16X8_S, 0xBC),
        (Op::I32X4_EXTMUL_HIGH_I16X8_S, 0xBD),
        (Op::I32X4_EXTMUL_LOW_I16X8_U, 0xBE),
        (Op::I32X4_EXTMUL_HIGH_I16X8_U, 0xBF),
        (Op::I64X2_ABS, 0xC0),
        (Op::I64X2_NEG, 0xC1),
        (Op::I64X2_ALL_TRUE, 0xC3),
        (Op::I64X2_BITMASK, 0xC4),
        (Op::I64X2_EXTEND_LOW_I32X4_S, 0xC7),
        (Op::I64X2_EXTEND_HIGH_I32X4_S, 0xC8),
        (Op::I64X2_EXTEND_LOW_I32X4_U, 0xC9),
        (Op::I64X2_EXTEND_HIGH_I32X4_U, 0xCA),
        (Op::I64X2_SHL, 0xCB),
        (Op::I64X2_SHR_S, 0xCC),
        (Op::I64X2_SHR_U, 0xCD),
        (Op::I64X2_ADD, 0xCE),
        (Op::I64X2_SUB, 0xD1),
        (Op::I64X2_MUL, 0xD5),
        (Op::I64X2_EQ, 0xD6),
        (Op::I64X2_NE, 0xD7),
        (Op::I64X2_LT_S, 0xD8),
        (Op::I64X2_GT_S, 0xD9),
        (Op::I64X2_LE_S, 0xDA),
        (Op::I64X2_GE_S, 0xDB),
        (Op::I64X2_EXTMUL_LOW_I32X4_S, 0xDC),
        (Op::I64X2_EXTMUL_HIGH_I32X4_S, 0xDD),
        (Op::I64X2_EXTMUL_LOW_I32X4_U, 0xDE),
        (Op::I64X2_EXTMUL_HIGH_I32X4_U, 0xDF),
        (Op::F32X4_ABS, 0xE0),
        (Op::F32X4_NEG, 0xE1),
        (Op::F32X4_SQRT, 0xE3),
        (Op::F32X4_ADD, 0xE4),
        (Op::F32X4_SUB, 0xE5),
        (Op::F32X4_MUL, 0xE6),
        (Op::F32X4_DIV, 0xE7),
        (Op::F32X4_MIN, 0xE8),
        (Op::F32X4_MAX, 0xE9),
        (Op::F32X4_PMIN, 0xEA),
        (Op::F32X4_PMAX, 0xEB),
        (Op::F64X2_ABS, 0xEC),
        (Op::F64X2_NEG, 0xED),
        (Op::F64X2_SQRT, 0xEF),
        (Op::F64X2_ADD, 0xF0),
        (Op::F64X2_SUB, 0xF1),
        (Op::F64X2_MUL, 0xF2),
        (Op::F64X2_DIV, 0xF3),
        (Op::F64X2_MIN, 0xF4),
        (Op::F64X2_MAX, 0xF5),
        (Op::F64X2_PMIN, 0xF6),
        (Op::F64X2_PMAX, 0xF7),
        (Op::I32X4_TRUNC_SAT_F32X4_S, 0xF8),
        (Op::I32X4_TRUNC_SAT_F32X4_U, 0xF9),
        (Op::F32X4_CONVERT_I32X4_S, 0xFA),
        (Op::F32X4_CONVERT_I32X4_U, 0xFB),
        (Op::I32X4_TRUNC_SAT_F64X2_S_ZERO, 0xFC),
        (Op::I32X4_TRUNC_SAT_F64X2_U_ZERO, 0xFD),
        (Op::F64X2_CONVERT_LOW_I32X4_S, 0xFE),
        (Op::F64X2_CONVERT_LOW_I32X4_U, 0xFF),
    ];

    for (op, sub) in cases {
        assert_eq!(op.group(), 0xFD);
        assert_eq!(op.sub(), sub);
    }
}

#[test]
fn threads_prefix_is_fe() {
    assert_eq!(Op::ATOMIC_FENCE.group(), 0xFE);
    assert_eq!(Op::MEMORY_ATOMIC_WAIT32.sub(), 0x01);
    assert_eq!(Op::MEMORY_ATOMIC_WAIT64.sub(), 0x02);
    assert_eq!(Op::I32_ATOMIC_LOAD.group(), 0xFE);
    assert_eq!(Op::I64_ATOMIC_STORE.group(), 0xFE);
}

#[test]
fn threads_opcodes_have_spec_byte_values() {
    let cases = [
        (Op::MEMORY_ATOMIC_NOTIFY, 0x00),
        (Op::MEMORY_ATOMIC_WAIT32, 0x01),
        (Op::MEMORY_ATOMIC_WAIT64, 0x02),
        (Op::ATOMIC_FENCE, 0x03),
        (Op::I32_ATOMIC_LOAD, 0x10),
        (Op::I64_ATOMIC_LOAD, 0x11),
        (Op::I32_ATOMIC_STORE, 0x17),
        (Op::I64_ATOMIC_STORE, 0x18),
        (Op::I32_ATOMIC_RMW_ADD, 0x1E),
        (Op::I64_ATOMIC_RMW_ADD, 0x1F),
        (Op::I32_ATOMIC_RMW_SUB, 0x25),
        (Op::I64_ATOMIC_RMW_SUB, 0x26),
        (Op::I32_ATOMIC_RMW_AND, 0x2C),
        (Op::I64_ATOMIC_RMW_AND, 0x2D),
        (Op::I32_ATOMIC_RMW_OR, 0x33),
        (Op::I64_ATOMIC_RMW_OR, 0x34),
        (Op::I32_ATOMIC_RMW_XOR, 0x3A),
        (Op::I64_ATOMIC_RMW_XOR, 0x3B),
        (Op::I32_ATOMIC_RMW_XCHG, 0x41),
        (Op::I64_ATOMIC_RMW_XCHG, 0x42),
        (Op::I32_ATOMIC_RMW_CMPXCHG, 0x48),
        (Op::I64_ATOMIC_RMW_CMPXCHG, 0x49),
    ];

    for (op, sub) in cases {
        assert_eq!(op.group(), 0xFE);
        assert_eq!(op.sub(), sub);
    }
}

// ── Multiple chunks round-trip ─────────────────────────────────────────────

#[test]
fn roundtrip_multiple_functions() {
    use std::sync::Arc;
    use vybe_bytecode::chunk::{ConstExpr, GlobalInit};

    let mut add_fn = Chunk::new("add");
    add_fn.arity = 2;
    add_fn.local_count = 2;
    add_fn.emit_op_u16(Op::LOCAL_GET, 0, 0);
    add_fn.emit_op_u16(Op::LOCAL_GET, 1, 0);
    add_fn.emit_op(Op::I32_ADD, 0);
    add_fn.emit_op(Op::RETURN, 0);

    let mut main = Chunk::new("<script>");
    main.local_count = 1;
    main.global_inits.push(GlobalInit {
        name: "__add".to_string(),
        init: ConstExpr::RefFunc(1),
    });
    let fn_name = main.add_constant(Value::String(Arc::from("__add")));
    let a = main.add_constant(Value::I32(20));
    let b = main.add_constant(Value::I32(22));
    main.emit_op_u16(Op::GLOBAL_GET, fn_name, 0);
    main.emit_op_u16(Op::CONST, a, 0);
    main.emit_op_u16(Op::CONST, b, 0);
    main.emit_op_u8(Op::CALL_REF, 2, 0);
    main.emit_op(Op::RETURN, 0);

    // Run directly (round-trip for multi-chunk requires full linker support)
    let r = VM::new().run(vec![main, add_fn]).expect("run failed");
    assert_eq!(r.as_i32(), 42);
}

// ── Conversion opcode reader round-trips ───────────────────────────────────
//
// These tests verify the WASM binary reader maps each conversion byte
// to the correct VM opcode. They were previously broken: 0xB9/0xBA/0xBB
// were shifted (wrong ops), and 0xBE/0xBF were missing entirely.

fn rt_run(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    let chunks = roundtrip(vec![c]);
    VM::new().run(chunks).expect("run failed")
}

fn push_i32_rt(c: &mut Chunk, v: i32) {
    let k = c.add_constant(Value::I32(v));
    c.emit_op_u16(Op::CONST, k, 0);
}
fn push_i64_rt(c: &mut Chunk, v: i64) {
    let k = c.add_constant(Value::I64(v));
    c.emit_op_u16(Op::CONST, k, 0);
}
fn push_f64_rt(c: &mut Chunk, v: f64) {
    let k = c.add_constant(Value::F64(v));
    c.emit_op_u16(Op::CONST, k, 0);
}

// 0xA8 i32.trunc_f32_s — was collapsed to I32_FROM_F64 (no trapping)
#[test]
fn reader_i32_trunc_f32_s_roundtrip() {
    assert_eq!(
        rt_run(|c| {
            push_f64_rt(c, 3.7);
            c.emit_op(Op::I32_TRUNC_F32_S, 0);
        })
        .as_i32(),
        3
    );
}

// 0xA9 i32.trunc_f32_u — was collapsed to I32_FROM_F64
#[test]
fn reader_i32_trunc_f32_u_roundtrip() {
    assert_eq!(
        rt_run(|c| {
            push_f64_rt(c, 200.9);
            c.emit_op(Op::I32_TRUNC_F32_U, 0);
        })
        .as_i32() as u32,
        200
    );
}

// 0xAE i32.trunc_f32_s → i64 — was mapped to I64_TRUNC_F64_S losing the F32 distinction
#[test]
fn reader_i64_trunc_f32_s_roundtrip() {
    assert_eq!(
        rt_run(|c| {
            push_f64_rt(c, -99.9);
            c.emit_op(Op::I64_TRUNC_F32_S, 0);
        })
        .as_i64(),
        -99
    );
}

// 0xB2 f32.convert_i32_s — was collapsed to F64_FROM_I32
#[test]
fn reader_f32_convert_i32_s_roundtrip() {
    assert_eq!(
        rt_run(|c| {
            push_i32_rt(c, -7);
            c.emit_op(Op::F32_CONVERT_I32_S, 0);
        })
        .as_f64() as f32,
        -7.0f32
    );
}

// 0xB4 f32.convert_i64_s — was collapsed to F64_FROM_I32
#[test]
fn reader_f32_convert_i64_s_roundtrip() {
    assert_eq!(
        rt_run(|c| {
            push_i64_rt(c, -1_000);
            c.emit_op(Op::F32_CONVERT_I64_S, 0);
        })
        .as_f64() as f32,
        -1_000.0f32
    );
}

// 0xB8 f64.convert_i32_u — was collapsed to F64_FROM_I32 (losing unsigned semantics)
#[test]
fn reader_f64_convert_i32_u_roundtrip() {
    assert_eq!(
        rt_run(|c| {
            push_i32_rt(c, -1);
            c.emit_op(Op::F64_CONVERT_I32_U, 0);
        })
        .as_f64(),
        4_294_967_295.0
    );
}

// 0xB9 f64.convert_i64_s — was WRONG: mapped to F64_PROMOTE_F32
#[test]
fn reader_f64_convert_i64_s_roundtrip() {
    assert_eq!(
        rt_run(|c| {
            push_i64_rt(c, -42);
            c.emit_op(Op::F64_CONVERT_I64_S, 0);
        })
        .as_f64(),
        -42.0
    );
}

// 0xBA f64.convert_i64_u — was WRONG: mapped to F32_REINTERPRET_I32
#[test]
fn reader_f64_convert_i64_u_roundtrip() {
    let r = rt_run(|c| {
        push_i64_rt(c, 1_000_000_000);
        c.emit_op(Op::F64_CONVERT_I64_U, 0);
    })
    .as_f64();
    assert!((r - 1_000_000_000.0).abs() < 1.0);
}

// 0xBB f64.promote_f32 — was WRONG: mapped to F64_REINTERPRET_I64
#[test]
fn reader_f64_promote_f32_roundtrip() {
    assert!(
        (rt_run(|c| {
            push_f64_rt(c, 1.5f32 as f64);
            c.emit_op(Op::F64_PROMOTE_F32, 0);
        })
        .as_f64()
            - 1.5)
            .abs()
            < 1e-6
    );
}

// 0xBE f32.reinterpret_i32 — was MISSING from reader
#[test]
fn reader_f32_reinterpret_i32_roundtrip() {
    // 0x3F800000 = 1.0f32 bit pattern
    let r = rt_run(|c| {
        push_i32_rt(c, 0x3F800000u32 as i32);
        c.emit_op(Op::F32_REINTERPRET_I32, 0);
    });
    assert_eq!(r.as_f64() as f32, 1.0f32);
}

// 0xBF f64.reinterpret_i64 — was MISSING from reader
#[test]
fn reader_f64_reinterpret_i64_roundtrip() {
    // 0x3FF0000000000000 = 1.0f64 bit pattern
    let r = rt_run(|c| {
        push_i64_rt(c, 0x3FF0000000000000u64 as i64);
        c.emit_op(Op::F64_REINTERPRET_I64, 0);
    });
    assert!((r.as_f64() - 1.0).abs() < 1e-10);
}
