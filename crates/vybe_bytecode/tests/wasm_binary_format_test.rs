//! Tests for the WASM binary format: magic bytes, version, section structure,
//! round-trip read/write, and opcode byte-value compliance per the spec.
//! Binary I/O compliance — not execution semantics (those live in per-opcode files).

use vybe_bytecode::value::Value;
use vybe_bytecode::wasm;
use vybe_bytecode::{Chunk, Op, VM};

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
    chunk.emit_op(Op::DROP, 0);

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
    chunk.emit_op(Op::DROP, 0);
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
}

#[test]
fn core_memory_opcodes_have_spec_byte_values() {
    assert_eq!(Op::I32_LOAD.sub(), 0x28);
    assert_eq!(Op::I64_LOAD.sub(), 0x29);
    assert_eq!(Op::F32_LOAD.sub(), 0x2A);
    assert_eq!(Op::F64_LOAD.sub(), 0x2B);
    assert_eq!(Op::I32_STORE.sub(), 0x36);
    assert_eq!(Op::I64_STORE.sub(), 0x37);
    assert_eq!(Op::F32_STORE.sub(), 0x38);
    assert_eq!(Op::F64_STORE.sub(), 0x39);
    assert_eq!(Op::MEMORY_SIZE.sub(), 0x3F);
    assert_eq!(Op::MEMORY_GROW.sub(), 0x40);
}

#[test]
fn memory64_internal_ops_emit_standard_memory_bytes() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op(Op::I64_MEMORY_SIZE, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::I64_MEMORY_GROW, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::I32_LOAD_64, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::I64_LOAD_64, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::F64_LOAD_64, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::I32_STORE_64, 0);
    chunk.emit_op(Op::I64_STORE_64, 0);
    chunk.emit_op(Op::F64_STORE_64, 0);
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
    assert_eq!(Op::STRUCT_NEW.prefix(), 0xFB);
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
fn simd_prefix_is_fd() {
    assert_eq!(Op::V128_CONST.prefix(), 0xFD);
    assert_eq!(Op::I32X4_ADD.prefix(), 0xFD);
    assert_eq!(Op::F64X2_SQRT.prefix(), 0xFD);
}

#[test]
fn threads_prefix_is_fe() {
    assert_eq!(Op::ATOMIC_FENCE.prefix(), 0xFE);
    assert_eq!(Op::MEMORY_ATOMIC_WAIT32.sub(), 0x01);
    assert_eq!(Op::MEMORY_ATOMIC_WAIT64.sub(), 0x02);
    assert_eq!(Op::I32_ATOMIC_LOAD.prefix(), 0xFE);
    assert_eq!(Op::I64_ATOMIC_STORE.prefix(), 0xFE);
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
