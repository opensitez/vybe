use super::*;
use super::raw_memory_tests::{invoke, module, operators, typed};

fn get(chunk: &mut Chunk, slot: u16) {
    chunk.emit_op(Op::LOCAL_GET, 0);
    chunk.code.extend_from_slice(&slot.to_be_bytes());
}

fn vector(chunk: &mut Chunk, bytes: [u8; 16]) {
    chunk.emit_op(Op::V128_CONST, 0);
    chunk.code.extend_from_slice(&bytes);
}

fn memory(chunk: &mut Chunk, op: Op, wide: bool, index: u32, offset: u64) {
    chunk.emit_op(op, 0);
    write_leb128_u32(&mut chunk.code, 0x80 | if wide { 0x100 } else { 0 } | if index != 0 { 0x40 } else { 0 });
    write_leb128_u64(&mut chunk.code, offset);
    if index != 0 { write_leb128_u32(&mut chunk.code, index); }
}

fn check_vector(chunk: &mut Chunk, expected: [u8; 16]) {
    vector(chunk, expected);
    chunk.emit_op(Op::I8X16_EQ, 0);
    chunk.emit_op(Op::I8X16_BITMASK, 0);
    chunk.emit_op(Op::RETURN, 0);
}

fn fixture(wide: bool, index: u32) -> Chunk {
    let mut chunk = typed(if wide { "i64" } else { "i32" }, "i32", wide);
    if index != 0 { chunk.memory_is_64 = vec![false, wide]; }
    chunk
}

fn address(wide: bool, value: i64) -> Value {
    if wide { Value::I64(value) } else { Value::I32(value as i32) }
}

fn assert_native(bytes: &[u8]) {
    let ops = operators(bytes);
    assert!(!ops.iter().any(|op| op.starts_with("Call {") || op.starts_with("LocalSet")),
        "typed SIMD must not introduce boxing calls or scratch spills: {ops:?}");
}

#[test]
fn native_simd_load_shapes_preserve_bits_memargs_and_bounds() {
    let seed = [0x81, 0xfe, 3, 0x80, 5, 0xff, 7, 0x90, 9, 10, 11, 12, 13, 14, 15, 16];
    let shapes = [Op::V128_LOAD, Op::V128_LOAD8X8_S, Op::V128_LOAD8X8_U,
        Op::V128_LOAD16X4_S, Op::V128_LOAD16X4_U, Op::V128_LOAD32X2_S, Op::V128_LOAD32X2_U,
        Op::V128_LOAD8_SPLAT, Op::V128_LOAD16_SPLAT, Op::V128_LOAD32_SPLAT,
        Op::V128_LOAD64_SPLAT, Op::V128_LOAD32_ZERO, Op::V128_LOAD64_ZERO];
    for wide in [false, true] {
        for index in [0, 1] {
            for op in shapes {
                let mut expected = [0; 16];
                match op.sub() {
                    0 => expected = seed,
                    1 | 2 => for lane in 0..8 {
                        let value = if op.sub() == 1 { i16::from(seed[lane] as i8) } else { i16::from(seed[lane]) };
                        expected[lane * 2..lane * 2 + 2].copy_from_slice(&value.to_le_bytes());
                    },
                    3 | 4 => for lane in 0..4 {
                        let value = u16::from_le_bytes(seed[lane * 2..lane * 2 + 2].try_into().unwrap());
                        let value = if op.sub() == 3 { i32::from(value as i16) } else { i32::from(value) };
                        expected[lane * 4..lane * 4 + 4].copy_from_slice(&value.to_le_bytes());
                    },
                    5 | 6 => for lane in 0..2 {
                        let value = u32::from_le_bytes(seed[lane * 4..lane * 4 + 4].try_into().unwrap());
                        let value = if op.sub() == 5 { i64::from(value as i32) } else { i64::from(value) };
                        expected[lane * 8..lane * 8 + 8].copy_from_slice(&value.to_le_bytes());
                    },
                    7..=10 => {
                        let width = 1usize << (op.sub() - 7);
                        for lane in expected.chunks_exact_mut(width) { lane.copy_from_slice(&seed[..width]); }
                    }
                    0x5c | 0x5d => {
                        let width = if op.sub() == 0x5c { 4 } else { 8 };
                        expected[..width].copy_from_slice(&seed[..width]);
                    }
                    _ => unreachable!(),
                }
                let mut chunk = fixture(wide, index);
                get(&mut chunk, 0);
                vector(&mut chunk, seed);
                memory(&mut chunk, Op::V128_STORE, wide, index, 5);
                get(&mut chunk, 0);
                memory(&mut chunk, op, wide, index, 5);
                check_vector(&mut chunk, expected);
                let bytes = module(chunk, wide);
                assert_native(&bytes);
                for at in [3, 65515] {
                    assert_eq!(invoke(&bytes, &[address(wide, at)], "i32").unwrap(), Value::I32(65535), "{op:?}, memory {index}, wide {wide}");
                }
                for at in [-1, 65516] { assert!(invoke(&bytes, &[address(wide, at)], "i32").is_err()); }
            }
        }
    }
}

#[test]
fn native_simd_lane_loads_and_stores_keep_lane_order_and_access_width() {
    let seed = [0x81, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 0xfe];
    for wide in [false, true] {
        for index in [0, 1] {
            for width_log in 0..4u16 {
                let width = 1usize << width_log;
                let lane = (16 / width - 1) as u8;
                for store in [false, true] {
                    let mut chunk = fixture(wide, index);
                    get(&mut chunk, 0);
                    vector(&mut chunk, if store { [0; 16] } else { seed });
                    memory(&mut chunk, Op::V128_STORE, wide, index, 5);
                    get(&mut chunk, 0);
                    vector(&mut chunk, if store { seed } else { [0; 16] });
                    let op = Op::decode(0xfd, (if store { 0x58 } else { 0x54 }) + width_log).unwrap();
                    memory(&mut chunk, op, wide, index, 5);
                    chunk.code.push(lane);
                    let mut expected = [0; 16];
                    if store {
                        expected[..width].copy_from_slice(&seed[16 - width..]);
                        get(&mut chunk, 0);
                        memory(&mut chunk, Op::V128_LOAD, wide, index, 5);
                    } else {
                        expected[16 - width..].copy_from_slice(&seed[..width]);
                    }
                    check_vector(&mut chunk, expected);
                    let bytes = module(chunk, wide);
                    assert_native(&bytes);
                    assert_eq!(invoke(&bytes, &[address(wide, 3)], "i32").unwrap(), Value::I32(65535));
                    assert!(invoke(&bytes, &[address(wide, -1)], "i32").is_err());
                }
                // A lane access at the very end is valid even when a full
                // v128 load would overrun. Do not widen the observable access.
                let mut chunk = fixture(wide, index);
                get(&mut chunk, 0);
                vector(&mut chunk, [0; 16]);
                memory(&mut chunk, Op::decode(0xfd, 0x54 + width_log).unwrap(), wide, index, 0);
                chunk.code.push(lane);
                check_vector(&mut chunk, [0; 16]);
                let bytes = module(chunk, wide);
                assert_eq!(invoke(&bytes, &[address(wide, (65536 - width) as i64)], "i32").unwrap(), Value::I32(65535));
                assert!(invoke(&bytes, &[address(wide, (65537 - width) as i64)], "i32").is_err());
            }
        }
    }
}

#[test]
fn native_simd_memory64_large_offsets_trap_without_truncation() {
    let mut chunk = fixture(true, 0);
    get(&mut chunk, 0);
    memory(&mut chunk, Op::V128_LOAD32_ZERO, true, 0, 1u64 << 32);
    check_vector(&mut chunk, [0; 16]);
    let bytes = module(chunk, true);
    assert_native(&bytes);
    assert!(invoke(&bytes, &[Value::I64(0)], "i32").is_err());
}

#[test]
fn declared_vector_formals_results_and_local_copies_remain_native() {
    let mut chunk = typed("v128", "v128", false);
    chunk.local_count = 2;
    get(&mut chunk, 0);
    chunk.emit_op(Op::LOCAL_TEE, 0);
    chunk.code.extend_from_slice(&1u16.to_be_bytes());
    get(&mut chunk, 1);
    chunk.emit_op(Op::V128_XOR, 0);
    chunk.emit_op(Op::V128_NOT, 0);
    chunk.emit_op(Op::RETURN, 0);
    let bytes = module(chunk, false);
    assert_native(&bytes);
    assert_eq!(invoke(&bytes, &[Value::V128([0x81; 16])], "v128").unwrap(), Value::V128([0xff; 16]));
}

#[test]
fn native_float_splat_extract_preserves_nan_payloads_and_signed_zero() {
    for wide in [false, true] {
        let mut chunk = typed(if wide { "i64" } else { "i32" }, if wide { "i64" } else { "i32" }, false);
        get(&mut chunk, 0);
        chunk.emit_op(if wide { Op::F64_REINTERPRET_I64 } else { Op::F32_REINTERPRET_I32 }, 0);
        chunk.emit_op(if wide { Op::F64X2_SPLAT } else { Op::F32X4_SPLAT }, 0);
        chunk.emit_op(if wide { Op::F64X2_EXTRACT_LANE } else { Op::F32X4_EXTRACT_LANE }, 0);
        chunk.code.push(if wide { 1 } else { 3 });
        chunk.emit_op(if wide { Op::I64_REINTERPRET_F64 } else { Op::I32_REINTERPRET_F32 }, 0);
        chunk.emit_op(Op::RETURN, 0);
        let bytes = module(chunk, false);
        assert_native(&bytes);
        let patterns: &[u64] = if wide {
            &[0, 1 << 63, 0x7ff0000000000001, 0xfff0000000000001, 0x7ff8123456789abc]
        } else { &[0, 1 << 31, 0x7f800001, 0xff800001, 0x7fc12345] };
        for &bits in patterns {
            let value = address(wide, bits as i64);
            assert_eq!(invoke(&bytes, &[value.clone()], if wide { "i64" } else { "i32" }).unwrap(), value);
        }
    }
}

#[test]
fn admitted_simd_stack_shapes_validate_against_standard_wasm_types() {
    let spelling = |ty| match ty {
        TYPE_I32 => "i32", TYPE_I64 => "i64", TYPE_F32 => "f32", TYPE_F64 => "f64", 0x7b => "v128",
        _ => panic!("unexpected SIMD stack type"),
    };
    let mut checked = 0;
    for sub in 0..=255 {
        let Some(op) = Op::decode(0xfd, sub) else { continue; };
        let Some((inputs, result)) = raw_simd_shape(op) else { continue; };
        let params = inputs.iter().map(|&ty| spelling(ty)).collect::<Vec<_>>().join(",");
        let mut chunk = typed(&params, spelling(result), false);
        for slot in 0..inputs.len() { get(&mut chunk, slot as u16); }
        chunk.emit_op(op, 0);
        match op.operand_format() {
            OperandFormat::U8 => chunk.code.push(0),
            OperandFormat::Shuffle => chunk.code.extend(0..16u8),
            _ => {},
        }
        chunk.emit_op(Op::RETURN, 0);
        let bytes = module(chunk, false);
        assert_native(&bytes);
        checked += 1;
    }
    assert!(checked > 100, "the typed SIMD shape gate must cover the full admitted set");
}

#[test]
fn native_simd_probe_is_smaller_than_the_mixed_value_control() {
    let kernel = |params: &str| {
        let mut chunk = typed(params, "i32", false);
        get(&mut chunk, 0);
        memory(&mut chunk, Op::V128_LOAD, false, 0, 0);
        check_vector(&mut chunk, [0; 16]);
        module(chunk, false)
    };
    let native = kernel("i32");
    // An unused reference formal deliberately prevents whole-body numeric
    // proof. The actual memory/SIMD operations and result are identical.
    let mixed = kernel("i32,externref");
    let body_size = |bytes: &[u8]| wasmparser::Parser::new(0).parse_all(bytes)
        .filter_map(|payload| match payload.unwrap() {
            wasmparser::Payload::CodeSectionEntry(body) => Some(body.get_binary_reader().bytes_remaining()), _ => None,
        }).nth(1).unwrap();
    let native_size = body_size(&native);
    let mixed_size = body_size(&mixed);
    let calls = |bytes: &[u8]| operators(bytes).iter().filter(|op| op.starts_with("Call {")).count();
    assert_eq!(calls(&native), 0);
    assert!(calls(&mixed) > 0);
    assert!(native_size < mixed_size);
    for at in [0, 3, 65520] {
        assert_eq!(super::raw_memory_tests::invoke_full_module(&native, &[Value::I32(at)]).unwrap(), Value::I32(65535));
        assert_eq!(super::raw_memory_tests::invoke_full_module(&mixed, &[Value::I32(at), Value::Null]).unwrap(), Value::I32(65535));
    }
    for at in [-1, 65521] {
        assert!(super::raw_memory_tests::invoke_full_module(&native, &[Value::I32(at)]).is_err());
        assert!(super::raw_memory_tests::invoke_full_module(&mixed, &[Value::I32(at), Value::Null]).is_err());
    }
    eprintln!("SIMD probe body: native={native_size} bytes, mixed={mixed_size} bytes; calls 0 vs {}", calls(&mixed));
}

#[test]
fn compiler_elided_simd_memargs_use_shared_natural_alignment() {
    let seed = [0x81, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 0xfe];
    for wide in [false, true] {
        let mut chunk = fixture(wide, 0);
        get(&mut chunk, 0);
        vector(&mut chunk, seed);
        chunk.emit_op(Op::V128_STORE, 0);
        get(&mut chunk, 0);
        chunk.emit_op(Op::V128_LOAD, 0);
        check_vector(&mut chunk, seed);
        let bytes = module(chunk, wide);
        assert_native(&bytes);
        for at in [0, 3, 65520] {
            assert_eq!(invoke(&bytes, &[address(wide, at)], "i32").unwrap(), Value::I32(65535));
        }
        for at in [-1, 65521] { assert!(invoke(&bytes, &[address(wide, at)], "i32").is_err()); }
    }
}
