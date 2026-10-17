use super::*;
use vybe_runtime::{Value, VM};

pub(super) fn typed(params: &str, result: &str, wide: bool) -> Chunk {
    let mut chunk = Chunk::new("typed_memory");
    chunk.arity = params.split(',').filter(|s| !s.is_empty()).count() as u8;
    chunk.local_count = u16::from(chunk.arity);
    chunk.result_arity = u8::from(!result.is_empty());
    chunk.func_sig = Some((params.to_owned(), result.to_owned()));
    chunk.memory_is_64 = vec![wide];
    chunk
}

pub(super) fn module(mut chunk: Chunk, wide: bool) -> Vec<u8> {
    let mut script = Chunk::new("<script>");
    let memories = chunk.memory_is_64.len().max(1);
    script.memory_min_pages = vec![1; memories];
    script.memory_max_pages = vec![Some(2); memories];
    script.memory_is_64 = if chunk.memory_is_64.is_empty() {
        vec![wide]
    } else { chunk.memory_is_64.clone() };
    script.data_segments = std::mem::take(&mut chunk.data_segments);
    script.active_data_segments = std::mem::take(&mut chunk.active_data_segments);
    script.emit_i32_const(0, 0);
    script.emit_op(Op::RETURN, 0);
    let bytes = super::super::write_wasm_with_options(&[script, chunk],
        super::super::WasmWriteOptions {
            include_name_section: true,
            ..Default::default()
        });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).expect("typed memory output must validate");
    bytes
}

pub(super) fn operators(bytes: &[u8]) -> Vec<String> {
    let body = wasmparser::Parser::new(0).parse_all(bytes).filter_map(|payload| {
        match payload.unwrap() {
            wasmparser::Payload::CodeSectionEntry(body) => Some(body),
            _ => None,
        }
    }).nth(1).unwrap();
    body.get_operators_reader().unwrap().into_iter()
        .map(|op| format!("{:?}", op.unwrap())).collect()
}

pub(super) fn invoke(bytes: &[u8], arguments: &[Value], result: &str) -> Result<Value, String> {
    // Independently exercise the exact self-contained native body, not a
    // reimplementation, with its standard signature and memory declarations.
    // Full-module controls below also retain imports and callable globals.
    let mut native = b"\0asm\x01\0\0\0".to_vec();
    let section = |output: &mut Vec<u8>, id: u8, payload: &[u8]| {
        output.push(id);
        write_leb128_u32(output, payload.len() as u32);
        output.extend_from_slice(payload);
    };
    let mut signature = vec![1, 0x60];
    write_leb128_u32(&mut signature, arguments.len() as u32);
    for value in arguments {
        signature.push(match value {
            Value::I32(_) => TYPE_I32,
            Value::I64(_) => TYPE_I64,
            Value::V128(_) => 0x7b,
            _ => panic!("unsupported native test argument"),
        });
    }
    if result.is_empty() {
        signature.push(0);
    } else {
        signature.extend_from_slice(&[1, match result { "i64" => TYPE_I64, "v128" => 0x7b, _ => TYPE_I32 }]);
    }
    section(&mut native, 1, &signature);
    section(&mut native, 3, &[1, 0]);
    // Copy actual standard declarations, including exact Memory64 and passive
    // data encodings. DataCount precedes code; Data follows code.
    let mut cursor = 8;
    let mut data_section = None;
    while cursor < bytes.len() {
        let start = cursor;
        let id = bytes[cursor];
        cursor += 1;
        let length = read_leb_u32(bytes, &mut cursor) as usize;
        let end = cursor.checked_add(length).expect("test section end");
        assert!(end <= bytes.len());
        match id {
            5 | 12 => native.extend_from_slice(&bytes[start..end]),
            11 => data_section = Some(&bytes[start..end]),
            _ => {}
        }
        cursor = end;
    }
    let mut function_body = None;
    let mut defined = 0;
    for payload in wasmparser::Parser::new(0).parse_all(bytes) {
        match payload.unwrap() {
            wasmparser::Payload::CodeSectionEntry(body) => {
                if defined == 1 {
                    let range = body.range();
                    let start = usize::try_from(range.start).expect("test body start fits usize");
                    let end = usize::try_from(range.end).expect("test body end fits usize");
                    function_body = Some(&bytes[start..end]);
                }
                defined += 1;
            }
            _ => {}
        }
    }
    let function_body = function_body.expect("native test body");
    let mut code = vec![1];
    write_leb128_u32(&mut code, function_body.len() as u32);
    code.extend_from_slice(function_body);
    section(&mut native, 10, &code);
    if let Some(data) = data_section { native.extend_from_slice(data); }
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&native).expect("isolated actual native body must validate");
    let mut chunks = crate::reader::read_wasm(&native).map_err(|e| format!("{e:?}"))?;
    chunks[0].code.clear();
    chunks[0].direct_defined_calls.clear();
    chunks[0].emit_op_u16(Op::REF_FUNC, 1, 0);
    chunks[0].code.push(0);
    chunks[0].emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    let function = vm.run(chunks).map_err(|e| format!("{e:?}"))?;
    vm.invoke(&function, arguments).map_err(|e| format!("{e:?}"))
}

pub(super) fn invoke_full_module(bytes: &[u8], arguments: &[Value]) -> Result<Value, String> {
    // Keep imports, native GC globals, active segments and original function
    // indices. Replace only the final entry result, after initialization.
    let mut chunks = crate::reader::read_wasm(bytes).map_err(|e| format!("{e:?}"))?;
    let target = chunks.iter().position(|chunk| chunk.name == "typed_memory")
        .ok_or_else(|| "typed memory fixture was not retained by the reader".to_owned())?;
    let mut ip = 0;
    let mut last = None;
    while ip < chunks[0].code.len() {
        let op = Op::from_code_at(&chunks[0].code, ip);
        last = Some((ip, op));
        ip += opcode_size(op, &chunks[0].code, ip);
    }
    let (end, op) = last.expect("fixture entry initialization");
    assert_eq!(op, Op::RETURN, "fixture entry has an explicit final return");
    chunks[0].code.truncate(end);
    chunks[0].direct_defined_calls.retain(|at, _| *at < end);
    for _ in 0..chunks[0].result_arity { chunks[0].emit_op(Op::DROP, 0); }
    chunks[0].result_arity = 1;
    chunks[0].emit_op_u16(Op::REF_FUNC, u16::try_from(target).expect("test function index"), 0);
    chunks[0].code.push(0);
    chunks[0].emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    vybe_runtime::js_builtins::register(&mut vm);
    vybe_platform_ecma::register(&mut vm);
    let function = vm.run(chunks).map_err(|e| format!("{e:?}"))?;
    vm.invoke(&function, arguments).map_err(|e| format!("{e:?}"))
}

#[test]
fn native_integer_memory_stores_loads_and_traps_keep_raw_boundaries() {
    for wide in [false, true] {
        for (store, load, value, result, expected) in [
            (Op::I32_STORE, Op::I32_LOAD, Value::I32(-123), "i32", Value::I32(-123)),
            (Op::I32_STORE8, Op::I32_LOAD8_S, Value::I32(0x81), "i32", Value::I32(-127)),
            (Op::I32_STORE16, Op::I32_LOAD16_U, Value::I32(0x8001), "i32", Value::I32(0x8001)),
            (Op::I64_STORE, Op::I64_LOAD, Value::I64(i64::MIN + 5), "i64", Value::I64(i64::MIN + 5)),
            (Op::I64_STORE32, Op::I64_LOAD32_S, Value::I64(0x8000_0001), "i64", Value::I64(-2147483647)),
        ] {
            let address = if wide { "i64" } else { "i32" };
            let mut chunk = typed(&format!("{address},{result}"), result, wide);
            chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
            chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
            chunk.emit_op(store, 0);
            chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
            chunk.emit_op(load, 0);
            chunk.emit_op(Op::RETURN, 0);
            let bytes = module(chunk, wide);
            assert!(!operators(&bytes).iter().any(|op| op.starts_with("Call {")));
            let address_value = if wide { Value::I64(3) } else { Value::I32(3) };
            assert_eq!(invoke(&bytes, &[address_value, value.clone()], result).unwrap(), expected);
            let invalid = if wide { Value::I64(65536) } else { Value::I32(65536) };
            assert!(invoke(&bytes, &[invalid, value], result).is_err());
        }
    }
}

#[test]
fn native_bulk_fill_copy_and_memory_size_need_no_boxing_or_spills() {
    let mut chunk = typed("i32,i32,i32", "i32", false);
    for slot in [0, 1, 2] { chunk.emit_op_u16(Op::LOCAL_GET, slot, 0); }
    chunk.emit_op_idx(Op::MEMORY_FILL, 0u32, 0);
    for slot in [1, 0, 2] { chunk.emit_op_u16(Op::LOCAL_GET, slot, 0); }
    chunk.emit_op_idx(Op::MEMORY_COPY, 0u32, 0);
    chunk.emit_leb_u32(0, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    chunk.emit_op(Op::I32_LOAD8_U, 0);
    chunk.emit_op(Op::RETURN, 0);
    let bytes = module(chunk, false);
    let ops = operators(&bytes);
    assert!(!ops.iter().any(|op| op.starts_with("Call {") || op.starts_with("LocalSet {")));
    assert_eq!(invoke(&bytes, &[Value::I32(0), Value::I32(171), Value::I32(16)], "i32").unwrap(), Value::I32(171));
    assert!(invoke(&bytes, &[Value::I32(65530), Value::I32(171), Value::I32(16)], "i32").is_err());

    for wide in [false, true] {
        let mut chunk = typed("", if wide { "i64" } else { "i32" }, wide);
        chunk.emit_op_idx(Op::MEMORY_SIZE, 0u32, 0);
        chunk.emit_op(Op::RETURN, 0);
        let bytes = module(chunk, wide);
        assert_eq!(invoke(&bytes, &[], if wide { "i64" } else { "i32" }).unwrap(), if wide { Value::I64(1) } else { Value::I32(1) });
    }
}

#[test]
fn boxed_selection_keeps_dynamic_value_abi_and_validates() {
    let mut chunk = Chunk::new("boxed_selection");
    for value in [11, 22, 1] { chunk.emit_i32_const(value, 0); }
    chunk.emit_op(Op::SELECT, 0);
    chunk.emit_op(Op::RETURN, 0);
    let bytes = module(chunk, false);
    assert!(operators(&bytes).iter().any(|op| op.starts_with("TypedSelect ")));
}

#[test]
fn native_selection_tee_and_unsigned_extension_preserve_word_bits() {
    for select in [Op::SELECT, Op::SELECT_T] {
        let mut chunk = typed("i32,i32,i32", "i64", false);
        chunk.local_count = 4;
        for slot in [0, 1, 2] { chunk.emit_op_u16(Op::LOCAL_GET, slot, 0); }
        chunk.emit_op(select, 0);
        chunk.emit_op_u16(Op::LOCAL_TEE, 3, 0);
        chunk.emit_op(Op::I64_EXTEND_I32_U, 0);
        chunk.emit_op(Op::RETURN, 0);
        let bytes = module(chunk, false);
        assert!(!operators(&bytes).iter().any(|op| op.starts_with("Call {")));
        assert_eq!(invoke(&bytes, &[Value::I32(-1), Value::I32(7), Value::I32(1)], "i64").unwrap(), Value::I64(4294967295));
        assert_eq!(invoke(&bytes, &[Value::I32(-1), Value::I32(7), Value::I32(0)], "i64").unwrap(), Value::I64(7));
    }
}

#[test]
fn native_mixed_memory_copies_keep_spec_address_and_length_widths() {
    for destination64 in [false, true] {
        for source64 in [false, true] {
            let length64 = destination64 && source64;
            let lane = |wide| if wide { "i64" } else { "i32" };
            let mut chunk = typed(&format!("{},{},i32,{}",
                lane(destination64), lane(source64), lane(length64)), "i32", destination64);
            chunk.memory_is_64 = vec![destination64, source64];
            for slot in [1, 2, 3] { chunk.emit_op_u16(Op::LOCAL_GET, slot, 0); }
            if source64 && !length64 { chunk.emit_op(Op::I64_EXTEND_I32_U, 0); }
            chunk.emit_op_idx(Op::MEMORY_FILL, 1u32, 0);
            for slot in [0, 1, 3] { chunk.emit_op_u16(Op::LOCAL_GET, slot, 0); }
            chunk.emit_op_idx(Op::MEMORY_COPY, 0u32, 0);
            chunk.emit_leb_u32(1, 0);
            chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
            chunk.emit_op(Op::I32_LOAD8_U, 0);
            chunk.emit_op(Op::RETURN, 0);
            let bytes = module(chunk, destination64);
            assert!(!operators(&bytes).iter().any(|op| op.starts_with("Call {") || op.starts_with("LocalSet {")));
            let word = |wide, n| if wide { Value::I64(n) } else { Value::I32(n as i32) };
            assert_eq!(invoke(&bytes, &[
                word(destination64, 7), word(source64, 11), Value::I32(0xA5), word(length64, 16)
            ], "i32").unwrap(), Value::I32(0xA5));
        }
    }
}

#[test]
fn native_memory_initialization_and_growth_preserve_results_and_traps() {
    for wide in [false, true] {
        let lane = if wide { "i64" } else { "i32" };
        let word = |n| if wide { Value::I64(n) } else { Value::I32(n as i32) };
        let mut chunk = typed(&format!("{lane},i32,i32"), "i32", wide);
        chunk.data_segments = vec![vec![0x81, 2, 3, 4]];
        for slot in [0, 1, 2] { chunk.emit_op_u16(Op::LOCAL_GET, slot, 0); }
        chunk.emit_op_idx(Op::MEMORY_INIT, 0u32, 0);
        chunk.emit_leb_u32(0, 0);
        chunk.emit_op_idx(Op::DATA_DROP, 0u32, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        chunk.emit_op(Op::I32_LOAD8_S, 0);
        chunk.emit_op(Op::RETURN, 0);
        let bytes = module(chunk, wide);
        assert!(!operators(&bytes).iter().any(|op| op.starts_with("Call {")));
        assert_eq!(invoke(&bytes, &[word(3), Value::I32(0), Value::I32(4)], "i32").unwrap(), Value::I32(-127));
        assert!(invoke(&bytes, &[word(3), Value::I32(3), Value::I32(4)], "i32").is_err());

        let mut chunk = typed(lane, lane, wide);
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        chunk.emit_op_idx(Op::MEMORY_GROW, 0u32, 0);
        chunk.emit_op(Op::RETURN, 0);
        let bytes = module(chunk, wide);
        assert_eq!(invoke(&bytes, &[word(1)], lane).unwrap(), word(1));
        assert_eq!(invoke(&bytes, &[word(2)], lane).unwrap(), word(-1));
    }

    let mut chunk = typed("i32,i32,i32", "", false);
    for slot in [0, 1, 2] { chunk.emit_op_u16(Op::LOCAL_GET, slot, 0); }
    chunk.emit_op_idx(Op::MEMORY_FILL, 0u32, 0);
    chunk.emit_op(Op::RETURN, 0);
    let bytes = module(chunk, false);
    assert!(!operators(&bytes).iter().any(|op| op.starts_with("Call {") || op.starts_with("LocalSet {")));
    assert!(invoke(&bytes, &[Value::I32(0), Value::I32(171), Value::I32(16)], "").is_ok());
}

#[test]
fn native_store_load_body_is_smaller_than_mixed_value_abi_control() {
    let mut native = typed("i32,i32", "i32", false);
    for slot in [0, 1] { native.emit_op_u16(Op::LOCAL_GET, slot, 0); }
    native.emit_op(Op::I32_STORE, 0);
    native.emit_op_u16(Op::LOCAL_GET, 0, 0);
    native.emit_op(Op::I32_LOAD, 0);
    native.emit_op(Op::RETURN, 0);
    // An unused noninteger formal deliberately retains the existing mixed
    // value-ABI path, with the exact same operation stream. No fake no-op
    // instructions are added to inflate this structural control.
    let mut mixed = native.clone();
    mixed.func_sig = Some(("i32,i32,externref".into(), "i32".into()));
    mixed.arity = 3;
    mixed.local_count = 3;
    let native = module(native, false);
    let mixed = module(mixed, false);
    let body_size = |bytes: &[u8]| -> u64 {
        wasmparser::Parser::new(0).parse_all(bytes).filter_map(|payload| {
            match payload.unwrap() {
                wasmparser::Payload::CodeSectionEntry(body) => {
                    let range = body.range();
                    Some(range.end - range.start)
                }
                _ => None,
            }
        }).nth(1).unwrap()
    };
    let native_size = body_size(&native);
    let mixed_size = body_size(&mixed);
    let calls = operators(&mixed).iter().filter(|op| op.starts_with("Call {")).count();
    eprintln!("typed store/load body: native={native_size} bytes, mixed-ABI control={mixed_size} bytes, removed calls={calls}");
    assert!(native_size < mixed_size);
    assert!(calls > 0);
}

// Build exact bytecode operands without relying on a source-language cast.
fn numeric_op(chunk: &mut Chunk, op: Op) {
    chunk.code.extend_from_slice(&op.group().to_be_bytes());
    chunk.code.extend_from_slice(&op.sub().to_be_bytes());
}
fn numeric_get(chunk: &mut Chunk, slot: u16) {
    numeric_op(chunk, Op::LOCAL_GET);
    chunk.code.extend_from_slice(&slot.to_be_bytes());
}

#[test]
fn native_float_memory_preserves_bits_and_bounds() {
    for wide in [false, true] {
        for is_double in [false, true] {
            let address = if wide { "i64" } else { "i32" };
            let bits = if is_double { "i64" } else { "i32" };
            let mut chunk = typed(&format!("{address},{bits}"), bits, wide);
            numeric_get(&mut chunk, 0);
            numeric_get(&mut chunk, 1);
            numeric_op(&mut chunk, if is_double { Op::F64_REINTERPRET_I64 } else { Op::F32_REINTERPRET_I32 });
            numeric_op(&mut chunk, if is_double { Op::F64_STORE } else { Op::F32_STORE });
            numeric_get(&mut chunk, 0);
            numeric_op(&mut chunk, if is_double { Op::F64_LOAD } else { Op::F32_LOAD });
            numeric_op(&mut chunk, if is_double { Op::I64_REINTERPRET_F64 } else { Op::I32_REINTERPRET_F32 });
            numeric_op(&mut chunk, Op::RETURN);
            let bytes = module(chunk, wide);
            assert!(!operators(&bytes).iter().any(|op| op.starts_with("Call ")));
            let patterns: &[u64] = if is_double {
                &[0, 1 << 63, 0x7ff0000000000000, 0xfff0000000000000, 0x7ff0000000000001, 0x7ff8123456789abc, 0x3ff0000000000000]
            } else {
                &[0, 1 << 31, 0x7f800000, 0xff800000, 0x7f800001, 0x7fc12345, 0x3f800000]
            };
            for &pattern in patterns {
                let value = if is_double { Value::I64(pattern as i64) } else { Value::I32(pattern as i32) };
                let addr = if wide { Value::I64(3) } else { Value::I32(3) };
                assert_eq!(invoke(&bytes, &[addr, value.clone()], bits).unwrap(), value);
            }
            let addr = if wide { Value::I64(65533) } else { Value::I32(65533) };
            let value = if is_double { Value::I64(0) } else { Value::I32(0) };
            assert!(invoke(&bytes, &[addr, value], bits).is_err());
        }
    }
}

#[test]
fn native_float_arithmetic_retains_typed_local() {
    for is_double in [false, true] {
        let bits = if is_double { "i64" } else { "i32" };
        let mut chunk = typed(bits, bits, false);
        chunk.local_count = 2;
        numeric_get(&mut chunk, 0);
        numeric_op(&mut chunk, if is_double { Op::F64_REINTERPRET_I64 } else { Op::F32_REINTERPRET_I32 });
        numeric_op(&mut chunk, Op::LOCAL_TEE);
        chunk.code.extend_from_slice(&1u16.to_be_bytes());
        numeric_get(&mut chunk, 1);
        numeric_op(&mut chunk, if is_double { Op::F64_MUL } else { Op::F32_MUL });
        numeric_op(&mut chunk, if is_double { Op::F64_SQRT } else { Op::F32_SQRT });
        numeric_op(&mut chunk, if is_double { Op::I64_REINTERPRET_F64 } else { Op::I32_REINTERPRET_F32 });
        numeric_op(&mut chunk, Op::RETURN);
        let bytes = module(chunk, false);
        assert!(!operators(&bytes).iter().any(|op| op.starts_with("Call ")));
        let three = if is_double { Value::I64(3f64.to_bits() as i64) } else { Value::I32(3f32.to_bits() as i32) };
        assert_eq!(invoke(&bytes, &[three.clone()], bits).unwrap(), three);
    }
}

#[test]
fn native_float_truncation_preserves_traps_and_saturation() {
    for saturating in [false, true] {
        let mut chunk = typed("i64", "i32", false);
        numeric_get(&mut chunk, 0);
        numeric_op(&mut chunk, Op::F64_REINTERPRET_I64);
        numeric_op(&mut chunk, if saturating { Op::I32_TRUNC_SAT_F64_S } else { Op::decode(0, 0xAA).expect("standard i32.trunc_f64_s") });
        numeric_op(&mut chunk, Op::RETURN);
        let bytes = module(chunk, false);
        for (input, expected) in [(1.9f64, 1), (-1.9, -1), (2147483647.0, i32::MAX), (-2147483648.0, i32::MIN)] {
            assert_eq!(invoke(&bytes, &[Value::I64(input.to_bits() as i64)], "i32").unwrap(), Value::I32(expected));
        }
        for (input, expected) in [(f64::NAN, 0), (f64::INFINITY, i32::MAX), (f64::NEG_INFINITY, i32::MIN), (2147483648.0, i32::MAX)] {
            let result = invoke(&bytes, &[Value::I64(input.to_bits() as i64)], "i32");
            if saturating { assert_eq!(result.unwrap(), Value::I32(expected)); }
            else { assert!(result.is_err()); }
        }
    }
}

#[test]
fn native_float_memory_is_smaller_than_mixed_value_control() {
    for double in [false, true] {
        let bits = if double { "i64" } else { "i32" };
        let mut raw = typed(&format!("i32,{bits}"), bits, false);
        numeric_get(&mut raw, 0);
        numeric_get(&mut raw, 1);
        numeric_op(&mut raw, if double { Op::F64_REINTERPRET_I64 } else { Op::F32_REINTERPRET_I32 });
        numeric_op(&mut raw, if double { Op::F64_STORE } else { Op::F32_STORE });
        numeric_get(&mut raw, 0);
        numeric_op(&mut raw, if double { Op::F64_LOAD } else { Op::F32_LOAD });
        numeric_op(&mut raw, if double { Op::I64_REINTERPRET_F64 } else { Op::I32_REINTERPRET_F32 });
        numeric_op(&mut raw, Op::RETURN);
        let mut mixed = raw.clone();
        mixed.arity = 3;
        mixed.local_count = 3;
        mixed.func_sig = Some((format!("i32,{bits},externref"), bits.to_owned()));
        let raw = module(raw, false);
        let mixed = module(mixed, false);
        let size = |bytes: &[u8]| {
            wasmparser::Parser::new(0).parse_all(bytes).filter_map(|payload| match payload.unwrap() {
                wasmparser::Payload::CodeSectionEntry(body) => Some(body.range().end - body.range().start),
                _ => None,
            }).nth(1).unwrap()
        };
        let calls = |bytes: &[u8]| operators(bytes).iter().filter(|op| op.starts_with("Call ")).count();
        assert!(size(&raw) < size(&mixed));
        assert_eq!(calls(&raw), 0);
        assert!(calls(&mixed) > 0);
        eprintln!("native f{} memory body: {} vs {} bytes; conversion calls: {} vs {}", if double { 64 } else { 32 }, size(&raw), size(&mixed), calls(&raw), calls(&mixed));
    }
}

#[test]
fn native_float_constants_and_results_keep_core_signature() {
    for double in [false, true] {
        let mut chunk = typed("", if double { "f64" } else { "f32" }, false);
        chunk.local_count = 1;
        numeric_op(&mut chunk, if double { Op::F64_CONST } else { Op::F32_CONST });
        if double { chunk.code.extend_from_slice(&9f64.to_le_bytes()); }
        else { chunk.code.extend_from_slice(&9f32.to_le_bytes()); }
        numeric_op(&mut chunk, Op::LOCAL_SET);
        chunk.code.extend_from_slice(&0u16.to_be_bytes());
        numeric_get(&mut chunk, 0);
        numeric_op(&mut chunk, if double { Op::F64_SQRT } else { Op::F32_SQRT });
        numeric_op(&mut chunk, Op::RETURN);
        let bytes = module(chunk, false);
        assert!(!operators(&bytes).iter().any(|op| op.starts_with("Call ")));
    }
}

#[test]
fn mixed_value_float_memory_bitcasts_do_not_quiet_nan_payloads() {
    for wide in [false, true] {
        let address = if wide { "i64" } else { "i32" };
        let mut chunk = typed(&format!("{address},i32"), "i32", wide);
        // A structured body deliberately uses the existing mixed-value path.
        numeric_op(&mut chunk, Op::BLOCK);
        chunk.code.extend_from_slice(&[0, 0]);
        numeric_op(&mut chunk, Op::END);
        numeric_get(&mut chunk, 0);
        numeric_get(&mut chunk, 1);
        numeric_op(&mut chunk, Op::F32_REINTERPRET_I32);
        numeric_op(&mut chunk, Op::F32_STORE);
        numeric_get(&mut chunk, 0);
        numeric_op(&mut chunk, Op::F32_LOAD);
        numeric_op(&mut chunk, Op::I32_REINTERPRET_F32);
        numeric_op(&mut chunk, Op::RETURN);
        let bytes = module(chunk, wide);
        for pattern in [0u32, 0x80000000, 0x7f800000, 0xff800000, 0x7f800001, 0xff800001, 0x7fc12345] {
            let addr = if wide { Value::I64(3) } else { Value::I32(3) };
            assert_eq!(invoke_full_module(&bytes, &[addr, Value::I32(pattern as i32)]).unwrap(), Value::I32(pattern as i32));
        }
        let addr = if wide { Value::I64(65533) } else { Value::I32(65533) };
        assert!(invoke_full_module(&bytes, &[addr, Value::I32(0)]).is_err());
    }
}

fn numeric_indirect_fixture(kind: u8, wide_table: bool, mixed: bool, tail: bool, mismatch: bool) -> Vec<u8> {
    numeric_indirect_fixture_repeated(kind, wide_table, mixed, tail, mismatch, 1)
}

fn numeric_indirect_fixture_repeated(kind: u8, wide_table: bool, mixed: bool, tail: bool, mismatch: bool, repetitions: usize) -> Vec<u8> {
    assert!(repetitions > 0 && (!tail || repetitions == 1));
    let mut bytes = b"\0asm\x01\0\0\0".to_vec();
    let bits = if matches!(kind, TYPE_I64 | TYPE_F64) { TYPE_I64 } else { TYPE_I32 };
    let address = if wide_table { TYPE_I64 } else { TYPE_I32 };
    let mut types = vec![if mismatch { 3 } else { 2 }];
    let signature = |out: &mut Vec<u8>, params: &[u8], result: u8| {
        out.push(TYPE_FUNC);
        write_leb128_u32(out, params.len() as u32);
        out.extend_from_slice(params);
        out.extend_from_slice(&[1, result]);
    };
    signature(&mut types, &[if mismatch { TYPE_I32 } else { kind }], if mismatch { TYPE_I32 } else { kind });
    signature(&mut types, &[bits, address], bits);
    if mismatch { signature(&mut types, &[kind], kind); }
    write_section(&mut bytes, SECTION_TYPE, &types);
    write_section(&mut bytes, SECTION_FUNCTION, &[2, 0, 1]);
    write_section(&mut bytes, 4, &[1, 0x70, if wide_table { 5 } else { 1 }, 2, 2]);
    let mut exports = vec![1];
    write_name(&mut exports, "typed_memory");
    exports.extend_from_slice(&[0, 1]);
    write_section(&mut bytes, SECTION_EXPORT, &exports);
    write_section(&mut bytes, 9, &[1, 0, if wide_table { 0x42 } else { 0x41 }, 0, 0x0b, 1, 0]);
    let identity = [0, 0x20, 0, 0x0f, 0x0b];
    let mut caller = vec![0];
    if mixed { caller.extend_from_slice(&[0x02, 0x40, 0x0b]); }
    for iteration in 0..repetitions {
    caller.extend_from_slice(&[0x20, 0]);
    if kind == TYPE_F32 { caller.push(0xbe); }
    if kind == TYPE_F64 { caller.push(0xbf); }
    caller.extend_from_slice(&[0x20, 1, if tail { 0x13 } else { 0x11 }, if mismatch { 2 } else { 0 }, 0]);
    if !tail {
        if kind == TYPE_F32 { caller.push(0xbc); }
        if kind == TYPE_F64 { caller.push(0xbd); }
        if mixed { caller.extend_from_slice(&[0x41, 1, 0x6a]); }
        caller.push(if iteration + 1 == repetitions { 0x0f } else { 0x1a });
    }
    }
    caller.push(0x0b);
    let mut code = vec![2];
    write_leb128_u32(&mut code, identity.len() as u32);
    code.extend_from_slice(&identity);
    write_leb128_u32(&mut code, caller.len() as u32);
    code.extend_from_slice(&caller);
    write_section(&mut bytes, SECTION_CODE, &code);
    let mut names = vec![2, 0];
    write_name(&mut names, "numeric_identity");
    names.push(1);
    write_name(&mut names, "typed_memory");
    let mut custom = Vec::new();
    write_name(&mut custom, "name");
    custom.push(1);
    write_leb128_u32(&mut custom, names.len() as u32);
    custom.extend(names);
    write_section(&mut bytes, SECTION_CUSTOM, &custom);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).expect("original numeric indirect fixture must validate");
    let chunks = crate::reader::read_wasm(&bytes).expect("read numeric indirect fixture");
    let rewritten = super::super::write_wasm_with_options(&chunks,
        super::super::WasmWriteOptions { include_name_section: true, ..Default::default() });
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&rewritten).unwrap_or_else(|error| {
            for chunk in &chunks {
                eprintln!("indirect fixture chunk {} signature {:?} sites {:?}", chunk.name, chunk.func_sig, chunk.call_indirect_sigs);
            }
            for payload in wasmparser::Parser::new(0).parse_all(&rewritten) {
                if let wasmparser::Payload::CodeSectionEntry(body) = payload.unwrap() {
                    let mut ops = body.get_operators_reader().unwrap();
                    while !ops.eof() {
                        let at = ops.original_position();
                        eprintln!("{at}: {:?}", ops.read().unwrap());
                    }
                }
            }
            panic!("rewritten numeric indirect fixture must validate: {error}");
        });
    rewritten
}

#[test]
fn repeated_declared_calls_share_typed_scratch_locals() {
    for wide in [false, true] {
        let count = |repetitions| {
            let bytes = numeric_indirect_fixture_repeated(TYPE_I32, wide, true, false, false, repetitions);
            let body = wasmparser::Parser::new(0).parse_all(&bytes).filter_map(|p| match p.unwrap() {
                wasmparser::Payload::CodeSectionEntry(body) => Some(body), _ => None,
            }).nth(2).unwrap();
            let locals: u32 = body.get_locals_reader().unwrap().into_iter()
                .map(|local| local.unwrap().0).sum();
            let calls = body.get_operators_reader().unwrap().into_iter()
                .filter(|op| matches!(op.as_ref().unwrap(), wasmparser::Operator::CallIndirect { .. })).count();
            assert_eq!(calls, repetitions);
            locals
        };
        assert_eq!(count(1), count(256), "scratch frame must not grow per call site");
    }
}

#[test]
fn native_indirect_numeric_calls_preserve_bits_and_table_traps() {
    for wide in [false, true] {
        for kind in [TYPE_I32, TYPE_I64, TYPE_F32, TYPE_F64] {
            let bytes = numeric_indirect_fixture(kind, wide, false, false, false);
            let caller = wasmparser::Parser::new(0).parse_all(&bytes).filter_map(|p| match p.unwrap() {
                wasmparser::Payload::CodeSectionEntry(body) => Some(body), _ => None,
            }).nth(2).expect("native indirect caller");
            let ops: Vec<_> = caller.get_operators_reader().unwrap().into_iter()
                .map(|op| format!("{:?}", op.unwrap())).collect();
            assert!(!ops.iter().any(|op| op.starts_with("Call {") || op.starts_with("LocalSet")));
            let patterns: &[u64] = if matches!(kind, TYPE_I64 | TYPE_F64) {
                &[0, 1 << 63, 0x7ff0000000000001, 0xfff0000000000001, 0x7ff8123456789abc]
            } else { &[0, 1 << 31, 0x7f800001, 0xff800001, 0x7fc12345] };
            for &pattern in patterns {
                let value = if matches!(kind, TYPE_I64 | TYPE_F64) { Value::I64(pattern as i64) } else { Value::I32(pattern as i32) };
                let index = if wide { Value::I64(0) } else { Value::I32(0) };
                assert_eq!(invoke_full_module(&bytes, &[value.clone(), index]).unwrap(), value);
            }
            for index in [1, 2, -1] {
                let index = if wide { Value::I64(index) } else { Value::I32(index as i32) };
                let value = if matches!(kind, TYPE_I64 | TYPE_F64) { Value::I64(0) } else { Value::I32(0) };
                assert!(invoke_full_module(&bytes, &[value, index]).is_err());
            }
        }
    }
}

#[test]
fn indirect_mixed_results_tail_calls_and_signature_traps_remain_standard() {
    for wide in [false, true] {
        let index = if wide { Value::I64(0) } else { Value::I32(0) };
        let mixed = numeric_indirect_fixture(TYPE_I32, wide, true, false, false);
        assert_eq!(invoke_full_module(&mixed, &[Value::I32(41), index.clone()]).unwrap(), Value::I32(42));
        for kind in [TYPE_I32, TYPE_I64] {
            let tail = numeric_indirect_fixture(kind, wide, false, true, false);
            let value = if kind == TYPE_I64 { Value::I64(i64::MIN + 5) } else { Value::I32(i32::MIN + 5) };
            assert_eq!(invoke_full_module(&tail, &[value.clone(), index.clone()]).unwrap(), value);
        }
        let mismatch = numeric_indirect_fixture(TYPE_I64, wide, false, false, true);
        assert!(invoke_full_module(&mismatch, &[Value::I64(5), index]).is_err());
    }
}

#[test]
fn data_segment_modes_indices_widths_and_instantiation_are_preserved() {
    for wide in [false, true] {
        for memory in [0u32, 1] {
            let address = if wide { "i64" } else { "i32" };
            let mut base = typed(&format!("{address},i32,i32"), "i32", wide);
            base.memory_is_64 = if memory == 0 { vec![wide] } else { vec![false, wide] };
            base.data_segments = vec![vec![0xa5, 0x5a], vec![7, 9]];
            base.active_data_segments = vec![vybe_runtime::chunk::ActiveDataSegment {
                memory_index: memory, offset: 3, data_index: 0,
            }];
            let load = |chunk: &mut Chunk| {
                numeric_get(chunk, 0);
                numeric_op(chunk, Op::I32_LOAD8_U);
                write_leb128_u32(&mut chunk.code, 0xc0);
                write_leb128_u32(&mut chunk.code, 0);
                write_leb128_u32(&mut chunk.code, memory);
                numeric_op(chunk, Op::RETURN);
            };
            let mut initial = base.clone();
            load(&mut initial);
            let bytes = module(initial, wide);
            let mut modes = Vec::new();
            for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
                if let wasmparser::Payload::DataSection(section) = payload.unwrap() {
                    for data in section {
                        match data.unwrap().kind {
                            wasmparser::DataKind::Active { memory_index, offset_expr } => {
                                assert_eq!(memory_index, memory);
                                let mut ops = offset_expr.get_operators_reader();
                                match ops.read().unwrap() {
                                    wasmparser::Operator::I64Const { value } if wide => assert_eq!(value, 3),
                                    wasmparser::Operator::I32Const { value } if !wide => assert_eq!(value, 3),
                                    other => panic!("wrong active offset width: {other:?}"),
                                }
                                modes.push(true);
                            }
                            wasmparser::DataKind::Passive => modes.push(false),
                        }
                    }
                }
            }
            assert_eq!(modes, [true, false]);
            let at = |offset| if wide { Value::I64(offset) } else { Value::I32(offset as i32) };
            assert_eq!(invoke_full_module(&bytes, &[at(3), Value::I32(0), Value::I32(0)]).unwrap(), Value::I32(0xa5));
            for segment in [0, 1] {
                let mut init = base.clone();
                for slot in [0, 1, 2] { numeric_get(&mut init, slot); }
                numeric_op(&mut init, Op::MEMORY_INIT);
                write_leb128_u32(&mut init.code, segment);
                write_leb128_u32(&mut init.code, memory);
                load(&mut init);
                let bytes = module(init, wide);
                let copied = invoke_full_module(&bytes, &[at(7), Value::I32(0), Value::I32(1)]);
                if segment == 0 {
                    assert!(copied.is_err(), "active data is unavailable to memory.init after instantiation");
                    assert_eq!(invoke_full_module(&bytes, &[at(7), Value::I32(0), Value::I32(0)]).unwrap(), Value::I32(0));
                    assert!(invoke_full_module(&bytes, &[at(7), Value::I32(1), Value::I32(0)]).is_err());
                } else {
                    assert_eq!(copied.unwrap(), Value::I32(7));
                }
            }
            let mut invalid = base;
            invalid.active_data_segments[0].offset = 65536;
            load(&mut invalid);
            let bytes = module(invalid, wide);
            assert!(invoke_full_module(&bytes, &[at(3), Value::I32(0), Value::I32(0)]).is_err());
        }
    }
}

#[test]
fn original_element_modes_and_bulk_table_indices_survive_emission() {
    for wide in [false, true] {
    for mixed in [false, true] {
    for segment in 0..3u8 {
        let mut input = b"\0asm\x01\0\0\0".to_vec();
        write_section(&mut input, SECTION_TYPE,
            &[2, 0x60, 0, 1, TYPE_I32, 0x60, 3, if wide { TYPE_I64 } else { TYPE_I32 }, TYPE_I32, TYPE_I32, 1, TYPE_I32]);
        write_section(&mut input, SECTION_FUNCTION, &[2, 0, 1]);
        write_section(&mut input, 4, &[1, 0x70, if wide { 5 } else { 1 }, 2, 2]);
        let mut exports = vec![1];
        write_name(&mut exports, "typed_memory");
        exports.extend_from_slice(&[0, 1]);
        write_section(&mut input, SECTION_EXPORT, &exports);
        // The original element order is deliberately passive, active,
        // declarative. Appended writer-owned callable segments must not
        // renumber any table.init/elem.drop operand.
        write_section(&mut input, 9, &[
            3,
            5, 0x70, 1, 0xd2, 0, 0x0b,
            6, 0, if wide { 0x42 } else { 0x41 }, 0, 0x0b, 0x70, 1, 0xd2, 0, 0x0b,
            7, 0x70, 1, 0xd2, 0, 0x0b,
        ]);
        let callee = [0, 0x41, 42, 0x0f, 0x0b];
        let mut caller = vec![0];
        if mixed { caller.extend_from_slice(&[0x02, 0x40, 0x0b]); }
        caller.extend_from_slice(&[0x20, 0, 0x20, 1, 0x20, 2, 0xfc, 12, segment, 0,
            0xfc, 13, segment]);
        let constant = if wide { 0x42 } else { 0x41 };
        caller.extend_from_slice(&[constant, 1, constant, 0, constant, 1, 0xfc, 14, 0, 0,
            constant, 1, 0x11, 0, 0, 0x0f, 0x0b]);
        let mut code = vec![2];
        for body in [&callee[..], &caller[..]] {
            write_leb128_u32(&mut code, body.len() as u32);
            code.extend_from_slice(body);
        }
        write_section(&mut input, SECTION_CODE, &code);
        let mut names = vec![1, 1];
        write_name(&mut names, "typed_memory");
        let mut custom = Vec::new();
        write_name(&mut custom, "name");
        custom.push(1);
        write_leb128_u32(&mut custom, names.len() as u32);
        custom.extend(names);
        write_section(&mut input, SECTION_CUSTOM, &custom);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&input).expect("original bulk table fixture");
        let chunks = crate::reader::read_wasm(&input).unwrap();
        let output = super::super::write_wasm_with_options(&chunks,
            super::super::WasmWriteOptions { include_name_section: true, ..Default::default() });
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&output).expect("rewritten bulk table fixture");
        let mut modes = Vec::new();
        for payload in wasmparser::Parser::new(0).parse_all(&output) {
            if let wasmparser::Payload::ElementSection(elements) = payload.unwrap() {
                for element in elements {
                    modes.push(match element.unwrap().kind {
                        wasmparser::ElementKind::Passive => 0,
                        wasmparser::ElementKind::Active { .. } => 1,
                        wasmparser::ElementKind::Declared => 2,
                    });
                }
            }
        }
        assert_eq!(&modes[..3], &[0, 1, 2]);
        let run = |source, length| invoke_full_module(&output,
            &[if wide { Value::I64(0) } else { Value::I32(0) }, Value::I32(source), Value::I32(length)]);
        assert_eq!(run(0, 0).unwrap(), Value::I32(42));
        if segment == 0 {
            assert_eq!(run(0, 1).unwrap(), Value::I32(42));
        } else {
            assert!(run(0, 1).is_err(), "active/declarative payloads are unavailable after instantiation");
            assert!(run(1, 0).is_err(), "zero-length table.init must still check the source bound");
        }
        assert!(run(1, 1).is_err());
    }
    }
    }
}
