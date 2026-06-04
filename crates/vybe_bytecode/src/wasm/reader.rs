//! WASM binary reader — decodes .wasm files into Chunk arrays.

use super::encoding::*;
use crate::value::Value;
use crate::{Chunk, Op};
use std::sync::Arc;

pub fn read_wasm(data: &[u8]) -> Result<Vec<Chunk>, String> {
    if data.len() < 8 || &data[0..4] != &WASM_MAGIC {
        return Err("Invalid WASM: bad magic".into());
    }
    let mut pos = 8;
    let mut custom_data: Option<Vec<u8>> = None;
    let mut type_section: Vec<u8> = Vec::new();
    let mut import_section: Vec<u8> = Vec::new();
    let mut func_section: Vec<u8> = Vec::new();
    let mut export_section: Vec<u8> = Vec::new();
    let mut code_section: Vec<u8> = Vec::new();

    while pos < data.len() {
        if pos >= data.len() {
            break;
        }
        let section_id = data[pos];
        pos += 1;
        let (size, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let section_end = (pos + size as usize).min(data.len());
        let section_data = data[pos..section_end].to_vec();

        match section_id {
            SECTION_CUSTOM => {
                // Check if it's our "vybe" custom section
                let (nlen, nr) = read_leb128_u32(&section_data);
                if nlen == 4 && section_data.get(nr..nr + 4) == Some(b"vybe") {
                    custom_data = Some(section_data);
                }
            }
            SECTION_TYPE => type_section = section_data,
            SECTION_IMPORT => import_section = section_data,
            SECTION_FUNCTION => func_section = section_data,
            SECTION_EXPORT => export_section = section_data,
            SECTION_CODE => code_section = section_data,
            _ => {} // skip memory, table, global, data, element
        }
        pos = section_end;
    }

    // If we have a vybe custom section, use that for round-trip (our format)
    if let Some(ref cd) = custom_data {
        return decode_vybe_section(cd);
    }

    // Otherwise, decode as standard WASM module
    if code_section.is_empty() {
        return Err("No code section in WASM module".into());
    }
    decode_standard_wasm(
        &type_section,
        &import_section,
        &func_section,
        &export_section,
        &code_section,
    )
}

/// Decode a standard WASM module (e.g. from Rust/C compiler)
fn decode_standard_wasm(
    type_sec: &[u8],
    import_sec: &[u8],
    func_sec: &[u8],
    export_sec: &[u8],
    code_sec: &[u8],
) -> Result<Vec<Chunk>, String> {
    // Parse type section to get function signatures
    let types = parse_type_section(type_sec);
    let func_type_indices = parse_function_section(func_sec);

    // Parse imports
    let imports = parse_import_section(import_sec);
    let import_func_count = imports.iter().filter(|(_, _, kind)| *kind == 0).count();

    // Parse exports to find function names
    let exports = parse_export_section(export_sec);

    // Parse code section
    let mut cpos = 0;
    let (func_count, read) = read_leb128_u32(&code_sec[cpos..]);
    cpos += read;

    let mut chunks = Vec::new();

    // Create a script chunk that calls exported functions
    let mut script = Chunk::new("<script>");
    script.local_count = 0;

    for i in 0..func_count as usize {
        let (body_size, read) = read_leb128_u32(&code_sec[cpos..]);
        cpos += read;
        let _body_start = cpos;
        let body_end = cpos + body_size as usize;

        // Parse locals
        let (local_groups, read) = read_leb128_u32(&code_sec[cpos..]);
        cpos += read;
        let mut local_count: u32 = 0;
        for _ in 0..local_groups {
            let (count, read) = read_leb128_u32(&code_sec[cpos..]);
            cpos += read;
            cpos += 1; // type byte
            local_count += count;
        }

        // Get function name from exports
        let func_idx = import_func_count + i;
        let name = exports
            .iter()
            .find(|(_, idx)| *idx == func_idx)
            .map(|(n, _)| n.clone())
            .unwrap_or_else(|| format!("func_{}", i));

        // Get arity + result arity from the function's type signature.
        let type_idx = func_type_indices.get(i).copied().unwrap_or(func_idx as u32) as usize;
        let (arity, result_arity) = types
            .get(type_idx)
            .map(|(params, results)| (params.len() as u8, (results.len() as u8).max(1)))
            .unwrap_or((0, 1));

        // Translate WASM opcodes to our Chunk format
        let wasm_code = &code_sec[cpos..body_end.saturating_sub(1)]; // -1 for trailing 'end'
        let mut chunk =
            translate_wasm_to_chunk(wasm_code, &name, arity, local_count, import_func_count);
        chunk.result_arity = result_arity;
        chunk.emit_op(Op::RETURN, 0);
        chunks.push(chunk);

        cpos = body_end;
    }

    // Add imports to script chunk
    for (module, name, _) in &imports {
        script.add_import(module, name);
    }

    // Insert script as chunk 0
    chunks.insert(0, script);

    Ok(chunks)
}

/// Translate WASM opcodes to our internal Chunk format.
/// Builds a proper constant pool and adjusts local indices.
fn translate_wasm_to_chunk(
    wasm: &[u8],
    name: &str,
    arity: u8,
    wasm_local_count: u32,
    _import_count: usize,
) -> Chunk {
    let mut chunk = Chunk::new(name);
    chunk.arity = arity;
    chunk.local_count = arity as u16 + wasm_local_count as u16;

    let mut pos = 0;
    let mut label_stack: Vec<()> = Vec::new();

    while pos < wasm.len() {
        let byte = wasm[pos];
        pos += 1;

        match byte {
            0x00 => chunk.emit_op(Op::HALT, 0),
            0x01 => {} // nop

            // block blocktype — forward jump target
            0x02 => {
                let result_count = read_block_result_count(wasm, &mut pos);
                chunk.emit_block_typed(0, result_count);
                label_stack.push(());
            }

            // loop blocktype — backward jump target
            0x03 => {
                let result_count = read_block_result_count(wasm, &mut pos);
                chunk.emit_loop_typed(0, result_count);
                label_stack.push(());
            }

            // if blocktype — conditional block
            0x04 => {
                let result_count = read_block_result_count(wasm, &mut pos);
                if result_count == 0 {
                    chunk.emit_if(0);
                } else {
                    chunk.emit_if_value(0);
                }
                label_stack.push(());
            }

            // else
            0x05 => {
                chunk.emit_else(0);
            }

            // end
            0x0B => {
                let _ = label_stack.pop();
                chunk.emit_end(0);
            }

            // br N — branch to Nth enclosing label
            0x0C => {
                let (depth, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_br(depth, 0);
            }

            // br_if N — conditional branch
            0x0D => {
                let (depth, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_br_if(depth, 0);
            }
            0x0E => {
                // br_table — branch table
                let (count, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                let mut depths = Vec::with_capacity(count as usize);
                for _ in 0..count {
                    let (depth, _) = read_leb128_u32(&wasm[pos..]);
                    skip_leb128(wasm, &mut pos);
                    depths.push(depth);
                }
                let (default_depth, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_br_table(&depths, default_depth, 0);
            }
            0x0F => chunk.emit_op(Op::RETURN, 0),
            0x1A => chunk.emit_op(Op::DROP, 0),
            0x1B => chunk.emit_op(Op::SELECT, 0),

            // call — adjust index (skip imports, offset to our chunk indices)
            0x10 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u8(Op::CALL, idx as u8, 0);
            }

            // local.get — slot 0 is the first argument, matching the VM.
            0x20 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u16(Op::LOCAL_GET, idx as u16, 0);
            }
            // local.set
            0x21 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u16(Op::LOCAL_SET, idx as u16, 0);
            }
            // local.tee
            0x22 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op_u16(Op::LOCAL_SET, idx as u16, 0);
            }

            // i32.const
            0x41 => {
                let (val, read) = read_leb128_i32(&wasm[pos..]);
                pos += read;
                let ci = chunk.add_constant(Value::I32(val));
                chunk.emit_op_u16(Op::CONST, ci, 0);
            }

            // i64.const
            0x42 => {
                let (val, read) = read_leb128_i64(&wasm[pos..]);
                pos += read;
                let ci = chunk.add_constant(Value::I64(val));
                chunk.emit_op_u16(Op::CONST, ci, 0);
            }

            // f64.const
            0x44 => {
                if pos + 8 <= wasm.len() {
                    let val = f64::from_le_bytes([
                        wasm[pos],
                        wasm[pos + 1],
                        wasm[pos + 2],
                        wasm[pos + 3],
                        wasm[pos + 4],
                        wasm[pos + 5],
                        wasm[pos + 6],
                        wasm[pos + 7],
                    ]);
                    pos += 8;
                    let ci = chunk.add_constant(Value::F64(val));
                    chunk.emit_op_u16(Op::CONST, ci, 0);
                }
            }

            // f32.const
            0x43 => {
                if pos + 4 <= wasm.len() {
                    let val = f32::from_le_bytes([
                        wasm[pos],
                        wasm[pos + 1],
                        wasm[pos + 2],
                        wasm[pos + 3],
                    ]);
                    pos += 4;
                    let ci = chunk.add_constant(Value::F64(val as f64));
                    chunk.emit_op_u16(Op::CONST, ci, 0);
                }
            }

            // i32 arithmetic — ALL opcodes
            0x67 => chunk.emit_op(Op::I32_CLZ, 0),
            0x68 => chunk.emit_op(Op::I32_CTZ, 0),
            0x69 => chunk.emit_op(Op::I32_POPCNT, 0),
            0x6A => chunk.emit_op(Op::I32_ADD, 0),
            0x6B => chunk.emit_op(Op::I32_SUB, 0),
            0x6C => chunk.emit_op(Op::I32_MUL, 0),
            0x6D => chunk.emit_op(Op::I32_DIV_S, 0),
            0x6E => chunk.emit_op(Op::I32_DIV_U, 0),
            0x6F => chunk.emit_op(Op::I32_REM_S, 0),
            0x70 => chunk.emit_op(Op::I32_REM_U, 0),
            0x71 => chunk.emit_op(Op::I32_AND, 0),
            0x72 => chunk.emit_op(Op::I32_OR, 0),
            0x73 => chunk.emit_op(Op::I32_XOR, 0),
            0x74 => chunk.emit_op(Op::I32_SHL, 0),
            0x75 => chunk.emit_op(Op::I32_SHR_S, 0),
            0x76 => chunk.emit_op(Op::I32_SHR_U, 0),
            0x77 => chunk.emit_op(Op::I32_ROTL, 0),
            0x78 => chunk.emit_op(Op::I32_ROTR, 0),

            // i64 arithmetic — ALL opcodes
            0x79 => chunk.emit_op(Op::I64_CLZ, 0),
            0x7A => chunk.emit_op(Op::I64_CTZ, 0),
            0x7B => chunk.emit_op(Op::I64_POPCNT, 0),
            0x7C => chunk.emit_op(Op::I64_ADD, 0),
            0x7D => chunk.emit_op(Op::I64_SUB, 0),
            0x7E => chunk.emit_op(Op::I64_MUL, 0),
            0x7F => chunk.emit_op(Op::I64_DIV_S, 0),
            0x80 => chunk.emit_op(Op::I64_DIV_U, 0),
            0x81 => chunk.emit_op(Op::I64_REM_S, 0),
            0x82 => chunk.emit_op(Op::I64_REM_U, 0),
            0x83 => chunk.emit_op(Op::I64_AND, 0),
            0x84 => chunk.emit_op(Op::I64_OR, 0),
            0x85 => chunk.emit_op(Op::I64_XOR, 0),
            0x86 => chunk.emit_op(Op::I64_SHL, 0),
            0x87 => chunk.emit_op(Op::I64_SHR_S, 0),
            0x88 => chunk.emit_op(Op::I64_SHR_U, 0),
            0x89 => chunk.emit_op(Op::I64_ROTL, 0),
            0x8A => chunk.emit_op(Op::I64_ROTR, 0),

            // i64 comparison
            0x50 => chunk.emit_op(Op::I64_EQZ, 0),
            0x51 => chunk.emit_op(Op::I64_EQ, 0),
            0x52 => chunk.emit_op(Op::I64_NE, 0),
            0x53 => chunk.emit_op(Op::I64_LT_S, 0),
            0x54 => chunk.emit_op(Op::I64_LT_U, 0),
            0x55 => chunk.emit_op(Op::I64_GT_S, 0),
            0x56 => chunk.emit_op(Op::I64_GT_U, 0),
            0x57 => chunk.emit_op(Op::I64_LE_S, 0),
            0x58 => chunk.emit_op(Op::I64_LE_U, 0),
            0x59 => chunk.emit_op(Op::I64_GE_S, 0),
            0x5A => chunk.emit_op(Op::I64_GE_U, 0),

            // i32 comparison
            0x45 => chunk.emit_op(Op::I32_EQZ, 0),
            0x46 => chunk.emit_op(Op::I32_EQ, 0),
            0x47 => chunk.emit_op(Op::I32_NE, 0),
            0x48 => chunk.emit_op(Op::I32_LT_S, 0),
            0x49 => chunk.emit_op(Op::I32_LT_U, 0),
            0x4A => chunk.emit_op(Op::I32_GT_S, 0),
            0x4B => chunk.emit_op(Op::I32_GT_U, 0),
            0x4C => chunk.emit_op(Op::I32_LE_S, 0),
            0x4D => chunk.emit_op(Op::I32_LE_U, 0),
            0x4E => chunk.emit_op(Op::I32_GE_S, 0),
            0x4F => chunk.emit_op(Op::I32_GE_U, 0),

            // f64 arithmetic — ALL opcodes
            0xA0 => chunk.emit_op(Op::F64_ADD, 0),
            0xA1 => chunk.emit_op(Op::F64_SUB, 0),
            0xA2 => chunk.emit_op(Op::F64_MUL, 0),
            0xA3 => chunk.emit_op(Op::F64_DIV, 0),
            0xA4 => chunk.emit_op(Op::F64_MIN, 0),
            0xA5 => chunk.emit_op(Op::F64_MAX, 0),
            0xA6 => chunk.emit_op(Op::F64_COPYSIGN, 0),

            // f32 comparison (mapped to f64 ops — Vybe uses f64 internally)
            0x5B => chunk.emit_op(Op::F64_EQ, 0),
            0x5C => chunk.emit_op(Op::F64_NE, 0),
            0x5D => chunk.emit_op(Op::F64_LT, 0),
            0x5E => chunk.emit_op(Op::F64_GT, 0),
            0x5F => chunk.emit_op(Op::F64_LE, 0),
            0x60 => chunk.emit_op(Op::F64_GE, 0),

            // f64 comparison
            0x61 => chunk.emit_op(Op::F64_EQ, 0),
            0x62 => chunk.emit_op(Op::F64_NE, 0),
            0x63 => chunk.emit_op(Op::F64_LT, 0),
            0x64 => chunk.emit_op(Op::F64_GT, 0),
            0x65 => chunk.emit_op(Op::F64_LE, 0),
            0x66 => chunk.emit_op(Op::F64_GE, 0),

            // Memory — ALL load/store opcodes
            0x28 => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I32_LOAD, 0);
            }
            0x29 => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I64_LOAD, 0);
            }
            0x2A => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::F32_LOAD, 0);
            }
            0x2B => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::F64_LOAD, 0);
            }
            0x2C => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I32_LOAD8_S, 0);
            }
            0x2D => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I32_LOAD8_U, 0);
            }
            0x2E => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I32_LOAD16_S, 0);
            }
            0x2F => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I32_LOAD16_U, 0);
            }
            0x30 => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I64_LOAD8_S, 0);
            }
            0x31 => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I64_LOAD8_U, 0);
            }
            0x32 => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I64_LOAD16_S, 0);
            }
            0x33 => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I64_LOAD16_U, 0);
            }
            0x34 => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I64_LOAD32_S, 0);
            }
            0x35 => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I64_LOAD32_U, 0);
            }
            0x36 => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I32_STORE, 0);
            }
            0x37 => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I64_STORE, 0);
            }
            0x38 => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::F32_STORE, 0);
            }
            0x39 => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::F64_STORE, 0);
            }
            0x3A => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I32_STORE8, 0);
            }
            0x3B => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I32_STORE16, 0);
            }
            0x3C => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I64_STORE8, 0);
            }
            0x3D => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I64_STORE16, 0);
            }
            0x3E => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::I64_STORE32, 0);
            }
            0x3F => {
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::MEMORY_SIZE, 0);
            }
            0x40 => {
                skip_leb128(wasm, &mut pos);
                chunk.emit_op(Op::MEMORY_GROW, 0);
            }

            // f32 arithmetic — ALL opcodes
            0x8B => chunk.emit_op(Op::F32_ABS, 0),
            0x8C => chunk.emit_op(Op::F32_NEG, 0),
            0x8D => chunk.emit_op(Op::F32_CEIL, 0),
            0x8E => chunk.emit_op(Op::F32_FLOOR, 0),
            0x8F => chunk.emit_op(Op::F32_TRUNC, 0),
            0x90 => chunk.emit_op(Op::F32_NEAREST, 0),
            0x91 => chunk.emit_op(Op::F32_SQRT, 0),
            0x92 => chunk.emit_op(Op::F64_ADD, 0), // f32.add (promoted)
            0x93 => chunk.emit_op(Op::F64_SUB, 0), // f32.sub (promoted)
            0x94 => chunk.emit_op(Op::F64_MUL, 0), // f32.mul (promoted)
            0x95 => chunk.emit_op(Op::F64_DIV, 0), // f32.div (promoted)
            0x96 => chunk.emit_op(Op::F32_MIN, 0),
            0x97 => chunk.emit_op(Op::F32_MAX, 0),
            0x98 => chunk.emit_op(Op::F32_COPYSIGN, 0),

            // f64 extra ops — ALL opcodes
            0x99 => chunk.emit_op(Op::F64_ABS, 0),
            0x9A => chunk.emit_op(Op::F64_NEG, 0),
            0x9B => chunk.emit_op(Op::F64_CEIL, 0),
            0x9C => chunk.emit_op(Op::F64_FLOOR, 0),
            0x9D => chunk.emit_op(Op::F64_TRUNC, 0),
            0x9E => chunk.emit_op(Op::F64_NEAREST, 0),
            0x9F => chunk.emit_op(Op::F64_SQRT, 0),

            // Conversions (WASM spec §5.3-binary.instructions 0xA7–0xBF)
            0xA7 => chunk.emit_op(Op::I32_WRAP_I64, 0), // i32.wrap_i64
            0xA8 => chunk.emit_op(Op::I32_TRUNC_F32_S, 0), // i32.trunc_f32_s
            0xA9 => chunk.emit_op(Op::I32_TRUNC_F32_U, 0), // i32.trunc_f32_u
            0xAA => chunk.emit_op(Op::I32_FROM_F64, 0), // i32.trunc_f64_s
            0xAB => chunk.emit_op(Op::I32_TRUNC_F64_U, 0), // i32.trunc_f64_u
            0xAC => chunk.emit_op(Op::I64_EXTEND_I32_S, 0), // i64.extend_i32_s
            0xAD => chunk.emit_op(Op::I64_EXTEND_I32_U, 0), // i64.extend_i32_u
            0xAE => chunk.emit_op(Op::I64_TRUNC_F32_S, 0), // i64.trunc_f32_s
            0xAF => chunk.emit_op(Op::I64_TRUNC_F32_U, 0), // i64.trunc_f32_u
            0xB0 => chunk.emit_op(Op::I64_TRUNC_F64_S, 0), // i64.trunc_f64_s
            0xB1 => chunk.emit_op(Op::I64_TRUNC_F64_U, 0), // i64.trunc_f64_u
            0xB2 => chunk.emit_op(Op::F32_CONVERT_I32_S, 0), // f32.convert_i32_s
            0xB3 => chunk.emit_op(Op::F32_CONVERT_I32_U, 0), // f32.convert_i32_u
            0xB4 => chunk.emit_op(Op::F32_CONVERT_I64_S, 0), // f32.convert_i64_s
            0xB5 => chunk.emit_op(Op::F32_CONVERT_I64_U, 0), // f32.convert_i64_u
            0xB6 => chunk.emit_op(Op::F32_DEMOTE_F64, 0), // f32.demote_f64
            0xB7 => chunk.emit_op(Op::F64_FROM_I32, 0), // f64.convert_i32_s
            0xB8 => chunk.emit_op(Op::F64_CONVERT_I32_U, 0), // f64.convert_i32_u
            0xB9 => chunk.emit_op(Op::F64_CONVERT_I64_S, 0), // f64.convert_i64_s
            0xBA => chunk.emit_op(Op::F64_CONVERT_I64_U, 0), // f64.convert_i64_u
            0xBB => chunk.emit_op(Op::F64_PROMOTE_F32, 0), // f64.promote_f32
            0xBC => chunk.emit_op(Op::I32_REINTERPRET_F32, 0), // i32.reinterpret_f32
            0xBD => chunk.emit_op(Op::I64_REINTERPRET_F64, 0), // i64.reinterpret_f64
            0xBE => chunk.emit_op(Op::F32_REINTERPRET_I32, 0), // f32.reinterpret_i32
            0xBF => chunk.emit_op(Op::F64_REINTERPRET_I64, 0), // f64.reinterpret_i64

            // Sign extension
            0xC0 => chunk.emit_op(Op::I32_EXTEND8_S, 0),
            0xC1 => chunk.emit_op(Op::I32_EXTEND16_S, 0),
            0xC2 => chunk.emit_op(Op::I64_EXTEND8_S, 0),
            0xC3 => chunk.emit_op(Op::I64_EXTEND16_S, 0),
            0xC4 => chunk.emit_op(Op::I64_EXTEND32_S, 0),

            // global.get/set — WASM globals mapped to global_get/set with index as name
            0x23 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                let name = format!("__wasm_global_{}", idx);
                let ci = chunk.add_constant(Value::String(Arc::from(name.as_str())));
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
            }
            0x24 => {
                let (idx, _) = read_leb128_u32(&wasm[pos..]);
                skip_leb128(wasm, &mut pos);
                let name = format!("__wasm_global_{}", idx);
                let ci = chunk.add_constant(Value::String(Arc::from(name.as_str())));
                chunk.emit_op_u16(Op::GLOBAL_SET, ci, 0);
            }

            // call_indirect
            0x11 => {
                skip_leb128(wasm, &mut pos);
                skip_leb128(wasm, &mut pos);
            }

            // 0xFC prefix — nontrapping-float-to-int (0x00–0x07) + bulk-memory/table ops
            0xFC => {
                let (sub, read) = read_leb128_u32(&wasm[pos..]);
                pos += read;
                match sub {
                    0x00 => chunk.emit_op(Op::I32_TRUNC_SAT_F32_S, 0),
                    0x01 => chunk.emit_op(Op::I32_TRUNC_SAT_F32_U, 0),
                    0x02 => chunk.emit_op(Op::I32_TRUNC_SAT_F64_S, 0),
                    0x03 => chunk.emit_op(Op::I32_TRUNC_SAT_F64_U, 0),
                    0x04 => chunk.emit_op(Op::I64_TRUNC_SAT_F32_S, 0),
                    0x05 => chunk.emit_op(Op::I64_TRUNC_SAT_F32_U, 0),
                    0x06 => chunk.emit_op(Op::I64_TRUNC_SAT_F64_S, 0),
                    0x07 => chunk.emit_op(Op::I64_TRUNC_SAT_F64_U, 0),
                    // bulk-memory / table ops have immediates — skip them
                    _ => {
                        skip_leb128(wasm, &mut pos);
                    }
                }
            }

            // GC proposal prefix.
            0xFB => {
                let (sub, read) = read_leb128_u32(&wasm[pos..]);
                pos += read;
                emit_gc_prefixed(&mut chunk, sub, wasm, &mut pos);
            }

            // SIMD and relaxed-SIMD proposal prefix.
            0xFD => {
                let (sub, read) = read_leb128_u32(&wasm[pos..]);
                pos += read;
                emit_simd_prefixed(&mut chunk, sub, wasm, &mut pos);
            }

            // Threads/atomics proposal prefix.
            0xFE => {
                let (sub, read) = read_leb128_u32(&wasm[pos..]);
                pos += read;
                emit_threads_prefixed(&mut chunk, sub, wasm, &mut pos);
            }

            // Unknown — skip
            _ => {}
        }
    }

    chunk
}

fn emit_gc_prefixed(chunk: &mut Chunk, sub: u32, wasm: &[u8], pos: &mut usize) {
    let Some(op) = u8::try_from(sub).ok().and_then(|s| Op::decode(0xFB, s)) else {
        return;
    };
    match op {
        _ if op == Op::STRUCT_NEW
            || op == Op::STRUCT_NEW_DEFAULT
            || op == Op::STRUCT_GET
            || op == Op::STRUCT_GET_S
            || op == Op::STRUCT_GET_U
            || op == Op::STRUCT_SET
            || op == Op::ARRAY_NEW
            || op == Op::ARRAY_NEW_DEFAULT
            || op == Op::ARRAY_GET
            || op == Op::ARRAY_GET_S
            || op == Op::ARRAY_GET_U
            || op == Op::ARRAY_SET
            || op == Op::ARRAY_LENGTH
            || op == Op::ARRAY_FILL =>
        {
            let (idx, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            match op.operand_format() {
                crate::opcode::OperandFormat::U16 => chunk.emit_op_u16(op, idx as u16, 0),
                _ => chunk.emit_op(op, 0),
            }
        }
        _ if op == Op::ARRAY_NEW_FIXED
            || op == Op::ARRAY_NEW_DATA
            || op == Op::ARRAY_NEW_ELEM
            || op == Op::ARRAY_INIT_DATA
            || op == Op::ARRAY_INIT_ELEM =>
        {
            let (_type_idx, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            let (extra, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            chunk.emit_op_u16(op, extra as u16, 0);
        }
        _ if op == Op::ARRAY_COPY => {
            skip_leb128(wasm, pos);
            skip_leb128(wasm, pos);
            chunk.emit_op(op, 0);
        }
        _ if op == Op::REF_TEST
            || op == Op::REF_TEST_NULL
            || op == Op::REF_CAST
            || op == Op::REF_CAST_NULL =>
        {
            skip_heaptype(wasm, pos);
            let idx = chunk.add_constant(Value::String(Arc::from("__wasm_heaptype")));
            chunk.emit_op_u16(op, idx, 0);
        }
        _ if op == Op::BR_ON_CAST || op == Op::BR_ON_CAST_FAIL => {
            skip_leb128(wasm, pos); // flags
            let (depth, read) = read_leb128_u32(&wasm[*pos..]);
            *pos += read;
            skip_heaptype(wasm, pos);
            skip_heaptype(wasm, pos);
            let idx = chunk.add_constant(Value::String(Arc::from("__wasm_heaptype")));
            chunk.emit_op_u16(op, idx, 0);
            chunk.emit(depth as u8, 0);
        }
        _ => chunk.emit_op(op, 0),
    }
}

fn emit_simd_prefixed(chunk: &mut Chunk, sub: u32, wasm: &[u8], pos: &mut usize) {
    if (0x100..=0x113).contains(&sub) {
        let relaxed_sub = (sub - 0x100) as u8;
        if let Some(op) = Op::decode(0xDD, relaxed_sub) {
            chunk.emit_op(op, 0);
        }
        return;
    }

    let Some(op) = u8::try_from(sub).ok().and_then(|s| Op::decode(0xFD, s)) else {
        return;
    };
    match op {
        _ if op == Op::V128_LOAD || op == Op::V128_STORE => {
            skip_memarg(wasm, pos);
            chunk.emit_op(op, 0);
        }
        _ if op == Op::V128_CONST => {
            chunk.emit_op(op, 0);
            for _ in 0..16 {
                let b = wasm.get(*pos).copied().unwrap_or(0);
                *pos += 1;
                chunk.emit(b, 0);
            }
        }
        _ if op == Op::I8X16_SHUFFLE => {
            chunk.emit_op(op, 0);
            for _ in 0..16 {
                let b = wasm.get(*pos).copied().unwrap_or(0);
                *pos += 1;
                chunk.emit(b, 0);
            }
        }
        _ if op.operand_format() == crate::opcode::OperandFormat::U8 => {
            let lane = wasm.get(*pos).copied().unwrap_or(0);
            *pos += 1;
            chunk.emit_op_u8(op, lane, 0);
        }
        _ => chunk.emit_op(op, 0),
    }
}

fn emit_threads_prefixed(chunk: &mut Chunk, sub: u32, wasm: &[u8], pos: &mut usize) {
    let Some(op) = u8::try_from(sub).ok().and_then(|s| Op::decode(0xFE, s)) else {
        return;
    };
    match op {
        _ if op == Op::ATOMIC_FENCE => {
            *pos = (*pos).saturating_add(1).min(wasm.len());
            chunk.emit_op(op, 0);
        }
        _ if op == Op::THREAD_SPAWN || op == Op::THREAD_JOIN => {
            chunk.emit_op(op, 0);
        }
        _ => {
            skip_memarg(wasm, pos);
            chunk.emit_op(op, 0);
        }
    }
}

fn read_leb128_i64(data: &[u8]) -> (i64, usize) {
    let mut result = 0i64;
    let mut shift = 0;
    let mut pos = 0;
    loop {
        if pos >= data.len() {
            break;
        }
        let byte = data[pos];
        pos += 1;
        result |= ((byte & 0x7F) as i64) << shift;
        shift += 7;
        if byte & 0x80 == 0 {
            if shift < 64 && (byte & 0x40 != 0) {
                result |= !0i64 << shift;
            }
            break;
        }
    }
    (result, pos)
}

fn parse_type_section(data: &[u8]) -> Vec<(Vec<u8>, Vec<u8>)> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut types = Vec::new();
    for _ in 0..count {
        if pos >= data.len() || data[pos] != TYPE_FUNC {
            pos += 1;
            continue;
        }
        pos += 1; // skip 0x60
        let (param_count, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let params: Vec<u8> = data[pos..pos + param_count as usize].to_vec();
        pos += param_count as usize;
        let (result_count, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let results: Vec<u8> = data[pos..pos + result_count as usize].to_vec();
        pos += result_count as usize;
        types.push((params, results));
    }
    types
}

fn parse_function_section(data: &[u8]) -> Vec<u32> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut funcs = Vec::with_capacity(count as usize);
    for _ in 0..count {
        let (type_idx, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        funcs.push(type_idx);
    }
    funcs
}

fn parse_import_section(data: &[u8]) -> Vec<(String, String, u8)> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut imports = Vec::new();
    for _ in 0..count {
        let (mlen, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let module = std::str::from_utf8(&data[pos..pos + mlen as usize])
            .unwrap_or("")
            .to_string();
        pos += mlen as usize;
        let (nlen, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let name = std::str::from_utf8(&data[pos..pos + nlen as usize])
            .unwrap_or("")
            .to_string();
        pos += nlen as usize;
        let kind = data[pos];
        pos += 1;
        skip_leb128(&data, &mut pos); // type index or other descriptor
        imports.push((module, name, kind));
    }
    imports
}

fn parse_export_section(data: &[u8]) -> Vec<(String, usize)> {
    if data.is_empty() {
        return vec![];
    }
    let mut pos = 0;
    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let mut exports = Vec::new();
    for _ in 0..count {
        let (nlen, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let name = std::str::from_utf8(&data[pos..pos + nlen as usize])
            .unwrap_or("")
            .to_string();
        pos += nlen as usize;
        let kind = data[pos];
        pos += 1;
        let (idx, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        if kind == 0 {
            // function export
            exports.push((name, idx as usize));
        }
    }
    exports
}

fn skip_leb128(data: &[u8], pos: &mut usize) {
    while *pos < data.len() {
        let byte = data[*pos];
        *pos += 1;
        if byte & 0x80 == 0 {
            break;
        }
    }
}

fn skip_memarg(data: &[u8], pos: &mut usize) {
    skip_leb128(data, pos); // align
    skip_leb128(data, pos); // offset
}

fn skip_heaptype(data: &[u8], pos: &mut usize) {
    skip_leb128(data, pos);
}

fn read_block_result_count(data: &[u8], pos: &mut usize) -> u8 {
    let first = data.get(*pos).copied().unwrap_or(0x40);
    skip_leb128(data, pos);
    if first == 0x40 { 0 } else { 1 }
}

fn read_leb128_i32(data: &[u8]) -> (i32, usize) {
    let mut result = 0i32;
    let mut shift = 0;
    let mut pos = 0;
    loop {
        if pos >= data.len() {
            break;
        }
        let byte = data[pos];
        pos += 1;
        result |= ((byte & 0x7F) as i32) << shift;
        shift += 7;
        if byte & 0x80 == 0 {
            if shift < 32 && (byte & 0x40 != 0) {
                result |= !0 << shift;
            }
            break;
        }
    }
    (result, pos)
}

fn decode_vybe_section(data: &[u8]) -> Result<Vec<Chunk>, String> {
    let mut pos = 0;
    let (name_len, read) = read_leb128_u32(&data[pos..]);
    pos += read;
    let _name = &data[pos..pos + name_len as usize];
    pos += name_len as usize;

    // Version byte
    let _version = data[pos];
    pos += 1;

    let (count, read) = read_leb128_u32(&data[pos..]);
    pos += read;

    let mut chunks = Vec::new();
    for _ in 0..count {
        let (nlen, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let name = std::str::from_utf8(&data[pos..pos + nlen as usize])
            .unwrap_or("")
            .to_string();
        pos += nlen as usize;
        let arity = data[pos];
        pos += 1;
        let (lc, read) = read_leb128_u32(&data[pos..]);
        pos += read;

        // Constants
        let (cc, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let mut constants = Vec::new();
        for _ in 0..cc {
            constants.push(decode_value(data, &mut pos));
        }

        // Imports
        let (ic, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let mut imports = Vec::new();
        for _ in 0..ic {
            let (mlen, read) = read_leb128_u32(&data[pos..]);
            pos += read;
            let module = std::str::from_utf8(&data[pos..pos + mlen as usize])
                .unwrap_or("")
                .to_string();
            pos += mlen as usize;
            let (nlen, read) = read_leb128_u32(&data[pos..]);
            pos += read;
            let iname = std::str::from_utf8(&data[pos..pos + nlen as usize])
                .unwrap_or("")
                .to_string();
            pos += nlen as usize;
            imports.push(crate::chunk::Import {
                module,
                name: iname,
            });
        }

        // Bytecode
        let (code_len, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let code = data[pos..pos + code_len as usize].to_vec();
        pos += code_len as usize;

        // Line info
        let (line_count, read) = read_leb128_u32(&data[pos..]);
        pos += read;
        let mut lines = Vec::with_capacity(line_count as usize);
        for _ in 0..line_count {
            let (line, read) = read_leb128_u32(&data[pos..]);
            pos += read;
            lines.push(line);
        }

        let mut chunk = Chunk::new(&name);
        chunk.arity = arity;
        chunk.local_count = lc as u16;
        chunk.constants = constants;
        chunk.imports = imports;
        chunk.code = code;
        chunk.lines = lines;
        chunks.push(chunk);
    }
    Ok(chunks)
}

// ============================================================
// Helpers
