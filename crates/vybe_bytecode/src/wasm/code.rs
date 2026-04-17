//! WASM code section encoding.
//! Translates internal bytecode to WASM binary format.

use super::encoding::*;
use super::types::WasmTypeContext;
use super::sections::{emit_rt_call, emit_box_i32, emit_box_f64, emit_unbox_f64, emit_unbox_i32};
use crate::{Chunk, Op};
use crate::value::Value;
use crate::opcode::OperandFormat;

pub fn encode_code_section(chunks: &[Chunk], rt_imports: &[(&str, &str)], type_ctx: &WasmTypeContext) -> Vec<u8> {
    let host_import_count = chunks.first().map(|c| c.imports.len()).unwrap_or(0);

    // Build import name → function index map (module:name for uniqueness)
    let mut rt_idx: std::collections::HashMap<(&str, &str), usize> = std::collections::HashMap::new();
    for (i, &(module, name)) in rt_imports.iter().enumerate() {
        rt_idx.insert((module, name), host_import_count + i);
    }

    let mut out = Vec::new();
    write_leb128_u32(&mut out, chunks.len() as u32);

    for chunk in chunks {
        let mut body = Vec::new();

        // Locals — all externref (universal value representation)
        if chunk.local_count > 0 {
            write_leb128_u32(&mut body, 1);
            write_leb128_u32(&mut body, chunk.local_count as u32);
            body.push(TYPE_EXTERNREF);
        } else {
            write_leb128_u32(&mut body, 0);
        }

        // Translate opcodes
        let mut ip = 0;
        while ip < chunk.code.len() {
            if ip + 1 >= chunk.code.len() { break; }
            let op = match Op::decode(chunk.code[ip], chunk.code[ip + 1]) {
                Some(op) => op,
                None => { ip += 2; continue; }
            };
            ip += 2;

            if op.prefix() == 0x00 && !op.is_vm_internal() {
                // ── Core WASM MVP ──
                emit_core_op(&mut body, op, chunk, &mut ip, &rt_idx);
            } else if op.prefix() == 0xFB {
                // ── GC ops → real WASM GC binary encoding ──
                emit_gc_op(&mut body, op, chunk, &mut ip, &rt_idx, type_ctx);
            } else if op.prefix() >= 0xFC && op.prefix() <= 0xFE {
                // ── Other prefixed WASM ops ──
                body.push(op.prefix());
                write_leb128_u32(&mut body, op.sub() as u32);
                ip += op.operand_format().fixed_size();
            } else {
                // ── VM-internal ops (0xFF) ──
                emit_vm_internal_op(&mut body, op, chunk, &mut ip, &rt_idx);
            }
        }
        body.push(0x0B); // end

        write_leb128_u32(&mut out, body.len() as u32);
        out.extend_from_slice(&body);
    }
    out
}

/// Emit a core WASM MVP opcode (prefix 0x00).
fn emit_core_op(body: &mut Vec<u8>, op: Op, chunk: &Chunk, ip: &mut usize, rt_idx: &std::collections::HashMap<(&str, &str), usize>) {
    match op {
        _ if op == Op::LOCAL_GET => { body.push(op.sub()); write_leb128_u32(body, read_u16(&chunk.code, ip) as u32); }
        _ if op == Op::LOCAL_SET => { body.push(0x22); write_leb128_u32(body, read_u16(&chunk.code, ip) as u32); } // local.tee
        _ if op == Op::CALL => { body.push(op.sub()); let argc = chunk.code[*ip]; *ip += 1; write_leb128_u32(body, argc as u32); }
        _ if op == Op::CALL_REF => { let _argc = chunk.code[*ip]; *ip += 1; body.push(op.sub()); write_leb128_u32(body, 0); }
        _ if op == Op::BR => { let _ = read_i16(&chunk.code, ip); body.push(op.sub()); write_leb128_u32(body, 0); }
        _ if op == Op::BR_IF_TRUE => { let _ = read_i16(&chunk.code, ip); body.push(op.sub()); write_leb128_u32(body, 0); }
        _ if op == Op::BLOCK => { let _ = read_u16(&chunk.code, ip); body.push(op.sub()); body.push(TYPE_VOID); }
        _ if op == Op::LOOP => { let _ = read_u16(&chunk.code, ip); body.push(op.sub()); body.push(TYPE_VOID); }
        _ if op == Op::MEMORY_SIZE || op == Op::MEMORY_GROW => { body.push(op.sub()); body.push(0x00); }
        // Memory load/store with alignment + offset
        _ if op == Op::I32_LOAD || op == Op::F32_LOAD => { body.push(op.sub()); body.push(0x02); body.push(0x00); }
        _ if op == Op::I64_LOAD || op == Op::F64_LOAD => { body.push(op.sub()); body.push(0x03); body.push(0x00); }
        _ if op == Op::I32_LOAD8_S || op == Op::I32_LOAD8_U || op == Op::I64_LOAD8_S || op == Op::I64_LOAD8_U => { body.push(op.sub()); body.push(0x00); body.push(0x00); }
        _ if op == Op::I32_LOAD16_S || op == Op::I32_LOAD16_U || op == Op::I64_LOAD16_S || op == Op::I64_LOAD16_U => { body.push(op.sub()); body.push(0x01); body.push(0x00); }
        _ if op == Op::I64_LOAD32_S || op == Op::I64_LOAD32_U => { body.push(op.sub()); body.push(0x02); body.push(0x00); }
        _ if op == Op::I32_STORE || op == Op::F32_STORE => { body.push(op.sub()); body.push(0x02); body.push(0x00); }
        _ if op == Op::I64_STORE || op == Op::F64_STORE => { body.push(op.sub()); body.push(0x03); body.push(0x00); }
        _ if op == Op::I32_STORE8 || op == Op::I64_STORE8 => { body.push(op.sub()); body.push(0x00); body.push(0x00); }
        _ if op == Op::I32_STORE16 || op == Op::I64_STORE16 => { body.push(op.sub()); body.push(0x01); body.push(0x00); }
        _ if op == Op::I64_STORE32 => { body.push(op.sub()); body.push(0x02); body.push(0x00); }
        // WASM global.get/set need global index — our operand is a string name const idx.
        // For .wasm output, emit as nop + skip operand (proper global section needed for real support).
        _ if op == Op::GLOBAL_GET || op == Op::GLOBAL_SET => {
            let _ = read_u16(&chunk.code, ip); // skip string name const idx
            body.push(0x01); // nop — TODO: emit real global.get/set with global section index
        }
        _ if op == Op::HALT => body.push(0x00), // unreachable
        // Numeric binary ops: unbox both operands, operate, re-box
        _ if op == Op::F64_ADD || op == Op::F64_SUB || op == Op::F64_MUL || op == Op::F64_DIV
          || op == Op::F64_MIN || op == Op::F64_MAX => {
            // Stack: [externref_a, externref_b] → unbox → f64 op → box
            // Need to unbox b first (TOS), save, unbox a, push b back, operate
            // Simpler: just emit the raw op and hope for type coercion
            // Actually: for correct .wasm, we need proper boxing.
            // For now, emit as runtime call (dyn_add handles the types)
            emit_rt_call(body, rt_idx, "dyn_add"); // placeholder
        }
        _ if op == Op::I32_ADD || op == Op::I32_SUB || op == Op::I32_MUL
          || op == Op::I32_DIV_S || op == Op::I32_REM_S => {
            emit_rt_call(body, rt_idx, "dyn_add"); // placeholder
        }
        _ => {
            // Other core ops: emit WASM byte directly
            body.push(op.sub());
            *ip += op.operand_format().fixed_size();
        }
    }
}

/// Emit a GC op (prefix 0xFB) — emit real WASM GC binary encoding with type indices.
fn emit_gc_op(body: &mut Vec<u8>, op: Op, chunk: &Chunk, ip: &mut usize, _rt_idx: &std::collections::HashMap<(&str, &str), usize>, type_ctx: &WasmTypeContext) {
    body.push(0xFB);
    write_leb128_u32(body, op.sub() as u32);

    match op {
        _ if op == Op::STRUCT_NEW => {
            let _prop_count = read_u16(&chunk.code, ip);
            // struct.new $typeidx — emit type index 0 (first struct type)
            // TODO: resolve actual type from compile-time context
            write_leb128_u32(body, 0);
        }
        _ if op == Op::STRUCT_GET => {
            let _field_name_idx = read_u16(&chunk.code, ip);
            // struct.get $typeidx $fieldidx
            write_leb128_u32(body, 0); // type index
            write_leb128_u32(body, 0); // field index
        }
        _ if op == Op::STRUCT_SET => {
            let _field_name_idx = read_u16(&chunk.code, ip);
            // struct.set $typeidx $fieldidx
            write_leb128_u32(body, 0);
            write_leb128_u32(body, 0);
        }
        _ if op == Op::ARRAY_NEW => {
            let elem_count = read_u16(&chunk.code, ip);
            // array.new_fixed $typeidx $count
            write_leb128_u32(body, type_ctx.array_type_idx);
            write_leb128_u32(body, elem_count as u32);
        }
        _ if op == Op::ARRAY_GET || op == Op::ARRAY_SET => {
            // array.get/set $typeidx
            write_leb128_u32(body, type_ctx.array_type_idx);
        }
        _ if op == Op::ARRAY_LENGTH => {
            // array.len — no operands in WASM GC (type is inferred from stack)
        }
        _ if op == Op::ARRAY_FILL => {
            write_leb128_u32(body, type_ctx.array_type_idx);
        }
        _ if op == Op::ARRAY_COPY => {
            write_leb128_u32(body, type_ctx.array_type_idx); // dst type
            write_leb128_u32(body, type_ctx.array_type_idx); // src type
        }
        _ if op == Op::ARRAY_NEW_DEFAULT => {
            write_leb128_u32(body, type_ctx.array_type_idx);
        }
        _ if op == Op::REF_TEST || op == Op::REF_CAST => {
            *ip += op.operand_format().fixed_size();
            // TODO: emit proper heap type reference
        }
        _ if op == Op::BR_ON_CAST || op == Op::BR_ON_CAST_FAIL => {
            *ip += op.operand_format().fixed_size();
        }
        _ => {
            // No-operand GC ops (i31.new, i31.get_s, i31.get_u)
            *ip += op.operand_format().fixed_size();
        }
    }
}

/// Emit a VM-internal op (prefix 0xFF) — lowered to WASM equivalents or runtime calls.
fn emit_vm_internal_op(body: &mut Vec<u8>, op: Op, chunk: &Chunk, ip: &mut usize, rt_idx: &std::collections::HashMap<(&str, &str), usize>) {
    match op {
        _ if op == Op::CONST => {
            let idx = read_u16(&chunk.code, ip);
            if let Some(val) = chunk.constants.get(idx as usize) {
                match val {
                    Value::F64(n) => { body.push(0x44); body.extend_from_slice(&n.to_le_bytes()); emit_box_f64(body, rt_idx); }
                    Value::I32(n) => { body.push(0x41); write_leb128_i32(body, *n); emit_box_i32(body, rt_idx); }
                    Value::I64(n) => { body.push(0x42); write_leb128_i64(body, *n); emit_box_i32(body, rt_idx); }
                    _ => { body.push(0xD0); body.push(0x6F); } // ref.null externref for non-numeric
                }
            }
        }
        _ if op == Op::NULL => { body.push(0xD0); body.push(0x6F); } // ref.null externref
        _ if op == Op::TRUE => { body.push(0x41); write_leb128_i32(body, 1); emit_box_i32(body, rt_idx); }
        _ if op == Op::FALSE || op == Op::I32_CONST_0 => { body.push(0x41); write_leb128_i32(body, 0); emit_box_i32(body, rt_idx); }
        _ if op == Op::I32_CONST_1 => { body.push(0x41); write_leb128_i32(body, 1); emit_box_i32(body, rt_idx); }
        _ if op == Op::F64_CONST_0 => { body.push(0x44); body.extend_from_slice(&0.0f64.to_le_bytes()); emit_box_f64(body, rt_idx); }
        _ if op == Op::CALL_IMPORT => {
            let import_idx = read_u16(&chunk.code, ip);
            let _argc = chunk.code[*ip]; *ip += 1;
            body.push(0x10);
            write_leb128_u32(body, import_idx as u32);
        }
        _ if op == Op::BR_IF_FALSE || op == Op::BR_IF_NULL => {
            let _ = read_i16(&chunk.code, ip);
            body.push(0x0D);
            write_leb128_u32(body, 0);
        }
        // Dynamic ops → runtime calls
        _ if op == Op::DYN_ADD => emit_rt_call(body, rt_idx, "dyn_add"),
        _ if op == Op::DYN_EQ => emit_rt_call(body, rt_idx, "dyn_eq"),
        _ if op == Op::DYN_NE => emit_rt_call(body, rt_idx, "dyn_ne"),
        _ if op == Op::DYN_LT => emit_rt_call(body, rt_idx, "dyn_lt"),
        _ if op == Op::DYN_GT => emit_rt_call(body, rt_idx, "dyn_gt"),
        _ if op == Op::DYN_LE => emit_rt_call(body, rt_idx, "dyn_le"),
        _ if op == Op::DYN_GE => emit_rt_call(body, rt_idx, "dyn_ge"),
        _ if op == Op::DYN_NEG => emit_rt_call(body, rt_idx, "dyn_neg"),
        _ if op == Op::DYN_NOT => emit_rt_call(body, rt_idx, "dyn_not"),
        _ if op == Op::DYN_TO_BOOL => emit_rt_call(body, rt_idx, "dyn_to_bool"),
        _ if op == Op::STR_CONCAT => emit_rt_call(body, rt_idx, "str_concat"),
        _ if op == Op::GLOBAL_GET => { let _ = read_u16(&chunk.code, ip); emit_rt_call(body, rt_idx, "global_get"); }
        _ if op == Op::GLOBAL_SET => { let _ = read_u16(&chunk.code, ip); emit_rt_call(body, rt_idx, "global_set"); }
        _ => {
            // Skip operands, emit nop
            let fmt = op.operand_format();
            match fmt {
                OperandFormat::Closure => {
                    let _ = read_u16(&chunk.code, ip);
                    let uv = chunk.code.get(*ip).copied().unwrap_or(0) as usize;
                    *ip += 1 + uv * 2;
                }
                OperandFormat::BrTable => {
                    let count = chunk.code.get(*ip).copied().unwrap_or(0) as usize;
                    *ip += 2 + count;
                }
                OperandFormat::TryTable => {
                    let count = chunk.code.get(*ip).copied().unwrap_or(0) as usize;
                    *ip += 1 + count * 3;
                }
                _ => { *ip += fmt.fixed_size(); }
            }
            body.push(0x01); // nop
        }
    }
}

/// Total instruction size: 2-byte opcode + operand bytes.
pub fn opcode_size(op: Op, code: &[u8], ip: usize) -> usize {
    let base = 2;
    match op.operand_format() {
        OperandFormat::Closure => {
            let uv_count = code.get(ip + 4).copied().unwrap_or(0) as usize;
            base + 2 + 1 + uv_count * 2
        }
        OperandFormat::BrTable => {
            let count = code.get(ip + 2).copied().unwrap_or(0) as usize;
            base + 2 + count
        }
        OperandFormat::TryTable => {
            let count = code.get(ip + 2).copied().unwrap_or(0) as usize;
            base + 1 + count * 3
        }
        fmt => base + fmt.fixed_size(),
    }
}
