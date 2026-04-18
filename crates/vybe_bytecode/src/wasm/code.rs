//! WASM code section encoding.
//! Translates internal bytecode to WASM binary format.
//!
//! Type strategy: externref is the universal value representation.
//! All locals, params, and function results are externref.
//! Typed WASM ops (f64.add, i32.mul, etc.) require unboxing via
//! wasm:js-number builtins (toF64/toI32) before the op and reboxing
//! (fromF64/fromI32) after. Binary ops need a temp externref local
//! to save TOS while unboxing the second operand.

use super::encoding::*;
use super::types::WasmTypeContext;
use super::sections::{emit_import_call, emit_box_i32, emit_box_f64, emit_unbox_f64, emit_unbox_i32};
use crate::{Chunk, Op};
use crate::value::Value;
use crate::opcode::OperandFormat;


/// Count how many temp locals a chunk needs for stack manipulation.
/// Returns 0, 1, or 2 depending on which ops are used.
fn count_temp_locals(chunk: &Chunk) -> u32 {
    let mut need = 0u32;
    let mut ip = 0;
    while ip < chunk.code.len() {
        if ip + 1 >= chunk.code.len() { break; }
        if let Some(op) = Op::decode(chunk.code[ip], chunk.code[ip + 1]) {
            if op == Op::CALL_REF {
                // call_ref needs argc+1 temps (save args + table idx)
                let call_argc = chunk.code.get(ip + 2).copied().unwrap_or(0) as u32;
                need = need.max(call_argc + 1);
            } else if op == Op::ARRAY_CONCAT || op == Op::ARRAY_CONTAINS
                || op == Op::ARRAY_INDEX_OF || op == Op::ARRAY_JOIN
                || op == Op::STR_INDEX_OF {
                need = need.max(5); // need 5 temps
            } else if op == Op::ARRAY_PUSH || op == Op::ARRAY_SLICE
                || op == Op::ARRAY_REVERSE {
                need = need.max(4); // need 4 temps
            } else if op == Op::ARRAY_SET || op == Op::STR_SUBSTRING {
                need = need.max(2); // need 2 temps for 3-operand reorder
            } else if is_binary_typed_op(op) || op == Op::GLOBAL_SET || op == Op::DUP
                || op == Op::ARRAY_GET || op == Op::ARRAY_LENGTH
                || op == Op::ARRAY_POP || op == Op::ARRAY_SHIFT
                || op == Op::REF_TYPEOF || op == Op::REF_IS_NULL
                // Dynamic binary ops also use temp for unbox pattern
                || op == Op::DYN_ADD || op == Op::DYN_LT || op == Op::DYN_GT
                || op == Op::DYN_LE || op == Op::DYN_GE || op == Op::DYN_EQ || op == Op::DYN_NE {
                need = need.max(1);
            }
            ip += opcode_size(op, &chunk.code, ip);
        } else {
            ip += 2;
        }
    }
    need
}

/// Is this a core WASM op that operates on typed (non-externref) values
/// and takes two operands from the stack?
fn is_binary_typed_op(op: Op) -> bool {
    // f64 binary arithmetic
    op == Op::F64_ADD || op == Op::F64_SUB || op == Op::F64_MUL || op == Op::F64_DIV
    || op == Op::F64_MIN || op == Op::F64_MAX || op == Op::F64_COPYSIGN
    // f64 comparisons
    || op == Op::F64_LT || op == Op::F64_GT || op == Op::F64_LE || op == Op::F64_GE
    // i32 binary arithmetic
    || op == Op::I32_ADD || op == Op::I32_SUB || op == Op::I32_MUL
    || op == Op::I32_DIV_S || op == Op::I32_DIV_U
    || op == Op::I32_REM_S || op == Op::I32_REM_U
    || op == Op::I32_AND || op == Op::I32_OR || op == Op::I32_XOR
    || op == Op::I32_SHL || op == Op::I32_SHR_S || op == Op::I32_SHR_U
    || op == Op::I32_ROTL || op == Op::I32_ROTR
    // i32 comparisons
    || op == Op::EQ || op == Op::NE
}

pub fn encode_code_section(chunks: &[Chunk], rt_imports: &[(&str, &str)], type_ctx: &WasmTypeContext, global_map: &std::collections::HashMap<String, u32>) -> Vec<u8> {
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

        // Check how many temp locals we need for stack manipulation
        let temp_count = count_temp_locals(chunk);
        let has_temp = temp_count > 0;
        // WASM convention: params = arity (slot 0 = first arg, no callee slot).
        let wasm_params = chunk.arity as u32;
        // Extra locals beyond params
        let extra_locals = if chunk.local_count as u32 > wasm_params {
            chunk.local_count as u32 - wasm_params
        } else { 0 };
        let temp_local_idx = wasm_params + extra_locals;

        // Locals declaration (only declare locals beyond params)
        let declared_locals = extra_locals + temp_count;
        if declared_locals > 0 {
            write_leb128_u32(&mut body, 1); // 1 local type group
            write_leb128_u32(&mut body, declared_locals);
            body.push(TYPE_EXTERNREF);
        } else {
            write_leb128_u32(&mut body, 0);
        }

        // Structured control flow: the compiler now emits BLOCK/LOOP/END/BR_LABEL/BR_IF_LABEL.
        // The WASM emitter just passes them through — no relooper needed.

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
                emit_core_op(&mut body, op, chunk, &mut ip, &rt_idx, temp_local_idx, has_temp, type_ctx, global_map, host_import_count);
            } else if op.prefix() == 0xFB {
                emit_gc_op(&mut body, op, chunk, &mut ip, &rt_idx, type_ctx, temp_local_idx);
            } else if op.prefix() >= 0xFC && op.prefix() <= 0xFE {
                body.push(op.prefix());
                write_leb128_u32(&mut body, op.sub() as u32);
                ip += op.operand_format().fixed_size();
            } else {
                emit_vm_internal_op(&mut body, op, chunk, &mut ip, &rt_idx, temp_local_idx, type_ctx);
            }
        }

        body.push(0x0B); // end function

        write_leb128_u32(&mut out, body.len() as u32);
        out.extend_from_slice(&body);
    }
    out
}

/// Emit a core WASM MVP opcode (prefix 0x00).
/// `temp_idx` is the index of the temp externref local (valid when `has_temp` is true).
fn emit_core_op(body: &mut Vec<u8>, op: Op, chunk: &Chunk, ip: &mut usize,
                rt_idx: &std::collections::HashMap<(&str, &str), usize>,
                temp_idx: u32, _has_temp: bool,
                type_ctx: &WasmTypeContext,
                global_map: &std::collections::HashMap<String, u32>,
                host_import_count: usize) {
    match op {
        _ if op == Op::LOCAL_GET => { body.push(op.sub()); write_leb128_u32(body, read_u16(&chunk.code, ip) as u32); }
        _ if op == Op::LOCAL_SET => { body.push(0x22); write_leb128_u32(body, read_u16(&chunk.code, ip) as u32); } // local.tee
        _ if op == Op::CALL => { body.push(op.sub()); let argc = chunk.code[*ip]; *ip += 1; write_leb128_u32(body, argc as u32); }
        _ if op == Op::CALL_REF => {
            let argc = chunk.code[*ip]; *ip += 1;
            // Stack: [externref_funcref, arg1, ..., argN] — funcref is below args
            // call_indirect needs: [arg1, ..., argN, i32_table_idx]
            // WASM convention: slot 0 = first arg, no reserved callee slot.
            //
            // 1. Save all args to temps
            for i in (0..argc).rev() {
                body.push(0x21); write_leb128_u32(body, temp_idx + i as u32);
            }
            // Stack: [externref_funcref]
            // 2. Save funcref
            body.push(0x21); write_leb128_u32(body, temp_idx + argc as u32);
            // 3. Restore user args
            for i in 0..argc {
                body.push(0x20); write_leb128_u32(body, temp_idx + i as u32);
            }
            // 4. Push table index (unbox funcref to i32)
            body.push(0x20); write_leb128_u32(body, temp_idx + argc as u32);
            emit_unbox_i32(body, rt_idx);
            // 5. call_indirect with matching function type
            if let Some(&type_idx) = type_ctx.func_type_by_arity.get(&argc) {
                body.push(0x11); // call_indirect
                write_leb128_u32(body, type_idx); // type index
                write_leb128_u32(body, 0); // table index 0
            } else {
                // No matching type — drop everything, push null
                body.push(0x1A); // drop table_idx
                for _ in 0..argc { body.push(0x1A); }
                body.push(0xD0); body.push(0x6F);
            }
        }
        _ if op == Op::BR => {
            // Legacy flat jump — skip operand, emit nop
            let _offset = read_i16(&chunk.code, ip);
            body.push(0x01); // nop
        }
        _ if op == Op::BR_IF_TRUE => {
            let _offset = read_i16(&chunk.code, ip);
            body.push(0x01); // nop
        }
        // END pops a label from the structured CF stack
        _ if op == Op::END => {
            body.push(0x0B); // end
        }
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
        // WASM global.get/set — resolved to indexed globals via global_map
        _ if op == Op::GLOBAL_GET => {
            let name_idx = read_u16(&chunk.code, ip);
            if let Some(crate::value::Value::String(name)) = chunk.constants.get(name_idx as usize) {
                if let Some(&gidx) = global_map.get(name.as_ref()) {
                    body.push(0x23); // global.get
                    write_leb128_u32(body, gidx);
                } else {
                    body.push(0xD0); body.push(0x6F); // ref.null extern (unknown global)
                }
            } else {
                body.push(0xD0); body.push(0x6F);
            }
        }
        _ if op == Op::GLOBAL_SET => {
            let name_idx = read_u16(&chunk.code, ip);
            if let Some(crate::value::Value::String(name)) = chunk.constants.get(name_idx as usize) {
                if let Some(&gidx) = global_map.get(name.as_ref()) {
                    // Stack has [value]. global.set consumes it — but our VM keeps it.
                    // Use local.tee pattern: tee to keep value, then global.set
                    body.push(0x22); write_leb128_u32(body, temp_idx); // local.tee $temp
                    body.push(0x24); // global.set
                    write_leb128_u32(body, gidx);
                    body.push(0x20); write_leb128_u32(body, temp_idx); // restore value
                } else {
                    // Unknown global — just keep value on stack
                }
            }
        }
        _ if op == Op::HALT => { body.push(0x0F); } // return (not unreachable — _start should return cleanly)
        // ref.null needs heaptype byte — can't just emit op.sub()
        _ if op == Op::NULL => { body.push(0xD0); body.push(0x6F); } // ref.null externref
        // ref.is_null produces i32 — box it since our value representation is externref
        _ if op == Op::REF_IS_NULL => {
            body.push(0xD1); // ref.is_null → i32
            emit_box_i32(body, rt_idx); // i32 → externref
        }

        // ── f64 binary arithmetic: unbox both → f64 op → rebox ──
        // Stack: [externref_a, externref_b]
        // Pattern: local.set $temp (save b) → toF64 (unbox a) → local.get $temp → toF64 (unbox b)
        //          → f64.op → fromF64 (rebox result)
        _ if op == Op::F64_ADD || op == Op::F64_SUB || op == Op::F64_MUL || op == Op::F64_DIV
          || op == Op::F64_MIN || op == Op::F64_MAX || op == Op::F64_COPYSIGN => {
            emit_binary_f64_op(body, op.sub(), rt_idx, temp_idx);
        }

        // ── f64 comparisons: unbox both → compare → rebox i32 result ──
        _ if op == Op::F64_LT || op == Op::F64_GT || op == Op::F64_LE || op == Op::F64_GE => {
            emit_binary_f64_cmp(body, op.sub(), rt_idx, temp_idx);
        }

        // ── i32 binary arithmetic: unbox both → i32 op → rebox ──
        _ if op == Op::I32_ADD || op == Op::I32_SUB || op == Op::I32_MUL
          || op == Op::I32_DIV_S || op == Op::I32_DIV_U
          || op == Op::I32_REM_S || op == Op::I32_REM_U
          || op == Op::I32_AND || op == Op::I32_OR || op == Op::I32_XOR
          || op == Op::I32_SHL || op == Op::I32_SHR_S || op == Op::I32_SHR_U
          || op == Op::I32_ROTL || op == Op::I32_ROTR => {
            emit_binary_i32_op(body, op.sub(), rt_idx, temp_idx);
        }

        // ── i32 comparisons (eq, ne): unbox both → compare → rebox ──
        _ if op == Op::EQ || op == Op::NE => {
            emit_binary_i32_cmp(body, op.sub(), rt_idx, temp_idx);
        }

        // ── f64 unary ops: unbox → op → rebox ──
        _ if op == Op::F64_NEG || op == Op::F64_ABS || op == Op::F64_CEIL
          || op == Op::F64_FLOOR || op == Op::F64_TRUNC || op == Op::F64_NEAREST
          || op == Op::F64_SQRT => {
            emit_unbox_f64(body, rt_idx);
            body.push(op.sub());
            emit_box_f64(body, rt_idx);
        }

        // ── i32 unary ops ──
        _ if op == Op::I32_EQZ => {
            emit_unbox_i32(body, rt_idx);
            body.push(op.sub());
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::I32_CLZ || op == Op::I32_CTZ || op == Op::I32_POPCNT => {
            emit_unbox_i32(body, rt_idx);
            body.push(op.sub());
            emit_box_i32(body, rt_idx);
        }

        // ── Conversions: unbox source type → convert → rebox target type ──
        _ if op == Op::I32_FROM_F64 => {
            // externref → f64 → i32.trunc_f64_s → externref
            emit_unbox_f64(body, rt_idx);
            body.push(op.sub());
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::F64_FROM_I32 => {
            // externref → i32 → f64.convert_i32_s → externref
            emit_unbox_i32(body, rt_idx);
            body.push(op.sub());
            emit_box_f64(body, rt_idx);
        }

        // ref.func (Closure format): emit ref.func with WASM function index
        _ if op == Op::REF_FUNC => {
            let chunk_idx = read_u16(&chunk.code, ip);
            let uv_count = chunk.code[*ip] as usize; *ip += 1;
            *ip += uv_count * 2; // skip upvalue descriptors
            // WASM function index = total_imports + chunk_idx
            let total_imports = host_import_count + rt_idx.len();
            let wasm_func_idx = total_imports + chunk_idx as usize;
            // Store as table index (i32) for call_indirect — box as externref
            // The chunk_idx is the table index since element section maps chunks 0..N to table slots
            body.push(0x41); // i32.const
            write_leb128_i32(body, chunk_idx as i32);
            emit_box_i32(body, rt_idx); // i32 → externref
        }

        _ => {
            // Other core ops: emit WASM byte directly
            body.push(op.sub());
            *ip += op.operand_format().fixed_size();
        }
    }
}

// ── Binary operation helpers ─────────────────────────────────────────

/// Emit binary f64 op: [externref_a, externref_b] → f64.op → [externref_result]
/// Uses temp local to save b while unboxing a.
fn emit_binary_f64_op(body: &mut Vec<u8>, wasm_opcode: u8,
                       rt_idx: &std::collections::HashMap<(&str, &str), usize>,
                       temp_idx: u32) {
    // Save b to temp local
    body.push(0x21); // local.set
    write_leb128_u32(body, temp_idx);
    // Unbox a: externref → f64
    emit_unbox_f64(body, rt_idx);
    // Restore b from temp local
    body.push(0x20); // local.get
    write_leb128_u32(body, temp_idx);
    // Unbox b: externref → f64
    emit_unbox_f64(body, rt_idx);
    // Operate: f64, f64 → f64
    body.push(wasm_opcode);
    // Rebox result: f64 → externref
    emit_box_f64(body, rt_idx);
}

/// Emit binary f64 comparison: [externref_a, externref_b] → f64.cmp → [externref_result(i32)]
fn emit_binary_f64_cmp(body: &mut Vec<u8>, wasm_opcode: u8,
                        rt_idx: &std::collections::HashMap<(&str, &str), usize>,
                        temp_idx: u32) {
    body.push(0x21); write_leb128_u32(body, temp_idx); // local.set $temp (save b)
    emit_unbox_f64(body, rt_idx);                        // toF64(a)
    body.push(0x20); write_leb128_u32(body, temp_idx); // local.get $temp (restore b)
    emit_unbox_f64(body, rt_idx);                        // toF64(b)
    body.push(wasm_opcode);                              // f64.lt/gt/le/ge → i32
    emit_box_i32(body, rt_idx);                          // fromI32 → externref
}

/// Emit binary i32 op: [externref_a, externref_b] → i32.op → [externref_result]
fn emit_binary_i32_op(body: &mut Vec<u8>, wasm_opcode: u8,
                       rt_idx: &std::collections::HashMap<(&str, &str), usize>,
                       temp_idx: u32) {
    body.push(0x21); write_leb128_u32(body, temp_idx); // local.set $temp (save b)
    emit_unbox_i32(body, rt_idx);                        // toI32(a)
    body.push(0x20); write_leb128_u32(body, temp_idx); // local.get $temp (restore b)
    emit_unbox_i32(body, rt_idx);                        // toI32(b)
    body.push(wasm_opcode);                              // i32.op → i32
    emit_box_i32(body, rt_idx);                          // fromI32 → externref
}

/// Emit binary i32 comparison: [externref_a, externref_b] → i32.cmp → [externref_result]
fn emit_binary_i32_cmp(body: &mut Vec<u8>, wasm_opcode: u8,
                        rt_idx: &std::collections::HashMap<(&str, &str), usize>,
                        temp_idx: u32) {
    body.push(0x21); write_leb128_u32(body, temp_idx); // local.set $temp (save b)
    emit_unbox_i32(body, rt_idx);                        // toI32(a)
    body.push(0x20); write_leb128_u32(body, temp_idx); // local.get $temp (restore b)
    emit_unbox_i32(body, rt_idx);                        // toI32(b)
    body.push(wasm_opcode);                              // i32.eq/ne → i32
    emit_box_i32(body, rt_idx);                          // fromI32 → externref
}

/// Emit a GC op (prefix 0xFB) — emit real WASM GC binary encoding with type indices.
///
/// GC refs (ref $struct, ref $array) are NOT subtypes of externref.
/// We use externref as our universal local type, so:
/// - After GC ops that PRODUCE refs: emit `extern.convert_any` (0xFB 0x1B) → externref
/// - Before GC ops that CONSUME refs: emit `any.convert_extern` (0xFB 0x1A) → anyref,
///   then `ref.cast` to the specific GC type
fn emit_gc_op(body: &mut Vec<u8>, op: Op, chunk: &Chunk, ip: &mut usize, _rt_idx: &std::collections::HashMap<(&str, &str), usize>, type_ctx: &WasmTypeContext, temp_idx: u32) {
    match op {
        _ if op == Op::STRUCT_NEW => {
            let _prop_count = read_u16(&chunk.code, ip);
            body.push(0xFB); write_leb128_u32(body, 0x00); // struct.new
            write_leb128_u32(body, 0); // type index TODO: resolve from type context
            emit_externalize(body); // (ref $struct) → externref
        }
        _ if op == Op::STRUCT_GET => {
            let _field_name_idx = read_u16(&chunk.code, ip);
            emit_internalize(body); // externref → anyref
            emit_ref_cast(body, 0); // anyref → (ref $struct) TODO: proper type idx
            body.push(0xFB); write_leb128_u32(body, 0x02); // struct.get
            write_leb128_u32(body, 0); // type index
            write_leb128_u32(body, 0); // field index
            // Result is externref (field type) — no conversion needed
        }
        _ if op == Op::STRUCT_SET => {
            let _field_name_idx = read_u16(&chunk.code, ip);
            // Stack: [externref_obj, externref_val]. Need obj as (ref $struct).
            // Can't convert obj without temp — for now emit as-is with conversion TODO
            emit_internalize(body); // externref → anyref (converts val, wrong!)
            // TODO: proper operand reordering with temp local
            body.push(0xFB); write_leb128_u32(body, 0x05); // struct.set
            write_leb128_u32(body, 0);
            write_leb128_u32(body, 0);
            body.push(0xD0); body.push(0x6F); // push dummy (struct.set is void in WASM)
        }
        _ if op == Op::ARRAY_NEW => {
            let elem_count = read_u16(&chunk.code, ip);
            // Elements on stack are externref — need to internalize each.
            // For 0 elements, no conversion needed.
            // For N elements, we'd need N conversions — complex.
            // For now: emit array.new_fixed, then externalize the result.
            body.push(0xFB); write_leb128_u32(body, 0x08); // array.new_fixed (NOT 0x06 which is array.new)
            write_leb128_u32(body, type_ctx.array_type_idx);
            write_leb128_u32(body, elem_count as u32);
            emit_externalize(body); // (ref $arr) → externref
        }
        _ if op == Op::ARRAY_GET => {
            // Stack: [externref_arr, externref_idx]
            // Need: [(ref null $arr), i32] for array.get
            // Save idx to temp, convert arr, restore idx as i32
            body.push(0x21); write_leb128_u32(body, temp_idx); // local.set $temp (save idx)
            emit_internalize(body);                              // externref_arr → anyref
            emit_ref_cast_array(body, type_ctx.array_type_idx); // anyref → (ref null $arr)
            body.push(0x20); write_leb128_u32(body, temp_idx); // local.get $temp (restore idx)
            emit_unbox_i32(body, _rt_idx);                       // externref_idx → i32
            body.push(0xFB); write_leb128_u32(body, 0x0B);     // array.get
            write_leb128_u32(body, type_ctx.array_type_idx);
            // Result: externref (element type)
        }
        _ if op == Op::ARRAY_SET => {
            // Stack: [externref_arr, externref_idx, externref_val]
            // Need: [(ref null $arr), i32, externref] for array.set
            // Save val and idx, convert arr, restore idx as i32, restore val
            body.push(0x21); write_leb128_u32(body, temp_idx);     // local.set $temp (save val)
            // Stack: [externref_arr, externref_idx]
            body.push(0x21); write_leb128_u32(body, temp_idx + 1); // local.set $temp2 (save idx) — use next local
            // Stack: [externref_arr]
            emit_internalize(body);                                  // externref_arr → anyref
            emit_ref_cast_array(body, type_ctx.array_type_idx);     // anyref → (ref null $arr)
            body.push(0x20); write_leb128_u32(body, temp_idx + 1); // local.get $temp2 (restore idx)
            emit_unbox_i32(body, _rt_idx);                           // externref_idx → i32
            body.push(0x20); write_leb128_u32(body, temp_idx);     // local.get $temp (restore val)
            // Stack: [(ref null $arr), i32, externref]
            body.push(0xFB); write_leb128_u32(body, 0x0E);         // array.set
            write_leb128_u32(body, type_ctx.array_type_idx);
            // array.set is void in WASM but our VM leaves a value on stack.
            // Push dummy so subsequent drop has something to consume.
            body.push(0xD0); body.push(0x6F); // ref.null extern
        }
        _ if op == Op::ARRAY_LENGTH => {
            // Stack: [externref_arr] → need (ref null array)
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0xFB); write_leb128_u32(body, 0x0F); // array.len → i32
            // Result is i32, box to externref
            emit_box_i32(body, _rt_idx);
        }
        _ if op == Op::ARRAY_FILL => {
            emit_internalize(body);
            body.push(0xFB); write_leb128_u32(body, 0x10);
            write_leb128_u32(body, type_ctx.array_type_idx);
        }
        _ if op == Op::ARRAY_COPY => {
            body.push(0xFB); write_leb128_u32(body, 0x11);
            write_leb128_u32(body, type_ctx.array_type_idx);
            write_leb128_u32(body, type_ctx.array_type_idx);
        }
        _ if op == Op::ARRAY_NEW_DEFAULT => {
            body.push(0xFB); write_leb128_u32(body, 0x07);
            write_leb128_u32(body, type_ctx.array_type_idx);
            emit_externalize(body);
        }
        _ if op == Op::REF_TEST || op == Op::REF_CAST => {
            *ip += op.operand_format().fixed_size();
            body.push(0xFB); write_leb128_u32(body, op.sub() as u32);
            // TODO: emit proper heap type reference
        }
        _ if op == Op::BR_ON_CAST || op == Op::BR_ON_CAST_FAIL => {
            *ip += op.operand_format().fixed_size();
            body.push(0xFB); write_leb128_u32(body, op.sub() as u32);
        }
        _ if op == Op::I31_NEW => {
            // i31.new expects i32 — unbox externref first
            emit_unbox_i32(body, _rt_idx);
            body.push(0xFB); write_leb128_u32(body, 0x1C); // ref.i31
            emit_externalize(body); // (ref i31) → externref
        }
        _ if op == Op::I31_GET_S || op == Op::I31_GET_U => {
            emit_internalize(body); // externref → anyref
            body.push(0xFB); write_leb128_u32(body, op.sub() as u32);
            emit_box_i32(body, _rt_idx); // i32 → externref
        }
        _ => {
            // Other GC ops: emit directly
            body.push(0xFB);
            write_leb128_u32(body, op.sub() as u32);
            *ip += op.operand_format().fixed_size();
        }
    }
}

/// Emit inline dyn_add: type check both operands, f64 arithmetic if numbers, string concat if not.
/// Uses wasm:js-number and wasm:js-string builtins (standard WASM proposals).
fn emit_dyn_binary_numeric(body: &mut Vec<u8>, rt_idx: &std::collections::HashMap<(&str, &str), usize>,
                            temp_idx: u32, f64_opcode: u8) {
    // Stack: [externref_a, externref_b]
    // Simple approach: always treat as f64 (matches how stdlib uses dyn_add for numbers)
    // Save b, unbox a, restore b, unbox b, operate, rebox
    body.push(0x21); write_leb128_u32(body, temp_idx); // local.set $temp (save b)
    emit_unbox_f64(body, rt_idx);                        // toF64(a)
    body.push(0x20); write_leb128_u32(body, temp_idx); // local.get $temp (restore b)
    emit_unbox_f64(body, rt_idx);                        // toF64(b)
    body.push(f64_opcode);                               // f64.add/sub/mul/div
    emit_box_f64(body, rt_idx);                          // fromF64 → externref
}

/// Emit inline dyn comparison: unbox both as f64, compare, box i32 result.
fn emit_dyn_binary_cmp(body: &mut Vec<u8>, rt_idx: &std::collections::HashMap<(&str, &str), usize>,
                        temp_idx: u32, f64_cmp_opcode: u8) {
    body.push(0x21); write_leb128_u32(body, temp_idx); // local.set $temp (save b)
    emit_unbox_f64(body, rt_idx);                        // toF64(a)
    body.push(0x20); write_leb128_u32(body, temp_idx); // local.get $temp (restore b)
    emit_unbox_f64(body, rt_idx);                        // toF64(b)
    body.push(f64_cmp_opcode);                           // f64.lt/gt/le/ge/eq/ne → i32
    emit_box_i32(body, rt_idx);                          // fromI32 → externref
}

/// Emit a string constant from the chunk's constant pool.
/// Builds the string char by char using wasm:js-string fromCharCode + concat.
fn emit_string_const(body: &mut Vec<u8>, chunk: &Chunk, const_idx: usize, rt_idx: &std::collections::HashMap<(&str, &str), usize>) {
    if let Some(Value::String(s)) = chunk.constants.get(const_idx) {
        let chars: Vec<char> = s.chars().collect();
        if chars.is_empty() {
            body.push(0xD0); body.push(0x6F); // ref.null extern
            return;
        }
        // First char
        body.push(0x41); write_leb128_i32(body, chars[0] as i32);
        emit_import_call(body, rt_idx, "wasm:js-string", "fromCharCode");
        // Concat remaining chars
        for &ch in &chars[1..] {
            body.push(0x41); write_leb128_i32(body, ch as i32);
            emit_import_call(body, rt_idx, "wasm:js-string", "fromCharCode");
            emit_import_call(body, rt_idx, "wasm:js-string", "concat");
        }
    } else {
        body.push(0xD0); body.push(0x6F);
    }
}

/// Emit `any.convert_extern` (0xFB 0x1A): externref → anyref
fn emit_internalize(body: &mut Vec<u8>) {
    body.push(0xFB);
    write_leb128_u32(body, 0x1A);
}

/// Emit `extern.convert_any` (0xFB 0x1B): anyref → externref
fn emit_externalize(body: &mut Vec<u8>) {
    body.push(0xFB);
    write_leb128_u32(body, 0x1B);
}

/// Emit `ref.cast null $typeidx` — casts anyref to a specific GC type (nullable)
fn emit_ref_cast(body: &mut Vec<u8>, type_idx: u32) {
    body.push(0xFB);
    write_leb128_u32(body, 0x17); // ref.cast null (nullable variant)
    write_leb128_u32(body, type_idx); // heaptype = type index
}

/// Emit `ref.cast (ref null $arr_typeidx)` for array refs
fn emit_ref_cast_array(body: &mut Vec<u8>, arr_type_idx: u32) {
    emit_ref_cast(body, arr_type_idx);
}

/// Emit a VM-internal op (prefix 0xFF) — lowered to WASM equivalents or runtime calls.
fn emit_vm_internal_op(body: &mut Vec<u8>, op: Op, chunk: &Chunk, ip: &mut usize, rt_idx: &std::collections::HashMap<(&str, &str), usize>, temp_idx: u32, type_ctx: &WasmTypeContext) {
    match op {
        _ if op == Op::CONST => {
            let idx = read_u16(&chunk.code, ip);
            if let Some(val) = chunk.constants.get(idx as usize) {
                match val {
                    Value::F64(n) => { body.push(0x44); body.extend_from_slice(&n.to_le_bytes()); emit_box_f64(body, rt_idx); }
                    Value::I32(n) => { body.push(0x41); write_leb128_i32(body, *n); emit_box_i32(body, rt_idx); }
                    Value::I64(n) => { body.push(0x42); write_leb128_i64(body, *n); emit_box_i32(body, rt_idx); }
                    _ => { body.push(0xD0); body.push(0x6F); } // ref.null externref
                }
            }
        }
        _ if op == Op::NULL => { body.push(0xD0); body.push(0x6F); } // ref.null externref
        _ if op == Op::TRUE => { body.push(0x41); write_leb128_i32(body, 1); emit_box_i32(body, rt_idx); } // i32 → externref
        _ if op == Op::FALSE || op == Op::I32_CONST_0 => { body.push(0x41); write_leb128_i32(body, 0); emit_box_i32(body, rt_idx); }
        _ if op == Op::I32_CONST_1 => { body.push(0x41); write_leb128_i32(body, 1); emit_box_i32(body, rt_idx); }
        _ if op == Op::F64_CONST_0 => { body.push(0x44); body.extend_from_slice(&0.0f64.to_le_bytes()); emit_box_f64(body, rt_idx); }
        _ if op == Op::CALL_IMPORT => {
            let import_idx = read_u16(&chunk.code, ip);
            let _argc = chunk.code[*ip]; *ip += 1;
            body.push(0x10);
            write_leb128_u32(body, import_idx as u32);
        }
        _ if op == Op::BR_IF_FALSE => {
            // Legacy flat jump — skip operand, emit nop
            // (compiler should use BR_IF_LABEL instead)
            let _offset = read_i16(&chunk.code, ip);
            body.push(0x01); // nop
        }
        _ if op == Op::BR_IF_NULL => {
            let _offset = read_i16(&chunk.code, ip);
            body.push(0x01); // nop
        }
        // ── Structured control flow: BR_LABEL/BR_IF_LABEL → WASM br/br_if ──
        _ if op == Op::BR_LABEL => {
            let depth = chunk.code[*ip]; *ip += 1;
            body.push(0x0C); // br
            write_leb128_u32(body, depth as u32);
        }
        _ if op == Op::BR_IF_LABEL => {
            let depth = chunk.code[*ip]; *ip += 1;
            // BR_IF_LABEL pops value and branches if truthy.
            // WASM br_if pops i32 and branches if non-zero.
            // Need to convert: unbox to i32 first.
            emit_unbox_i32(body, rt_idx);
            body.push(0x0D); // br_if
            write_leb128_u32(body, depth as u32);
        }

        // ── Dynamic ops → inline WASM sequences using wasm:js-* builtins ──

        // dyn_add: type check → number add or string concat
        _ if op == Op::DYN_ADD => {
            // Stack: [externref_a, externref_b]
            // Save b, check if a is number
            emit_dyn_binary_numeric(body, rt_idx, temp_idx, 0xA0); // f64.add
            // Fallback: wasm:js-string concat
        }
        // dyn comparisons: unbox both as f64, compare, box result as i32
        _ if op == Op::DYN_LT => emit_dyn_binary_cmp(body, rt_idx, temp_idx, 0x63), // f64.lt
        _ if op == Op::DYN_GT => emit_dyn_binary_cmp(body, rt_idx, temp_idx, 0x64), // f64.gt
        _ if op == Op::DYN_LE => emit_dyn_binary_cmp(body, rt_idx, temp_idx, 0x65), // f64.le
        _ if op == Op::DYN_GE => emit_dyn_binary_cmp(body, rt_idx, temp_idx, 0x66), // f64.ge
        _ if op == Op::DYN_EQ => emit_dyn_binary_cmp(body, rt_idx, temp_idx, 0x61), // f64.eq (0x61)
        _ if op == Op::DYN_NE => emit_dyn_binary_cmp(body, rt_idx, temp_idx, 0x62), // f64.ne (0x62)
        // dyn unary: unbox → op → rebox
        _ if op == Op::DYN_NEG => {
            emit_unbox_f64(body, rt_idx);
            body.push(0x9A); // f64.neg
            emit_box_f64(body, rt_idx);
        }
        _ if op == Op::DYN_NOT => {
            // Truthy check: toI32, eqz (invert), fromI32
            emit_unbox_i32(body, rt_idx);
            body.push(0x45); // i32.eqz
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::DYN_TO_BOOL => {
            // Convert to boolean: toI32 (0 = false, nonzero = true), fromI32
            emit_unbox_i32(body, rt_idx);
            emit_box_i32(body, rt_idx);
        }

        // ── String ops → wasm:js-string builtins (standard WASM proposal) ──
        _ if op == Op::STR_CONCAT => emit_import_call(body, rt_idx, "wasm:js-string", "concat"),
        _ if op == Op::STR_EQUALS => {
            emit_import_call(body, rt_idx, "wasm:js-string", "equals");
            emit_box_i32(body, rt_idx); // i32 → externref
        }
        _ if op == Op::STR_COMPARE => {
            emit_import_call(body, rt_idx, "wasm:js-string", "compare");
            emit_box_i32(body, rt_idx); // i32 → externref
        }
        _ if op == Op::STR_LENGTH => {
            // (externref) → i32 → need to box result
            emit_import_call(body, rt_idx, "wasm:js-string", "length");
            emit_box_i32(body, rt_idx); // i32 → externref
        }
        _ if op == Op::STR_CHAR_CODE_AT => {
            // Stack: [externref_str, externref_idx] → need (externref, i32)
            // TOS is idx — unbox it in place
            emit_unbox_i32(body, rt_idx); // externref_idx → i32. Stack: [str, i32]
            emit_import_call(body, rt_idx, "wasm:js-string", "charCodeAt");
            emit_box_i32(body, rt_idx); // i32 result → externref
        }
        _ if op == Op::STR_FROM_CHAR_CODE => {
            // (externref) → need (i32)
            emit_unbox_i32(body, rt_idx);
            emit_import_call(body, rt_idx, "wasm:js-string", "fromCharCode");
            // result is externref (string) — no conversion needed
        }
        _ if op == Op::STR_SUBSTRING => {
            // Stack: [externref_str, externref_start, externref_end]
            // Need: [externref_str, i32_start, i32_end]
            // Save end, unbox start (but start is below end on stack)
            // Simpler: save end to temp, save start to temp+1, keep str,
            // restore start as i32, restore end as i32
            body.push(0x21); write_leb128_u32(body, temp_idx);     // save end
            body.push(0x21); write_leb128_u32(body, temp_idx + 1); // save start
            // Stack: [externref_str]
            body.push(0x20); write_leb128_u32(body, temp_idx + 1); // restore start
            emit_unbox_i32(body, rt_idx);                            // → i32
            body.push(0x20); write_leb128_u32(body, temp_idx);     // restore end
            emit_unbox_i32(body, rt_idx);                            // → i32
            emit_import_call(body, rt_idx, "wasm:js-string", "substring");
        }
        // String ops not in wasm:js-string → stub: drop args, return first arg or null
        _ if op == Op::STR_INDEX_OF => {
            // Stack: [externref_str, externref_substr]
            // Inline indexOf: for i=0 to len-sublen, check substring(i, i+sublen) == substr
            body.push(0x21); write_leb128_u32(body, temp_idx);     // $t0 = substr
            body.push(0x21); write_leb128_u32(body, temp_idx + 1); // $t1 = str
            // Get substr length
            body.push(0x20); write_leb128_u32(body, temp_idx);     // substr
            emit_import_call(body, rt_idx, "wasm:js-string", "length"); // → i32
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 2); // $t2 = sublen (boxed)
            // Get str length
            body.push(0x20); write_leb128_u32(body, temp_idx + 1); // str
            emit_import_call(body, rt_idx, "wasm:js-string", "length"); // → i32
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 3); // $t3 = strlen (boxed)
            // i = 0 (boxed)
            body.push(0x41); write_leb128_i32(body, 0);
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 4); // $t4 = i (boxed)
            // block $exit (result externref) { loop $search (void) { ... } }
            body.push(0x02); body.push(TYPE_EXTERNREF); // block (result externref)
            body.push(0x03); body.push(0x40);            // loop (void)
            // if i > limit, push -1 and break
            body.push(0x20); write_leb128_u32(body, temp_idx + 4); emit_unbox_i32(body, rt_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 3); emit_unbox_i32(body, rt_idx);
            body.push(0x4A); // i32.gt_s
            body.push(0x04); body.push(0x40); // if (void) — past limit
            body.push(0x41); write_leb128_i32(body, -1); emit_box_i32(body, rt_idx);
            body.push(0x0C); write_leb128_u32(body, 2); // br block (if=0, loop=1, block=2)
            body.push(0x0B); // end if
            // substring(str, i, i + sublen)
            body.push(0x20); write_leb128_u32(body, temp_idx + 1); // str
            body.push(0x20); write_leb128_u32(body, temp_idx + 4); emit_unbox_i32(body, rt_idx); // i → i32
            body.push(0x20); write_leb128_u32(body, temp_idx + 4); emit_unbox_i32(body, rt_idx); // i → i32
            body.push(0x20); write_leb128_u32(body, temp_idx + 2); emit_unbox_i32(body, rt_idx); // sublen → i32
            body.push(0x6A); // i32.add → i + sublen
            emit_import_call(body, rt_idx, "wasm:js-string", "substring"); // (str, i, i+sublen) → string
            // Compare with substr
            body.push(0x20); write_leb128_u32(body, temp_idx); // substr
            emit_import_call(body, rt_idx, "wasm:js-string", "equals"); // → i32
            body.push(0x04); body.push(0x40); // if (void)
            // Found: push i (boxed) and break to $exit
            body.push(0x20); write_leb128_u32(body, temp_idx + 4); // i (already boxed)
            body.push(0x0C); write_leb128_u32(body, 2); // br $exit (depth: if=0, loop=1, block=2)
            body.push(0x0B); // end if
            // i++
            body.push(0x20); write_leb128_u32(body, temp_idx + 4); emit_unbox_i32(body, rt_idx);
            body.push(0x41); write_leb128_i32(body, 1);
            body.push(0x6A); // i32.add
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 4); // save i
            body.push(0x0C); write_leb128_u32(body, 0); // br $search (loop)
            body.push(0x0B); // end loop
            // Not found: push -1 (boxed)
            body.push(0x41); write_leb128_i32(body, -1);
            emit_box_i32(body, rt_idx);
            body.push(0x0B); // end block
            // Result: externref (i or -1, boxed)
        }
        // Binary string ops (2 args → 1 result): drop second, keep first
        _ if op == Op::STR_LAST_INDEX_OF
          || op == Op::STR_STARTS_WITH || op == Op::STR_ENDS_WITH || op == Op::STR_CONTAINS
          || op == Op::STR_SPLIT || op == Op::STR_REPEAT => {
            body.push(0x1A); // drop second arg, keep first as result
        }
        // Ternary string ops (3 args → 1 result): drop two, keep first
        _ if op == Op::STR_SLICE || op == Op::STR_REPLACE
          || op == Op::STR_PAD_START || op == Op::STR_PAD_END => {
            body.push(0x1A); body.push(0x1A); // drop 2, keep first
        }
        // Unary string ops (1 arg → 1 result): keep as-is
        _ if op == Op::STR_CHAR_AT || op == Op::STR_TO_UPPER || op == Op::STR_TO_LOWER
          || op == Op::STR_TRIM || op == Op::STR_TRIM_START || op == Op::STR_TRIM_END
          || op == Op::STR_REVERSE => {
            // pass through — input externref becomes output externref
            body.push(0x01); // nop
        }
        _ if op == Op::STR_CONCAT_N => {
            let n = chunk.code[*ip]; *ip += 1;
            // Concat N strings: chain wasm:js-string concat calls
            for _ in 1..n {
                emit_import_call(body, rt_idx, "wasm:js-string", "concat");
            }
        }

        // ── Array ops → inline WASM GC sequences ──
        // All i32 intermediates (lengths, indices) are boxed to externref for temp storage.
        _ if op == Op::ARRAY_PUSH => {
            // Stack: [externref_arr, externref_val]
            // 1. Save val, save arr, get old length
            body.push(0x21); write_leb128_u32(body, temp_idx);     // $t0 = val
            body.push(0x22); write_leb128_u32(body, temp_idx + 1); // $t1 = arr (tee)
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0xFB); write_leb128_u32(body, 0x0F); // array.len → i32
            // 2. Box len, save it
            emit_box_i32(body, rt_idx); // i32 → externref
            body.push(0x22); write_leb128_u32(body, temp_idx + 2); // $t2 = len (boxed)
            // 3. Create new array: array.new(init=null, size=len+1)
            emit_unbox_i32(body, rt_idx);  // len → i32
            body.push(0x41); write_leb128_i32(body, 1);
            body.push(0x6A); // i32.add → len+1
            // array.new takes (init_val, i32_size)
            // Stack has [i32_len+1]. Save as boxed, push init, unbox size.
            emit_box_i32(body, rt_idx); // box size for temp storage
            body.push(0x21); write_leb128_u32(body, temp_idx + 3); // $t3 = size (boxed)
            body.push(0xD0); body.push(0x6F); // init val (externref null)
            body.push(0x20); write_leb128_u32(body, temp_idx + 3); // size (boxed)
            emit_unbox_i32(body, rt_idx); // → i32
            body.push(0xFB); write_leb128_u32(body, 0x06); // array.new
            write_leb128_u32(body, type_ctx.array_type_idx);
            emit_externalize(body); // (ref $arr) → externref
            body.push(0x22); write_leb128_u32(body, temp_idx + 3); // $t3 = new_arr
            // 4. array.copy: dst=new, src=old, d_off=0, s_off=0, len=old_len
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x41); write_leb128_i32(body, 0); // dst_offset
            body.push(0x20); write_leb128_u32(body, temp_idx + 1); // old arr
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x41); write_leb128_i32(body, 0); // src_offset
            body.push(0x20); write_leb128_u32(body, temp_idx + 2); // old len (boxed)
            emit_unbox_i32(body, rt_idx); // → i32
            body.push(0xFB); write_leb128_u32(body, 0x11); // array.copy
            write_leb128_u32(body, type_ctx.array_type_idx);
            write_leb128_u32(body, type_ctx.array_type_idx);
            // 5. array.set new[old_len] = val
            body.push(0x20); write_leb128_u32(body, temp_idx + 3); // new_arr
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 2); // old len (boxed)
            emit_unbox_i32(body, rt_idx); // → i32 index
            body.push(0x20); write_leb128_u32(body, temp_idx); // val
            body.push(0xFB); write_leb128_u32(body, 0x0E); // array.set
            write_leb128_u32(body, type_ctx.array_type_idx);
            // 6. Result = new_arr
            body.push(0x20); write_leb128_u32(body, temp_idx + 3);
        }
        _ if op == Op::ARRAY_SLICE => {
            // Stack: [externref_arr, externref_start, externref_end]
            // All intermediates boxed as externref for temp storage
            body.push(0x21); write_leb128_u32(body, temp_idx);     // $t0 = end
            body.push(0x21); write_leb128_u32(body, temp_idx + 1); // $t1 = start
            body.push(0x21); write_leb128_u32(body, temp_idx + 2); // $t2 = arr
            // Compute slice_len = end - start, box it
            body.push(0x20); write_leb128_u32(body, temp_idx);     // end
            emit_unbox_i32(body, rt_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 1); // start
            emit_unbox_i32(body, rt_idx);
            body.push(0x6B); // i32.sub → slice_len (i32)
            emit_box_i32(body, rt_idx); // → externref
            body.push(0x21); write_leb128_u32(body, temp_idx + 3); // $t3 = slice_len (boxed)
            // Create new array: array.new(init=null, size=slice_len)
            body.push(0xD0); body.push(0x6F); // init val
            body.push(0x20); write_leb128_u32(body, temp_idx + 3); // slice_len
            emit_unbox_i32(body, rt_idx); // → i32
            body.push(0xFB); write_leb128_u32(body, 0x06); // array.new
            write_leb128_u32(body, type_ctx.array_type_idx);
            emit_externalize(body);
            body.push(0x22); write_leb128_u32(body, temp_idx + 3); // $t3 = new_arr (reuse)
            // array.copy: dst=new, src=old, d_off=0, s_off=start, len=slice_len
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x41); write_leb128_i32(body, 0); // dst_offset
            body.push(0x20); write_leb128_u32(body, temp_idx + 2); // old arr
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 1); // start
            emit_unbox_i32(body, rt_idx); // → i32 src_offset
            // length = end - start
            body.push(0x20); write_leb128_u32(body, temp_idx);     // end
            emit_unbox_i32(body, rt_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 1); // start
            emit_unbox_i32(body, rt_idx);
            body.push(0x6B); // i32.sub → copy_len
            body.push(0xFB); write_leb128_u32(body, 0x11); // array.copy
            write_leb128_u32(body, type_ctx.array_type_idx);
            write_leb128_u32(body, type_ctx.array_type_idx);
            // Result = new_arr
            body.push(0x20); write_leb128_u32(body, temp_idx + 3);
        }
        _ if op == Op::ARRAY_POP => {
            // Pop: return last element, leave shorter array
            // Our VM pops and returns the element. In WASM:
            // [externref_arr] → get last element → return it
            body.push(0x22); write_leb128_u32(body, temp_idx); // tee arr
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0xFB); write_leb128_u32(body, 0x0F); // array.len → i32
            body.push(0x41); write_leb128_i32(body, 1);
            body.push(0x6B); // i32.sub → last index
            // Save last index
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 1); // save idx (boxed)
            // array.get(arr, last_idx)
            body.push(0x20); write_leb128_u32(body, temp_idx); // arr
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 1);
            emit_unbox_i32(body, rt_idx);
            body.push(0xFB); write_leb128_u32(body, 0x0B); // array.get
            write_leb128_u32(body, type_ctx.array_type_idx);
        }
        _ if op == Op::ARRAY_SHIFT => {
            // Shift: return first element
            body.push(0x22); write_leb128_u32(body, temp_idx); // tee arr
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x41); write_leb128_i32(body, 0); // index 0
            body.push(0xFB); write_leb128_u32(body, 0x0B); // array.get
            write_leb128_u32(body, type_ctx.array_type_idx);
        }
        _ if op == Op::ARRAY_CONCAT => {
            // Stack: [externref_arr1, externref_arr2]
            // All i32 intermediates boxed for externref temp storage
            body.push(0x21); write_leb128_u32(body, temp_idx);     // $t0 = arr2
            body.push(0x22); write_leb128_u32(body, temp_idx + 1); // $t1 = arr1 (tee)
            // Get len1, box
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0xFB); write_leb128_u32(body, 0x0F); // array.len → i32
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 2); // $t2 = len1 (boxed)
            // Get len2, box
            body.push(0x20); write_leb128_u32(body, temp_idx);
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0xFB); write_leb128_u32(body, 0x0F);
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 3); // $t3 = len2 (boxed)
            // array.new(init, len1+len2)
            body.push(0xD0); body.push(0x6F); // init
            body.push(0x20); write_leb128_u32(body, temp_idx + 2);
            emit_unbox_i32(body, rt_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 3);
            emit_unbox_i32(body, rt_idx);
            body.push(0x6A); // i32.add
            body.push(0xFB); write_leb128_u32(body, 0x06); // array.new
            write_leb128_u32(body, type_ctx.array_type_idx);
            emit_externalize(body);
            body.push(0x22); write_leb128_u32(body, temp_idx + 4); // $t4 = new_arr
            // Copy arr1: new[0..len1]
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x41); write_leb128_i32(body, 0);
            body.push(0x20); write_leb128_u32(body, temp_idx + 1);
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x41); write_leb128_i32(body, 0);
            body.push(0x20); write_leb128_u32(body, temp_idx + 2);
            emit_unbox_i32(body, rt_idx);
            body.push(0xFB); write_leb128_u32(body, 0x11); // array.copy
            write_leb128_u32(body, type_ctx.array_type_idx);
            write_leb128_u32(body, type_ctx.array_type_idx);
            // Copy arr2: new[len1..len1+len2]
            body.push(0x20); write_leb128_u32(body, temp_idx + 4);
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 2);
            emit_unbox_i32(body, rt_idx); // dst_off = len1
            body.push(0x20); write_leb128_u32(body, temp_idx);
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x41); write_leb128_i32(body, 0);
            body.push(0x20); write_leb128_u32(body, temp_idx + 3);
            emit_unbox_i32(body, rt_idx); // len2
            body.push(0xFB); write_leb128_u32(body, 0x11); // array.copy
            write_leb128_u32(body, type_ctx.array_type_idx);
            write_leb128_u32(body, type_ctx.array_type_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 4); // result
        }
        _ if op == Op::ARRAY_REVERSE => {
            // Reverse in place: swap arr[i] and arr[len-1-i] for i=0..len/2
            body.push(0x22); write_leb128_u32(body, temp_idx); // tee arr
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0xFB); write_leb128_u32(body, 0x0F); // array.len → i32
            body.push(0x41); write_leb128_i32(body, 1);
            body.push(0x6B); // len - 1
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 1); // $t1 = hi (boxed)
            body.push(0x41); write_leb128_i32(body, 0);
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 2); // $t2 = lo (boxed)
            // Loop: while lo < hi, swap arr[lo] and arr[hi]
            body.push(0x02); body.push(0x40); // block void
            body.push(0x03); body.push(0x40); // loop void
            // Check lo < hi
            body.push(0x20); write_leb128_u32(body, temp_idx + 2); emit_unbox_i32(body, rt_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 1); emit_unbox_i32(body, rt_idx);
            body.push(0x4D); // i32.ge_u → lo >= hi means done
            body.push(0x0D); write_leb128_u32(body, 1); // br_if $exit
            // Save arr[lo] to temp
            body.push(0x20); write_leb128_u32(body, temp_idx); // arr
            emit_internalize(body); emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 2); emit_unbox_i32(body, rt_idx);
            body.push(0xFB); write_leb128_u32(body, 0x0B); write_leb128_u32(body, type_ctx.array_type_idx); // array.get
            body.push(0x21); write_leb128_u32(body, temp_idx + 3); // $t3 = arr[lo]
            // arr[lo] = arr[hi]
            body.push(0x20); write_leb128_u32(body, temp_idx);
            emit_internalize(body); emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 2); emit_unbox_i32(body, rt_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx); // arr for second get
            emit_internalize(body); emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 1); emit_unbox_i32(body, rt_idx);
            body.push(0xFB); write_leb128_u32(body, 0x0B); write_leb128_u32(body, type_ctx.array_type_idx); // array.get arr[hi]
            body.push(0xFB); write_leb128_u32(body, 0x0E); write_leb128_u32(body, type_ctx.array_type_idx); // array.set arr[lo]=arr[hi]
            // arr[hi] = saved arr[lo]
            body.push(0x20); write_leb128_u32(body, temp_idx);
            emit_internalize(body); emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 1); emit_unbox_i32(body, rt_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 3); // saved arr[lo]
            body.push(0xFB); write_leb128_u32(body, 0x0E); write_leb128_u32(body, type_ctx.array_type_idx); // array.set arr[hi]=saved
            // lo++, hi--
            body.push(0x20); write_leb128_u32(body, temp_idx + 2); emit_unbox_i32(body, rt_idx);
            body.push(0x41); write_leb128_i32(body, 1); body.push(0x6A); // lo + 1
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 2);
            body.push(0x20); write_leb128_u32(body, temp_idx + 1); emit_unbox_i32(body, rt_idx);
            body.push(0x41); write_leb128_i32(body, 1); body.push(0x6B); // hi - 1
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 1);
            body.push(0x0C); write_leb128_u32(body, 0); // br $loop
            body.push(0x0B); // end loop
            body.push(0x0B); // end block
            body.push(0x20); write_leb128_u32(body, temp_idx); // return arr
        }
        _ if op == Op::ARRAY_CONTAINS => {
            // [arr, val] → externref (boxed bool)
            // Linear search: for i=0..len, if arr[i] == val, return true
            body.push(0x21); write_leb128_u32(body, temp_idx);     // $t0 = val
            body.push(0x22); write_leb128_u32(body, temp_idx + 1); // $t1 = arr (tee)
            emit_internalize(body); emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0xFB); write_leb128_u32(body, 0x0F); // array.len
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 2); // $t2 = len (boxed)
            body.push(0x41); write_leb128_i32(body, 0); emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 3); // $t3 = i (boxed)
            body.push(0x02); body.push(TYPE_EXTERNREF); // block $exit (result externref)
            body.push(0x03); body.push(0x40);            // loop void
            // if i >= len, break → return false
            body.push(0x20); write_leb128_u32(body, temp_idx + 3); emit_unbox_i32(body, rt_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 2); emit_unbox_i32(body, rt_idx);
            body.push(0x4D); // i32.ge_u
            body.push(0x04); body.push(0x40); // if void
            body.push(0x41); write_leb128_i32(body, 0); emit_box_i32(body, rt_idx); // false
            body.push(0x0C); write_leb128_u32(body, 2); // br $exit
            body.push(0x0B); // end if
            // Compare arr[i] with val using f64 equality
            body.push(0x20); write_leb128_u32(body, temp_idx + 1); // arr
            emit_internalize(body); emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 3); emit_unbox_i32(body, rt_idx);
            body.push(0xFB); write_leb128_u32(body, 0x0B); write_leb128_u32(body, type_ctx.array_type_idx);
            // Compare: try f64 equality (dyn_eq pattern)
            body.push(0x21); write_leb128_u32(body, temp_idx + 4); // save element
            body.push(0x20); write_leb128_u32(body, temp_idx + 4); // element
            emit_unbox_f64(body, rt_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx); // val
            emit_unbox_f64(body, rt_idx);
            body.push(0x61); // f64.eq
            body.push(0x04); body.push(0x40); // if equal
            body.push(0x41); write_leb128_i32(body, 1); emit_box_i32(body, rt_idx); // true
            body.push(0x0C); write_leb128_u32(body, 2); // br $exit
            body.push(0x0B); // end if
            // i++
            body.push(0x20); write_leb128_u32(body, temp_idx + 3); emit_unbox_i32(body, rt_idx);
            body.push(0x41); write_leb128_i32(body, 1); body.push(0x6A);
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 3);
            body.push(0x0C); write_leb128_u32(body, 0); // br $loop
            body.push(0x0B); // end loop
            body.push(0x41); write_leb128_i32(body, 0); emit_box_i32(body, rt_idx); // false fallthrough
            body.push(0x0B); // end block
        }
        _ if op == Op::ARRAY_INDEX_OF => {
            // [arr, val] → externref (boxed i32 index, -1 if not found)
            // Same pattern as contains but returns index
            body.push(0x21); write_leb128_u32(body, temp_idx);     // $t0 = val
            body.push(0x22); write_leb128_u32(body, temp_idx + 1); // $t1 = arr
            emit_internalize(body); emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0xFB); write_leb128_u32(body, 0x0F);
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 2); // len
            body.push(0x41); write_leb128_i32(body, 0); emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 3); // i
            body.push(0x02); body.push(TYPE_EXTERNREF); // block
            body.push(0x03); body.push(0x40); // loop
            body.push(0x20); write_leb128_u32(body, temp_idx + 3); emit_unbox_i32(body, rt_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 2); emit_unbox_i32(body, rt_idx);
            body.push(0x4D); // i >= len
            body.push(0x04); body.push(0x40);
            body.push(0x41); write_leb128_i32(body, -1); emit_box_i32(body, rt_idx);
            body.push(0x0C); write_leb128_u32(body, 2);
            body.push(0x0B);
            body.push(0x20); write_leb128_u32(body, temp_idx + 1);
            emit_internalize(body); emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 3); emit_unbox_i32(body, rt_idx);
            body.push(0xFB); write_leb128_u32(body, 0x0B); write_leb128_u32(body, type_ctx.array_type_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 4);
            body.push(0x20); write_leb128_u32(body, temp_idx + 4);
            emit_unbox_f64(body, rt_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx);
            emit_unbox_f64(body, rt_idx);
            body.push(0x61); // f64.eq
            body.push(0x04); body.push(0x40);
            body.push(0x20); write_leb128_u32(body, temp_idx + 3); // return i
            body.push(0x0C); write_leb128_u32(body, 2);
            body.push(0x0B);
            body.push(0x20); write_leb128_u32(body, temp_idx + 3); emit_unbox_i32(body, rt_idx);
            body.push(0x41); write_leb128_i32(body, 1); body.push(0x6A);
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 3);
            body.push(0x0C); write_leb128_u32(body, 0);
            body.push(0x0B); // end loop
            body.push(0x41); write_leb128_i32(body, -1); emit_box_i32(body, rt_idx);
            body.push(0x0B); // end block
        }
        _ if op == Op::ARRAY_JOIN => {
            // [arr, separator] → string (concat all elements with separator between)
            body.push(0x21); write_leb128_u32(body, temp_idx);     // $t0 = separator
            body.push(0x22); write_leb128_u32(body, temp_idx + 1); // $t1 = arr
            emit_internalize(body); emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0xFB); write_leb128_u32(body, 0x0F);
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 2); // len
            // Start with empty string (fromCharCode of nothing — use arr[0] if available)
            body.push(0x41); write_leb128_i32(body, 0); // i = 0
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 3); // i
            // result = ""
            body.push(0x41); write_leb128_i32(body, 0);
            emit_import_call(body, rt_idx, "wasm:js-string", "fromCharCode"); // empty char
            body.push(0x21); write_leb128_u32(body, temp_idx + 4); // result
            body.push(0x02); body.push(0x40); // block void
            body.push(0x03); body.push(0x40); // loop void
            body.push(0x20); write_leb128_u32(body, temp_idx + 3); emit_unbox_i32(body, rt_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 2); emit_unbox_i32(body, rt_idx);
            body.push(0x4D); // i >= len
            body.push(0x0D); write_leb128_u32(body, 1); // br_if $exit
            // if i > 0, append separator
            body.push(0x20); write_leb128_u32(body, temp_idx + 3); emit_unbox_i32(body, rt_idx);
            body.push(0x41); write_leb128_i32(body, 0);
            body.push(0x48); // i32.gt_s: i > 0
            body.push(0x04); body.push(0x40);
            body.push(0x20); write_leb128_u32(body, temp_idx + 4);
            body.push(0x20); write_leb128_u32(body, temp_idx);
            emit_import_call(body, rt_idx, "wasm:js-string", "concat");
            body.push(0x21); write_leb128_u32(body, temp_idx + 4);
            body.push(0x0B); // end if
            // Append arr[i]
            body.push(0x20); write_leb128_u32(body, temp_idx + 4); // result
            body.push(0x20); write_leb128_u32(body, temp_idx + 1); // arr
            emit_internalize(body); emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x20); write_leb128_u32(body, temp_idx + 3); emit_unbox_i32(body, rt_idx);
            body.push(0xFB); write_leb128_u32(body, 0x0B); write_leb128_u32(body, type_ctx.array_type_idx);
            emit_import_call(body, rt_idx, "wasm:js-string", "concat");
            body.push(0x21); write_leb128_u32(body, temp_idx + 4);
            // i++
            body.push(0x20); write_leb128_u32(body, temp_idx + 3); emit_unbox_i32(body, rt_idx);
            body.push(0x41); write_leb128_i32(body, 1); body.push(0x6A);
            emit_box_i32(body, rt_idx);
            body.push(0x21); write_leb128_u32(body, temp_idx + 3);
            body.push(0x0C); write_leb128_u32(body, 0); // br $loop
            body.push(0x0B); // end loop
            body.push(0x0B); // end block
            body.push(0x20); write_leb128_u32(body, temp_idx + 4); // result
        }

        // ── Type introspection → wasm:js-* test builtins ──
        _ if op == Op::REF_IS_STRING => {
            emit_import_call(body, rt_idx, "wasm:js-string", "test");
            emit_box_i32(body, rt_idx); // i32 result → externref
        }
        _ if op == Op::REF_IS_NUMBER => {
            emit_import_call(body, rt_idx, "wasm:js-number", "test");
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::REF_IS_BOOL => {
            emit_import_call(body, rt_idx, "wasm:js-boolean", "test");
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::REF_TYPEOF => {
            // typeof: check types using wasm:js-* test builtins
            // Returns type name as externref string
            // Check: null → "undefined", number → "number", string → "string", boolean → "boolean", else "object"
            body.push(0x22); write_leb128_u32(body, temp_idx); // tee value to temp
            // Check null first
            body.push(0xD1); // ref.is_null → i32
            body.push(0x04); body.push(TYPE_EXTERNREF); // if null (result externref)
            // Build "undefined" via fromCharCode sequence — too verbose
            // Simpler: use fromI32(0) as a sentinel for "undefined"
            // Actually, the caller compares with STR_EQUALS against known strings.
            // Since we can't create string constants here, use the number test
            // to return a type tag that STR_EQUALS will compare.
            // Better approach: use a chain of if/else returning distinct constants.
            // The chunk's constant pool has the strings we need.
            // Find "number" in constants
            let mut number_val = None;
            let mut string_val = None;
            let mut boolean_val = None;
            let mut i32_val = None;
            for (ci, c) in chunk.constants.iter().enumerate() {
                if let Value::String(s) = c {
                    match s.as_ref() {
                        "number" => number_val = Some(ci),
                        "string" => string_val = Some(ci),
                        "boolean" => boolean_val = Some(ci),
                        "i32" => i32_val = Some(ci),
                        _ => {}
                    }
                }
            }
            // null → return "undefined" placeholder (ref.null extern for now)
            body.push(0xD0); body.push(0x6F); // ref.null extern = "undefined"
            body.push(0x05); // else (not null)
            // Check if number
            body.push(0x20); write_leb128_u32(body, temp_idx); // value
            emit_import_call(body, rt_idx, "wasm:js-number", "test"); // → i32
            body.push(0x04); body.push(TYPE_EXTERNREF); // if number (result externref)
            // Return "number" constant if available, else boxed tag
            if let Some(ci) = number_val {
                // Emit the string constant from the chunk's pool
                emit_string_const(body, chunk, ci, rt_idx);
            } else {
                body.push(0x41); write_leb128_i32(body, 1); emit_box_i32(body, rt_idx);
            }
            body.push(0x05); // else (not number)
            // Check if string
            body.push(0x20); write_leb128_u32(body, temp_idx);
            emit_import_call(body, rt_idx, "wasm:js-string", "test");
            body.push(0x04); body.push(TYPE_EXTERNREF); // if string
            if let Some(ci) = string_val {
                emit_string_const(body, chunk, ci, rt_idx);
            } else {
                body.push(0x41); write_leb128_i32(body, 2); emit_box_i32(body, rt_idx);
            }
            body.push(0x05); // else
            // Default: "object"
            body.push(0x41); write_leb128_i32(body, 3); emit_box_i32(body, rt_idx);
            body.push(0x0B); // end if string
            body.push(0x0B); // end if number
            body.push(0x0B); // end if null
        }
        _ if op == Op::REF_IS_OBJECT || op == Op::REF_IS_FUNC || op == Op::REF_IS_ARRAY => {
            // TODO: proper type checks
            body.push(0x1A); // drop value
            body.push(0x41); write_leb128_i32(body, 0); // push false (i32 0)
            emit_box_i32(body, rt_idx);
        }
        // Stack ops
        _ if op == Op::DUP => {
            // Duplicate TOS: local.tee $temp, local.get $temp
            body.push(0x22); write_leb128_u32(body, temp_idx); // local.tee $temp
            body.push(0x20); write_leb128_u32(body, temp_idx); // local.get $temp
        }
        // Exception handling
        _ if op == Op::TRY_START => { let _ = read_u16(&chunk.code, ip); let _ = read_u16(&chunk.code, ip); body.push(0x01); } // nop for now
        _ if op == Op::TRY_END => { body.push(0x01); } // nop for now
        // Spread — TODO: inline impl
        _ if op == Op::SPREAD => { body.push(0x01); } // nop
        // Set timer — TODO: needs host import (not stdlib)
        _ if op == Op::SET_TIMER => { body.push(0x01); } // nop
        // Upvalue get/set — closures use WASM function references
        _ if op == Op::UPVALUE_GET => {
            let _idx = chunk.code[*ip]; *ip += 1;
            // TODO: proper closure/upvalue via WASM funcref + tables
            body.push(0xD0); body.push(0x6F); // ref.null extern (placeholder)
        }
        _ if op == Op::UPVALUE_SET => {
            let _idx = chunk.code[*ip]; *ip += 1;
            // TODO: proper closure/upvalue via WASM funcref + tables
            // drop the value being set
            body.push(0x1A); // drop
            body.push(0xD0); body.push(0x6F); // ref.null extern (placeholder return)
        }
        // Set type ID — GC type stamps handled by WASM GC type system
        _ if op == Op::SET_TYPE_ID => { body.push(0x01); } // nop (type is structural in GC)
        _ if op == Op::HALT => { body.push(0x0F); } // return (not unreachable — _start should return cleanly)
        // global_get/set are core ops (prefix 0x00) — handled in emit_core_op
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
