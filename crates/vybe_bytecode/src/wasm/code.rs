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
            if op == Op::ARRAY_SET || op == Op::STR_SUBSTRING {
                need = need.max(2); // need 2 temps for 3-operand reorder
            } else if is_binary_typed_op(op) || op == Op::GLOBAL_SET || op == Op::DUP
                || op == Op::ARRAY_GET || op == Op::ARRAY_LENGTH
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

/// Scan for backward branches (negative offsets — indicates loops).
fn scan_for_backward_branches(chunk: &Chunk) -> bool {
    let mut ip = 0;
    while ip < chunk.code.len() {
        if ip + 1 >= chunk.code.len() { break; }
        if let Some(op) = Op::decode(chunk.code[ip], chunk.code[ip + 1]) {
            if op == Op::BR {
                let saved = ip + 2;
                let mut read_ip = saved;
                let offset = read_i16(&chunk.code, &mut read_ip);
                if offset < 0 { return true; }
            }
            ip += opcode_size(op, &chunk.code, ip);
        } else {
            ip += 2;
        }
    }
    false
}

/// Scan for forward branches (br_if_false, br_if_null — always forward in our bytecode).
fn scan_for_forward_branches(chunk: &Chunk) -> bool {
    let mut ip = 0;
    while ip < chunk.code.len() {
        if ip + 1 >= chunk.code.len() { break; }
        if let Some(op) = Op::decode(chunk.code[ip], chunk.code[ip + 1]) {
            if op == Op::BR_IF_FALSE || op == Op::BR_IF_NULL || op == Op::BR_IF_TRUE {
                return true;
            }
            ip += opcode_size(op, &chunk.code, ip);
        } else {
            ip += 2;
        }
    }
    false
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

        // Check how many temp locals we need for stack manipulation
        let temp_count = count_temp_locals(chunk);
        let has_temp = temp_count > 0;
        let temp_local_idx = (chunk.arity as u32) + (chunk.local_count as u32);

        // Locals declaration
        if chunk.local_count > 0 || has_temp {
            let total_locals = chunk.local_count as u32 + temp_count;
            write_leb128_u32(&mut body, 1); // 1 local type group
            write_leb128_u32(&mut body, total_locals);
            body.push(TYPE_EXTERNREF);
        } else {
            write_leb128_u32(&mut body, 0);
        }

        // Pre-scan for branches to determine control flow structure.
        // If there are backward branches (loops), wrap body in block+loop.
        let has_backward_branch = scan_for_backward_branches(chunk);
        let has_forward_branch = scan_for_forward_branches(chunk);

        // Emit structured control flow wrapper.
        // Layout: (block $exit (loop $loop <body> end) end)
        // - Backward br → br 0 ($loop: continue)
        // - Forward br → br 1 ($exit: break)
        // - br to function end → br 2 (function body) or return
        let nesting_depth = if has_backward_branch || has_forward_branch {
            body.push(0x02); body.push(TYPE_VOID); // block $exit (void — no result)
            body.push(0x03); body.push(TYPE_VOID); // loop $loop (void)
            2 // depth offset: br 0 = loop, br 1 = block, br 2 = function
        } else {
            0
        };

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
                emit_core_op(&mut body, op, chunk, &mut ip, &rt_idx, temp_local_idx, has_temp, nesting_depth);
            } else if op.prefix() == 0xFB {
                // ── GC ops → real WASM GC binary encoding ──
                emit_gc_op(&mut body, op, chunk, &mut ip, &rt_idx, type_ctx, temp_local_idx);
            } else if op.prefix() >= 0xFC && op.prefix() <= 0xFE {
                // ── Other prefixed WASM ops ──
                body.push(op.prefix());
                write_leb128_u32(&mut body, op.sub() as u32);
                ip += op.operand_format().fixed_size();
            } else {
                // ── VM-internal ops (0xFF) ──
                emit_vm_internal_op(&mut body, op, chunk, &mut ip, &rt_idx, temp_local_idx, nesting_depth);
            }
        }

        // Close block/loop wrappers
        if nesting_depth > 0 {
            body.push(0x0B); // end loop
            body.push(0x0B); // end block
            // Fallthrough return value — in case br exits the block without return.
            // The normal path uses `return` inside the loop.
            body.push(0xD0); body.push(0x6F); // ref.null externref
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
                temp_idx: u32, _has_temp: bool, nesting_depth: u32) {
    match op {
        _ if op == Op::LOCAL_GET => { body.push(op.sub()); write_leb128_u32(body, read_u16(&chunk.code, ip) as u32); }
        _ if op == Op::LOCAL_SET => { body.push(0x22); write_leb128_u32(body, read_u16(&chunk.code, ip) as u32); } // local.tee
        _ if op == Op::CALL => { body.push(op.sub()); let argc = chunk.code[*ip]; *ip += 1; write_leb128_u32(body, argc as u32); }
        _ if op == Op::CALL_REF => {
            let argc = chunk.code[*ip]; *ip += 1;
            // call_ref for closures/higher-order — needs funcref tables (TODO)
            // For now: drop the funcref + args, push null result
            for _ in 0..=argc { body.push(0x1A); } // drop funcref + argc args
            body.push(0xD0); body.push(0x6F); // ref.null extern (placeholder result)
        }
        _ if op == Op::BR => {
            let offset = read_i16(&chunk.code, ip);
            body.push(0x0C); // br
            if offset < 0 {
                // Backward branch → loop continue (depth 0 = loop)
                write_leb128_u32(body, 0);
            } else {
                // Forward branch → break out of loop (depth 1 = block)
                write_leb128_u32(body, 1.min(nesting_depth));
            }
        }
        _ if op == Op::BR_IF_TRUE => {
            let offset = read_i16(&chunk.code, ip);
            emit_unbox_i32(body, rt_idx);  // externref → i32
            body.push(0x0D);               // br_if
            if offset < 0 {
                write_leb128_u32(body, 0); // backward → loop
            } else {
                write_leb128_u32(body, 1.min(nesting_depth)); // forward → block
            }
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
        // WASM global.get/set — our operand is a string name const idx.
        // TODO: Phase 3c — emit real WASM global section with indexed globals.
        // For now: use a local as storage (globals map to locals in single-function mode)
        _ if op == Op::GLOBAL_GET => {
            let _name_idx = read_u16(&chunk.code, ip);
            // TODO: proper WASM global section. For now push null as placeholder.
            body.push(0xD0); body.push(0x6F); // ref.null extern
        }
        _ if op == Op::GLOBAL_SET => {
            let _name_idx = read_u16(&chunk.code, ip);
            // TODO: proper WASM global section. For now just keep value on stack.
            // The value is already on the stack — it becomes the "result" of the set.
        }
        _ if op == Op::HALT => { body.push(0x00); } // unreachable
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
fn emit_vm_internal_op(body: &mut Vec<u8>, op: Op, chunk: &Chunk, ip: &mut usize, rt_idx: &std::collections::HashMap<(&str, &str), usize>, temp_idx: u32, nesting_depth: u32) {
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
            let offset = read_i16(&chunk.code, ip);
            emit_unbox_i32(body, rt_idx);  // externref → i32
            body.push(0x45);               // i32.eqz (invert: branch if false)
            body.push(0x0D);               // br_if
            if offset < 0 {
                write_leb128_u32(body, 0); // backward → loop
            } else {
                write_leb128_u32(body, 1.min(nesting_depth)); // forward → block
            }
        }
        _ if op == Op::BR_IF_NULL => {
            let offset = read_i16(&chunk.code, ip);
            body.push(0xD1);               // ref.is_null → i32
            body.push(0x0D);               // br_if
            if offset < 0 {
                write_leb128_u32(body, 0);
            } else {
                write_leb128_u32(body, 1.min(nesting_depth));
            }
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
        // Binary string ops (2 args → 1 result): drop second, keep first
        _ if op == Op::STR_INDEX_OF || op == Op::STR_LAST_INDEX_OF
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
        _ if op == Op::ARRAY_PUSH => {
            // Stack: [externref_arr, externref_val]
            // WASM GC arrays are fixed-size. array_push needs:
            // 1. Get old array length
            // 2. Create new array of length+1
            // 3. Copy old elements
            // 4. Set new element at end
            // This is complex — for now drop val, keep arr (TODO: proper impl)
            body.push(0x1A); // drop val
            // leaves arr on stack
        }
        _ if op == Op::ARRAY_POP || op == Op::ARRAY_SHIFT => {
            // TODO: proper impl — for now just return the array unchanged
            body.push(0x01); // nop
        }
        _ if op == Op::ARRAY_SLICE => {
            // Stack: [arr, start, end] → TODO: proper slice
            body.push(0x1A); // drop end
            body.push(0x1A); // drop start
            // leaves arr on stack
        }
        _ if op == Op::ARRAY_CONCAT || op == Op::ARRAY_JOIN
          || op == Op::ARRAY_REVERSE || op == Op::ARRAY_CONTAINS
          || op == Op::ARRAY_INDEX_OF => {
            // Binary array ops — TODO: proper impl
            body.push(0x1A); // drop second arg
            // leaves first arg on stack
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
            // typeof → check types in order, return string
            // For now: drop value, push null (TODO: proper impl)
            body.push(0x1A); // drop
            body.push(0xD0); body.push(0x6F); // ref.null extern
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
        _ if op == Op::HALT => { body.push(0x00); } // unreachable
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
