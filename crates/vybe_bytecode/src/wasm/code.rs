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
use super::sections::{
    emit_box_f64, emit_box_i32, emit_import_call, emit_unbox_f64, emit_unbox_i32,
};
use super::types::WasmTypeContext;
use crate::opcode::{OperandFormat, read_leb_u32};
use crate::value::Value;
use crate::{Chunk, Op};

/// A try/catch region extracted from the bytecode.
///
/// The compiler emits:
/// ```text
/// TRY_START catch_off fin_off  ; 6 bytes
///   ...body...
/// TRY_END                       ; 2 bytes
/// [else body]
/// BR skip_to_finally            ; 4 bytes; target = after_ip
///   ...catch dispatch + body...
/// after_ip:
/// ```
///
/// We translate that to a structural WASM try_table:
/// ```wasm
/// block $after (void)
///   block $catch (result externref)
///     try_table (catch $vybe_exception $catch)
///       ...body...
///     end
///     br $after            ;; skip catch on normal path
///   end $catch             ;; exception externref on stack (from tag param)
///   ...catch handler...
/// end $after
/// ```
#[derive(Clone, Copy)]
struct TryRegion {
    /// Byte offset of the TRY_START opcode in the source chunk.
    try_start_pos: usize,
    /// Byte offset of the matching TRY_END opcode.
    try_end_pos: usize,
    /// Byte offset where catch-dispatch code begins (just after the
    /// compiler-emitted BR that skips catch on success).
    catch_ip: usize,
    /// Byte offset where normal control flow resumes after the whole
    /// try region (target of the skip-BR).
    after_ip: usize,
}

/// Walk the bytecode collecting try regions keyed by TRY_START offset.
///
/// Each catch_ip must be preceded by a 4-byte BR instruction; we use
/// that BR's target to locate `after_ip`. If the BR isn't where we
/// expect, the region is skipped (emitter falls back to nop-ing the
/// try markers, preserving pre-exception-handling behavior).
fn collect_try_regions(chunk: &Chunk) -> std::collections::HashMap<usize, TryRegion> {
    let mut regions = std::collections::HashMap::new();
    let mut ip = 0;
    while ip + 1 < chunk.code.len() {
        let Some(op) = Op::decode(chunk.code[ip], chunk.code[ip + 1] as u16) else {
            ip += 2;
            continue;
        };
        if op == Op::TRY_START {
            let op_pos = ip;
            // Read i16 catch_offset immediately after the 2-byte opcode.
            let catch_off = ((chunk.code[ip + 2] as i16) << 8) | (chunk.code[ip + 3] as i16 & 0xFF);
            // finally_offset at ip+4..ip+6 — reserved, unused for now.
            let operands_end = ip + 6;
            // catch_ip is relative to the byte *after* the 4 operand
            // bytes (the VM reads both u16s before adding the offset).
            let catch_ip = (operands_end as i64 + catch_off as i64) as usize;
            ip = operands_end;

            // Find the matching TRY_END (nested TRY_STARTs count).
            let mut depth = 1i32;
            let mut try_end_pos: Option<usize> = None;
            let mut scan = ip;
            while scan + 1 < chunk.code.len() {
                let Some(inner) = Op::decode(chunk.code[scan], chunk.code[scan + 1] as u16) else {
                    scan += 2;
                    continue;
                };
                if inner == Op::TRY_START {
                    depth += 1;
                    scan += opcode_size(inner, &chunk.code, scan);
                    continue;
                }
                if inner == Op::TRY_END {
                    depth -= 1;
                    if depth == 0 {
                        try_end_pos = Some(scan);
                        break;
                    }
                    scan += opcode_size(inner, &chunk.code, scan);
                    continue;
                }
                scan += opcode_size(inner, &chunk.code, scan);
            }

            if let Some(te_pos) = try_end_pos {
                // Verify the 4 bytes before catch_ip are a BR opcode.
                if catch_ip >= 4 && catch_ip <= chunk.code.len() {
                    let br_pos = catch_ip - 4;
                    let is_br = chunk.code[br_pos] == Op::BR.prefix()
                        && chunk.code[br_pos + 1] == Op::BR.sub_u8();
                    if is_br {
                        let br_off = ((chunk.code[br_pos + 2] as i16) << 8)
                            | (chunk.code[br_pos + 3] as i16 & 0xFF);
                        let after_ip = (catch_ip as i64 + br_off as i64) as usize;
                        regions.insert(
                            op_pos,
                            TryRegion {
                                try_start_pos: op_pos,
                                try_end_pos: te_pos,
                                catch_ip,
                                after_ip,
                            },
                        );
                    }
                }
            }
        } else {
            ip += opcode_size(op, &chunk.code, ip);
        }
    }
    regions
}

/// Byte size of the LEB128-encoded value (unsigned).
fn leb128_u32_size(v: u32) -> u32 {
    if v < 0x80 {
        1
    } else if v < 0x4000 {
        2
    } else if v < 0x20_0000 {
        3
    } else if v < 0x1000_0000 {
        4
    } else {
        5
    }
}

/// Size in bytes of the locals declaration prefix written at the start of a
/// function body. The spec says branch hint offsets are relative to this
/// position, so the hint scanner adds this value to every `chunk.code` offset.
pub(crate) fn locals_prefix_size(chunk: &Chunk) -> u32 {
    let wasm_params = chunk.arity as u32;
    let extra_locals = if chunk.local_count as u32 > wasm_params {
        chunk.local_count as u32 - wasm_params
    } else {
        0
    };
    let temp_count = count_temp_locals(chunk);
    let declared_locals = extra_locals + temp_count;
    if declared_locals > 0 {
        // 1 group count (1 byte) + leb128(declared_locals) + 1 type byte
        1 + leb128_u32_size(declared_locals) + 1
    } else {
        // vec-count of 0 (1 byte for LEB128 encoding of 0)
        1
    }
}

/// Count how many temp locals a chunk needs for stack manipulation.
/// Returns 0, 1, or 2 depending on which ops are used.
fn count_temp_locals(chunk: &Chunk) -> u32 {
    let mut need = 0u32;
    let mut ip = 0;
    while ip < chunk.code.len() {
        if ip + 1 >= chunk.code.len() {
            break;
        }
        if let Some(op) = Op::decode(chunk.code[ip], chunk.code[ip + 1] as u16) {
            if op == Op::CALL_REF {
                // call_ref needs argc+1 temps (save args + table idx)
                let call_argc = chunk.code.get(ip + 2).copied().unwrap_or(0) as u32;
                need = need.max(call_argc + 1);
            } else if op == Op::STR_INDEX_OF {
                need = need.max(5); // need 5 temps
            } else if op == Op::ARRAY_SET || op == Op::STRUCT_SET || op == Op::STR_SUBSTRING {
                need = need.max(2); // need 2 temps for 3-operand reorder
            } else if is_binary_typed_op(op)
                || op == Op::GLOBAL_SET
                || op == Op::DUP
                || op == Op::ARRAY_GET
                || op == Op::ARRAY_LENGTH
                || op == Op::REF_TYPEOF
                || op == Op::REF_IS_NULL
            {
                need = need.max(1);
            }
            // Phase E: the `0xFF` ARRAY_* (PUSH/POP/SLICE/JOIN/REVERSE/
            // CONTAINS/INDEX_OF/CONTAINS/CONCAT/SHIFT) no longer exist.
            // Their call-site temps are no longer needed here either.
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

fn next_bytes_decode_opcode(chunk: &Chunk, ip: usize) -> bool {
    ip + 1 < chunk.code.len() && Op::decode(chunk.code[ip], chunk.code[ip + 1] as u16).is_some()
}

fn read_optional_memidx_immediate(chunk: &Chunk, ip: &mut usize) -> u32 {
    if chunk.code.get(*ip) == Some(&0xEE) && chunk.code.get(*ip + 1) == Some(&0x00) {
        *ip += 2;
        return read_leb_u32(&chunk.code, ip);
    }
    if next_bytes_decode_opcode(chunk, *ip) {
        return 0;
    }
    read_leb_u32(&chunk.code, ip)
}

fn read_optional_memarg(chunk: &Chunk, ip: &mut usize, default_align: u32) -> (u32, u64, u32) {
    if next_bytes_decode_opcode(chunk, *ip) {
        return (default_align, 0, 0);
    }
    let align = read_leb_u32(&chunk.code, ip);
    let offset = read_leb_u32(&chunk.code, ip) as u64;
    let memidx = if align & 0x40 != 0 {
        read_leb_u32(&chunk.code, ip)
    } else {
        0
    };
    (align & !0x40, offset, memidx)
}

fn read_optional_memarg64(chunk: &Chunk, ip: &mut usize, default_align: u32) -> (u32, u64, u32) {
    if next_bytes_decode_opcode(chunk, *ip) {
        return (default_align, 0, 0);
    }
    let align = read_leb_u32(&chunk.code, ip);
    let offset = read_leb_u64(&chunk.code, ip);
    let memidx = if align & 0x40 != 0 {
        read_leb_u32(&chunk.code, ip)
    } else {
        0
    };
    (align & !0x40, offset, memidx)
}

fn emit_stack_switch_handlers(body: &mut Vec<u8>, chunk: &Chunk, op_start: usize) {
    if let Some(handlers) = chunk.stack_switch_handlers.get(&op_start) {
        write_leb128_u32(body, handlers.len() as u32);
        for handler in handlers {
            body.push(handler.kind);
            write_leb128_u32(body, handler.tag_index);
            if handler.kind == 0 {
                write_leb128_u32(body, handler.label_index);
            }
        }
    } else {
        write_leb128_u32(body, 0);
    }
}

fn wasm_struct_type_for_chunk_type(chunk: &Chunk, type_ctx: &WasmTypeContext, typeidx: u16) -> u32 {
    chunk
        .types
        .get(typeidx as usize)
        .and_then(|ty| type_ctx.struct_type(&ty.name))
        .unwrap_or(typeidx as u32)
}

fn wasm_struct_type_matching_field_count(
    chunk: &Chunk,
    type_ctx: &WasmTypeContext,
    field_count: u16,
) -> u32 {
    let mut matches = chunk
        .types
        .iter()
        .filter(|ty| ty.fields.len() == field_count as usize)
        .filter_map(|ty| type_ctx.struct_type(&ty.name));
    let first = matches.next();
    if matches.next().is_none() {
        first.unwrap_or(0)
    } else {
        0
    }
}

fn wasm_struct_field_for_name(
    chunk: &Chunk,
    type_ctx: &WasmTypeContext,
    field_name_idx: u16,
) -> (u32, u32) {
    let Some(value) = chunk.constants.get(field_name_idx as usize) else {
        return (0, 0);
    };
    let field_name = format!("{}", value).to_ascii_lowercase();
    let mut matches = chunk.types.iter().filter_map(|ty| {
        let field_idx = ty
            .fields
            .iter()
            .position(|field| field.eq_ignore_ascii_case(&field_name))?;
        Some((
            type_ctx.struct_type(&ty.name).unwrap_or(0),
            field_idx as u32,
        ))
    });
    let first = matches.next();
    if matches.next().is_none() {
        first.unwrap_or((0, 0))
    } else {
        (0, 0)
    }
}

fn read_leb_u64(code: &[u8], ip: &mut usize) -> u64 {
    let mut result = 0u64;
    let mut shift = 0u32;
    while *ip < code.len() && shift < 64 {
        let byte = code[*ip];
        *ip += 1;
        result |= ((byte & 0x7f) as u64) << shift;
        if byte & 0x80 == 0 {
            break;
        }
        shift += 7;
    }
    result
}

pub fn encode_code_section(
    chunks: &[Chunk],
    rt_imports: &[(&str, &str)],
    type_ctx: &WasmTypeContext,
    global_map: &std::collections::HashMap<String, u32>,
) -> Vec<u8> {
    let host_import_count = chunks.first().map(|c| c.imports.len()).unwrap_or(0);

    // Build import name → function index map (module:name for uniqueness)
    let mut rt_idx: std::collections::HashMap<(&str, &str), usize> =
        std::collections::HashMap::new();
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
        } else {
            0
        };
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

        // Structured control flow: the compiler now emits BLOCK/LOOP/END/BR/BR_IF.
        // The WASM emitter just passes them through — no relooper needed.

        // Pre-pass: identify try regions so we can wrap them in proper
        // structural WASM try_table blocks (exception-handling proposal).
        // Five events, keyed by bytecode offset, drive the emission:
        //   try_start → emit `block $after; block $catch; try_table …`
        //   try_end   → close try_table (an `else` body, if any, follows
        //               normally and runs inside $catch, unprotected)
        //   skip_br   → rewrite the compiler's BR to `br 1` (to $after)
        //   catch_ip  → close $catch (exception externref on stack)
        //   after_ip  → close $after
        let try_regions = collect_try_regions(chunk);
        let mut try_start_events: std::collections::HashMap<usize, TryRegion> =
            std::collections::HashMap::new();
        let mut try_end_events: std::collections::HashSet<usize> = std::collections::HashSet::new();
        let mut skip_br_events: std::collections::HashSet<usize> = std::collections::HashSet::new();
        let mut catch_ip_events: std::collections::HashSet<usize> =
            std::collections::HashSet::new();
        let mut after_ip_events: std::collections::HashMap<usize, usize> =
            std::collections::HashMap::new();
        for (_pos, region) in &try_regions {
            try_start_events.insert(region.try_start_pos, *region);
            try_end_events.insert(region.try_end_pos);
            skip_br_events.insert(region.catch_ip - 4);
            catch_ip_events.insert(region.catch_ip);
            *after_ip_events.entry(region.after_ip).or_insert(0) += 1;
        }

        // Translate opcodes
        let mut ip = 0;
        while ip < chunk.code.len() {
            // Emit the $catch end marker when we land on a catch_ip.
            if catch_ip_events.contains(&ip) {
                body.push(0x0B); // end $catch — exception externref on stack
            }
            // Emit any $after closes — at after_ip the whole try region
            // is done. Nested regions may pile up here.
            if let Some(n) = after_ip_events.get(&ip) {
                for _ in 0..*n {
                    body.push(0x0B);
                }
            }

            if ip + 1 >= chunk.code.len() {
                break;
            }
            let op = match Op::decode(chunk.code[ip], chunk.code[ip + 1] as u16) {
                Some(op) => op,
                None => {
                    ip += 2;
                    continue;
                }
            };

            // TRY_START → open structural blocks.
            if op == Op::TRY_START && try_start_events.contains_key(&ip) {
                body.push(0x02);
                body.push(TYPE_VOID); // block $after
                body.push(0x02);
                body.push(TYPE_EXTERNREF); // block $catch (result externref)
                body.push(0x1F); // try_table
                body.push(TYPE_VOID); // block type: void
                write_leb128_u32(&mut body, 1); // 1 catch clause
                body.push(0x00); // variant: catch (tag, label)
                write_leb128_u32(&mut body, super::exception_handling::VYBE_EXCEPTION_TAG);
                write_leb128_u32(&mut body, 0); // label 0 = $catch
                ip += 6;
                continue;
            }
            // TRY_END → close try_table. An `else` body (Python/Ruby) runs
            // inside $catch, unprotected, between this point and skip_br.
            if op == Op::TRY_END && try_end_events.contains(&ip) {
                body.push(0x0B); // end (closes try_table)
                ip += 2;
                continue;
            }
            // Skip-BR (compiler-emitted BR that jumps over catch on the
            // success path) → replace with `br $after`.
            if op == Op::BR && skip_br_events.contains(&ip) {
                body.push(0x0C);
                write_leb128_u32(&mut body, 1);
                ip += 4;
                continue;
            }
            let op_start = ip;
            ip += 2;

            if op.prefix() == 0x00 && !op.is_vm_internal() {
                emit_core_op(
                    &mut body,
                    op,
                    chunk,
                    &mut ip,
                    op_start,
                    &rt_idx,
                    temp_local_idx,
                    has_temp,
                    type_ctx,
                    global_map,
                    host_import_count,
                );
            } else if op.prefix() == 0xFB {
                emit_gc_op(
                    &mut body,
                    op,
                    chunk,
                    &mut ip,
                    &rt_idx,
                    type_ctx,
                    temp_local_idx,
                );
            } else if op.prefix() == 0xFC {
                // 0xFC-prefix ops per the bulk-memory / reference-types spec
                // need specific trailing immediates that aren't captured by
                // our `operand_format` in bytecode. Translate each case.
                body.push(op.prefix());
                write_leb128_u32(&mut body, op.sub() as u32);
                match op {
                    Op::MEMORY_INIT => {
                        // spec: data_idx, memory_idx
                        let data_idx = chunk.code[ip];
                        ip += 1;
                        let memidx = read_optional_memidx_immediate(chunk, &mut ip);
                        write_leb128_u32(&mut body, data_idx as u32);
                        write_leb128_u32(&mut body, memidx);
                    }
                    Op::DATA_DROP => {
                        let data_idx = chunk.code[ip];
                        ip += 1;
                        write_leb128_u32(&mut body, data_idx as u32);
                    }
                    Op::MEMORY_COPY => {
                        // spec: dst_mem, src_mem
                        let dst_mem = read_optional_memidx_immediate(chunk, &mut ip);
                        let src_mem = read_optional_memidx_immediate(chunk, &mut ip);
                        write_leb128_u32(&mut body, dst_mem);
                        write_leb128_u32(&mut body, src_mem);
                    }
                    Op::MEMORY_FILL => {
                        let memidx = read_optional_memidx_immediate(chunk, &mut ip);
                        write_leb128_u32(&mut body, memidx);
                    }
                    Op::TABLE_INIT => {
                        let elem_idx = chunk.code[ip];
                        ip += 1;
                        let table_idx = chunk.code[ip];
                        ip += 1;
                        write_leb128_u32(&mut body, elem_idx as u32);
                        write_leb128_u32(&mut body, table_idx as u32);
                    }
                    Op::ELEM_DROP => {
                        let elem_idx = chunk.code[ip];
                        ip += 1;
                        write_leb128_u32(&mut body, elem_idx as u32);
                    }
                    Op::TABLE_COPY => {
                        let dst_table = chunk.code[ip];
                        ip += 1;
                        let src_table = chunk.code[ip];
                        ip += 1;
                        write_leb128_u32(&mut body, dst_table as u32);
                        write_leb128_u32(&mut body, src_table as u32);
                    }
                    Op::TABLE_GROW | Op::TABLE_SIZE | Op::TABLE_FILL => {
                        let table_idx = chunk.code[ip];
                        ip += 1;
                        write_leb128_u32(&mut body, table_idx as u32);
                    }
                    _ => {
                        ip += op.operand_format().size_in(&chunk.code, ip);
                    }
                }
            } else if op.prefix() == 0xFD {
                emit_simd_prefixed_op(&mut body, op, chunk, &mut ip);
            } else if op.prefix() == 0xFE {
                emit_thread_prefixed_op(&mut body, op, chunk, &mut ip);
            } else {
                emit_vm_internal_op(
                    &mut body,
                    op,
                    chunk,
                    &mut ip,
                    op_start,
                    &rt_idx,
                    temp_local_idx,
                    type_ctx,
                );
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
fn emit_core_op(
    body: &mut Vec<u8>,
    op: Op,
    chunk: &Chunk,
    ip: &mut usize,
    op_start: usize,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    temp_idx: u32,
    _has_temp: bool,
    type_ctx: &WasmTypeContext,
    global_map: &std::collections::HashMap<String, u32>,
    _host_import_count: usize,
) {
    match op {
        _ if op == Op::LOCAL_GET => {
            body.push(op.sub_u8());
            write_leb128_u32(body, read_u16(&chunk.code, ip) as u32);
        }
        _ if op == Op::LOCAL_SET => {
            body.push(0x22);
            write_leb128_u32(body, read_u16(&chunk.code, ip) as u32);
        } // local.tee
        _ if op == Op::CALL => {
            body.push(op.sub_u8());
            let func_idx = read_u16(&chunk.code, ip);
            let _argc = chunk.code[*ip];
            *ip += 1;
            write_leb128_u32(body, func_idx as u32);
        }
        _ if op == Op::CALL_REF => {
            let argc = chunk.code[*ip];
            *ip += 1;
            // Stack: [externref_funcref, arg1, ..., argN] — funcref is below args
            // call_indirect needs: [arg1, ..., argN, i32_table_idx]
            // WASM convention: slot 0 = first arg, no reserved callee slot.
            //
            // 1. Save all args to temps
            for i in (0..argc).rev() {
                body.push(0x21);
                write_leb128_u32(body, temp_idx + i as u32);
            }
            // Stack: [externref_funcref]
            // 2. Save funcref
            body.push(0x21);
            write_leb128_u32(body, temp_idx + argc as u32);
            // 3. Restore user args
            for i in 0..argc {
                body.push(0x20);
                write_leb128_u32(body, temp_idx + i as u32);
            }
            // 4. Push table index (unbox funcref to i32)
            body.push(0x20);
            write_leb128_u32(body, temp_idx + argc as u32);
            emit_unbox_i32(body, rt_idx);
            // 5. call_indirect with matching function type
            if let Some(&type_idx) = type_ctx.func_type_by_arity.get(&argc) {
                body.push(0x11); // call_indirect
                write_leb128_u32(body, type_idx); // type index
                write_leb128_u32(body, 0); // table index 0
            } else {
                // No matching type — drop everything, push null
                body.push(0x1A); // drop table_idx
                for _ in 0..argc {
                    body.push(0x1A);
                }
                body.push(0xD0);
                body.push(0x6F);
            }
        }
        _ if op == Op::CALL_INDIRECT => {
            let argc = chunk.code[*ip];
            *ip += 1;
            let table_idx = chunk.code[*ip];
            *ip += 1;
            if let Some(&type_idx) = type_ctx.func_type_by_arity.get(&argc) {
                body.push(0x11);
                write_leb128_u32(body, type_idx);
                write_leb128_u32(body, table_idx as u32);
            } else {
                body.push(0x1A); // drop table_idx
                for _ in 0..argc {
                    body.push(0x1A);
                }
                body.push(0xD0);
                body.push(0x6F);
            }
        }
        _ if op == Op::BR => {
            let depth = read_leb_u32(&chunk.code, ip);
            body.push(0x0C);
            write_leb128_u32(body, depth);
        }
        _ if op == Op::BR_IF => {
            let depth = read_leb_u32(&chunk.code, ip);
            body.push(0x0D);
            write_leb128_u32(body, depth);
        }
        _ if op == Op::BR_TABLE => {
            let count = read_leb_u32(&chunk.code, ip);
            body.push(0x0E);
            write_leb128_u32(body, count);
            for _ in 0..count {
                let depth = read_leb_u32(&chunk.code, ip);
                write_leb128_u32(body, depth);
            }
            let default_depth = read_leb_u32(&chunk.code, ip);
            write_leb128_u32(body, default_depth);
        }
        // END pops a label from the structured CF stack
        _ if op == Op::END => {
            body.push(0x0B); // end
        }
        // BLOCK/LOOP/IF carry a result_count byte (0=void, 1=externref, N≥2=type-idx).
        // Translate to WASM blocktype (negative valtype or positive type-index LEB128).
        _ if op == Op::BLOCK || op == Op::LOOP || op == Op::IF => {
            let result_count = chunk.code[*ip];
            *ip += 1;
            body.push(op.sub_u8()); // 0x02 / 0x03 / 0x04
            match result_count {
                0 => body.push(TYPE_VOID),
                1 => body.push(TYPE_EXTERNREF),
                n => {
                    let tidx = *type_ctx
                        .block_type_by_results
                        .get(&n)
                        .expect("block multi-value type was not pre-registered");
                    write_leb128_i32(body, tidx as i32);
                }
            }
        }
        // ELSE: no operands.
        _ if op == Op::ELSE => {
            body.push(0x05); // else
        }
        _ if op == Op::MEMORY_SIZE || op == Op::MEMORY_GROW => {
            body.push(op.sub_u8());
            let memidx = read_optional_memidx_immediate(chunk, ip);
            write_leb128_u32(body, memidx);
        }
        // Memory load/store with alignment + offset
        _ if op == Op::I32_LOAD || op == Op::F32_LOAD => {
            body.push(op.sub_u8());
            let (align, offset, memidx) = read_optional_memarg(chunk, ip, 2);
            encode_memarg_with_memidx(body, align, offset, memidx);
        }
        _ if op == Op::I64_LOAD || op == Op::F64_LOAD => {
            body.push(op.sub_u8());
            let (align, offset, memidx) = read_optional_memarg(chunk, ip, 3);
            encode_memarg_with_memidx(body, align, offset, memidx);
        }
        _ if op == Op::I32_LOAD8_S
            || op == Op::I32_LOAD8_U
            || op == Op::I64_LOAD8_S
            || op == Op::I64_LOAD8_U =>
        {
            body.push(op.sub_u8());
            let (align, offset, memidx) = read_optional_memarg(chunk, ip, 0);
            encode_memarg_with_memidx(body, align, offset, memidx);
        }
        _ if op == Op::I32_LOAD16_S
            || op == Op::I32_LOAD16_U
            || op == Op::I64_LOAD16_S
            || op == Op::I64_LOAD16_U =>
        {
            body.push(op.sub_u8());
            let (align, offset, memidx) = read_optional_memarg(chunk, ip, 1);
            encode_memarg_with_memidx(body, align, offset, memidx);
        }
        _ if op == Op::I64_LOAD32_S || op == Op::I64_LOAD32_U => {
            body.push(op.sub_u8());
            let (align, offset, memidx) = read_optional_memarg(chunk, ip, 2);
            encode_memarg_with_memidx(body, align, offset, memidx);
        }
        _ if op == Op::I32_STORE || op == Op::F32_STORE => {
            body.push(op.sub_u8());
            let (align, offset, memidx) = read_optional_memarg(chunk, ip, 2);
            encode_memarg_with_memidx(body, align, offset, memidx);
        }
        _ if op == Op::I64_STORE || op == Op::F64_STORE => {
            body.push(op.sub_u8());
            let (align, offset, memidx) = read_optional_memarg(chunk, ip, 3);
            encode_memarg_with_memidx(body, align, offset, memidx);
        }
        _ if op == Op::I32_STORE8 || op == Op::I64_STORE8 => {
            body.push(op.sub_u8());
            let (align, offset, memidx) = read_optional_memarg(chunk, ip, 0);
            encode_memarg_with_memidx(body, align, offset, memidx);
        }
        _ if op == Op::I32_STORE16 || op == Op::I64_STORE16 => {
            body.push(op.sub_u8());
            let (align, offset, memidx) = read_optional_memarg(chunk, ip, 1);
            encode_memarg_with_memidx(body, align, offset, memidx);
        }
        _ if op == Op::I64_STORE32 => {
            body.push(op.sub_u8());
            let (align, offset, memidx) = read_optional_memarg(chunk, ip, 2);
            encode_memarg_with_memidx(body, align, offset, memidx);
        }
        // WASM global.get/set — resolved to indexed globals via global_map.
        // WASM global indices are module-wide: imported globals come first
        // (indices 0..imported_globals_count), user globals follow. The
        // `rt_globals()` list (UNDEFINED/TRUE/FALSE from
        // js-primitive-builtins) accounts for the 3 imported ones, so
        // user-global idx N in the bytecode maps to WASM idx N +
        // imported_globals_count. Without this offset, `global.set 0`
        // would target the first imported `wasm:js-*` global (immutable)
        // and v8 correctly rejects it.
        _ if op == Op::GLOBAL_GET => {
            let name_idx = read_u16(&chunk.code, ip);
            if let Some(crate::value::Value::String(name)) = chunk.constants.get(name_idx as usize)
            {
                if let Some(&gidx) = global_map.get(name.as_ref()) {
                    let wasm_gidx = gidx + crate::wasm::sections::rt_globals().len() as u32;
                    body.push(0x23); // global.get
                    write_leb128_u32(body, wasm_gidx);
                } else {
                    body.push(0xD0);
                    body.push(0x6F); // ref.null extern (unknown global)
                }
            } else {
                body.push(0xD0);
                body.push(0x6F);
            }
        }
        _ if op == Op::GLOBAL_SET => {
            let name_idx = read_u16(&chunk.code, ip);
            if let Some(crate::value::Value::String(name)) = chunk.constants.get(name_idx as usize)
            {
                if let Some(&gidx) = global_map.get(name.as_ref()) {
                    let wasm_gidx = gidx + crate::wasm::sections::rt_globals().len() as u32;
                    // Stack has [value]. global.set consumes it — but our VM keeps it.
                    // Use local.tee pattern: tee to keep value, then global.set
                    body.push(0x22);
                    write_leb128_u32(body, temp_idx); // local.tee $temp
                    body.push(0x24); // global.set
                    write_leb128_u32(body, wasm_gidx);
                    body.push(0x20);
                    write_leb128_u32(body, temp_idx); // restore value
                } else {
                    // Unknown global — just keep value on stack
                }
            }
        }
        _ if op == Op::HALT => {
            body.push(0x0F);
        } // return (not unreachable — _start should return cleanly)
        // Exception-handling proposal. THROW takes the exception value
        // from TOS and raises it via the single `$vybe_exception` tag
        // (declared in the tag section). The tag's signature is
        // `(externref) -> ()`, so the throw consumes the externref
        // already on the stack — no additional packing needed.
        //
        // THROW_REF is spec'd to take an `exnref` off the stack; since
        // our value type is externref (not exnref), we re-throw through
        // the same tag to stay within the single-tag design.
        _ if op == Op::THROW => {
            body.push(0x08); // throw
            write_leb128_u32(body, super::exception_handling::VYBE_EXCEPTION_TAG);
        }
        _ if op == Op::THROW_REF => {
            body.push(0x08);
            write_leb128_u32(body, super::exception_handling::VYBE_EXCEPTION_TAG);
        }
        // Reference-types `table.get tbl` / `table.set tbl` (core prefix).
        // Bytecode carries a single-byte table index; WASM binary uses a
        // LEB128 tableidx, so we widen on the way out.
        _ if op == Op::TABLE_GET => {
            let tbl = chunk.code[*ip];
            *ip += 1;
            body.push(0x25);
            write_leb128_u32(body, tbl as u32);
        }
        _ if op == Op::TABLE_SET => {
            let tbl = chunk.code[*ip];
            *ip += 1;
            body.push(0x26);
            write_leb128_u32(body, tbl as u32);
        }
        // Typed `select t` (0x1C): same stack semantics as untyped
        // `select` but carries a `vec(valtype)` operand. Our uniform ABI
        // always selects among externref values, so we emit the canonical
        // `[1 × externref]` result-type vector inline.
        _ if op == Op::SELECT_T => {
            body.push(0x1C);
            write_leb128_u32(body, 1); // 1 result type
            body.push(TYPE_EXTERNREF);
        }
        // `ref.null extern` — core opcode path. Typed variants
        // (NULL_FUNC / NULL_ANY / NULL_NONE) use the 0xFF prefix and
        // lower in `emit_vm_internal_op`.
        _ if op == Op::NULL => {
            body.push(0xD0);
            body.push(HT_EXTERN);
        }
        // ref.is_null produces i32 — box it since our value representation is externref
        _ if op == Op::REF_IS_NULL => {
            body.push(0xD1); // ref.is_null → i32
            emit_box_i32(body, rt_idx); // i32 → externref
        }
        // GC proposal (core prefix): ref.eq produces i32 — rebox as externref.
        _ if op == Op::REF_EQ => {
            body.push(0xD3);
            emit_box_i32(body, rt_idx);
        }
        // ref.as_non_null is identity at the WASM level — it only
        // distinguishes a non-null reference type at validation time.
        // Since our values are externref (nullable) throughout, emit
        // the opcode directly; the engine will trap on null per spec.
        _ if op == Op::REF_AS_NON_NULL => {
            body.push(0xD4);
        }
        // br_on_null / br_on_non_null take an LEB128 u32 label immediate.
        // Our bytecode stores a 2-byte signed offset; we cannot turn an
        // offset into a structural label depth here (same gap as BR/BR_IF),
        // so we emit the spec-correct byte with a conservative label 0.
        // Until structural CF is fully wired, this is equivalent to the
        // existing BR treatment: producing spec-visible bytes rather than
        // silent nops.
        _ if op == Op::BR_ON_NULL => {
            let _offset = read_i16(&chunk.code, ip);
            body.push(0xD5);
            write_leb128_u32(body, 0);
        }
        _ if op == Op::BR_ON_NON_NULL => {
            let _offset = read_i16(&chunk.code, ip);
            body.push(0xD6);
            write_leb128_u32(body, 0);
        }

        // ── f64 binary arithmetic: unbox both → f64 op → rebox ──
        // Stack: [externref_a, externref_b]
        // Pattern: local.set $temp (save b) → toF64 (unbox a) → local.get $temp → toF64 (unbox b)
        //          → f64.op → fromF64 (rebox result)
        _ if op == Op::F64_ADD
            || op == Op::F64_SUB
            || op == Op::F64_MUL
            || op == Op::F64_DIV
            || op == Op::F64_MIN
            || op == Op::F64_MAX
            || op == Op::F64_COPYSIGN =>
        {
            emit_binary_f64_op(body, op.sub_u8(), rt_idx, temp_idx);
        }

        // ── f64 comparisons: unbox both → compare → rebox i32 result ──
        _ if op == Op::F64_LT || op == Op::F64_GT || op == Op::F64_LE || op == Op::F64_GE => {
            emit_binary_f64_cmp(body, op.sub_u8(), rt_idx, temp_idx);
        }

        // ── i32 binary arithmetic: unbox both → i32 op → rebox ──
        _ if op == Op::I32_ADD
            || op == Op::I32_SUB
            || op == Op::I32_MUL
            || op == Op::I32_DIV_S
            || op == Op::I32_DIV_U
            || op == Op::I32_REM_S
            || op == Op::I32_REM_U
            || op == Op::I32_AND
            || op == Op::I32_OR
            || op == Op::I32_XOR
            || op == Op::I32_SHL
            || op == Op::I32_SHR_S
            || op == Op::I32_SHR_U
            || op == Op::I32_ROTL
            || op == Op::I32_ROTR =>
        {
            emit_binary_i32_op(body, op.sub_u8(), rt_idx, temp_idx);
        }

        // ── i32 comparisons (eq, ne): unbox both → compare → rebox ──
        _ if op == Op::EQ || op == Op::NE => {
            emit_binary_i32_cmp(body, op.sub_u8(), rt_idx, temp_idx);
        }

        // ── f64 unary ops: unbox → op → rebox ──
        _ if op == Op::F64_NEG
            || op == Op::F64_ABS
            || op == Op::F64_CEIL
            || op == Op::F64_FLOOR
            || op == Op::F64_TRUNC
            || op == Op::F64_NEAREST
            || op == Op::F64_SQRT =>
        {
            emit_unbox_f64(body, rt_idx);
            body.push(op.sub_u8());
            emit_box_f64(body, rt_idx);
        }

        // ── i32 unary ops ──
        _ if op == Op::I32_EQZ => {
            emit_unbox_i32(body, rt_idx);
            body.push(op.sub_u8());
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::I32_CLZ || op == Op::I32_CTZ || op == Op::I32_POPCNT => {
            emit_unbox_i32(body, rt_idx);
            body.push(op.sub_u8());
            emit_box_i32(body, rt_idx);
        }

        // ── Conversions: unbox source type → convert → rebox target type ──
        _ if op == Op::I32_FROM_F64 => {
            // externref → f64 → i32.trunc_f64_s → externref
            emit_unbox_f64(body, rt_idx);
            body.push(op.sub_u8());
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::F64_FROM_I32 => {
            // externref → i32 → f64.convert_i32_s → externref
            emit_unbox_i32(body, rt_idx);
            body.push(op.sub_u8());
            emit_box_f64(body, rt_idx);
        }

        // ref.func (Closure format): emit ref.func with WASM function index
        _ if op == Op::REF_FUNC => {
            let chunk_idx = read_u16(&chunk.code, ip);
            let uv_count = chunk.code[*ip] as usize;
            *ip += 1;
            *ip += uv_count * 2; // skip upvalue descriptors
            // Store as table index (i32) for call_indirect — box as externref.
            // chunk_idx is the table index because the element section maps chunks 0..N to table slots.
            body.push(0x41); // i32.const
            write_leb128_i32(body, chunk_idx as i32);
            emit_box_i32(body, rt_idx); // i32 → externref
        }
        _ if op == Op::RETHROW || op == Op::DELEGATE => {
            body.push(op.sub_u8());
            let depth = read_leb_u32(&chunk.code, ip);
            write_leb128_u32(body, depth);
        }
        // ── Stack-switching proposal: real core opcodes 0xE0..=0xE6 ──
        _ if op == Op::CONT_NEW => {
            body.push(super::stack_switching::OP_CONT_NEW);
            write_leb128_u32(body, type_ctx.continuation_type_idx);
        }
        _ if op == Op::SUSPEND => {
            let tag_idx = read_u16(&chunk.code, ip);
            body.push(super::stack_switching::OP_SUSPEND);
            write_leb128_u32(body, tag_idx as u32);
        }
        _ if op == Op::RESUME => {
            let _ = read_u16(&chunk.code, ip);
            body.push(super::stack_switching::OP_RESUME);
            write_leb128_u32(body, type_ctx.continuation_type_idx);
            emit_stack_switch_handlers(body, chunk, op_start);
        }
        _ if op == Op::SWITCH => {
            let tag_idx = read_u16(&chunk.code, ip);
            body.push(super::stack_switching::OP_SWITCH);
            write_leb128_u32(body, type_ctx.continuation_type_idx);
            write_leb128_u32(body, tag_idx as u32);
        }
        _ if op == Op::CONT_BIND => {
            let _argc = chunk.code[*ip];
            *ip += 1;
            body.push(super::stack_switching::OP_CONT_BIND);
            write_leb128_u32(body, type_ctx.continuation_type_idx);
            write_leb128_u32(body, type_ctx.continuation_type_idx);
        }
        _ if op == Op::RESUME_THROW => {
            let tag_idx = read_u16(&chunk.code, ip);
            body.push(super::stack_switching::OP_RESUME_THROW);
            write_leb128_u32(body, type_ctx.continuation_type_idx);
            write_leb128_u32(body, tag_idx as u32);
            emit_stack_switch_handlers(body, chunk, op_start);
        }
        _ if op == Op::RESUME_THROW_REF => {
            body.push(super::stack_switching::OP_RESUME_THROW_REF);
            write_leb128_u32(body, type_ctx.continuation_type_idx);
            emit_stack_switch_handlers(body, chunk, op_start);
        }

        _ => {
            // Other core ops: emit WASM byte directly
            body.push(op.sub_u8());
            *ip += op.operand_format().size_in(&chunk.code, *ip);
        }
    }
}

fn emit_thread_prefixed_op(body: &mut Vec<u8>, op: Op, chunk: &Chunk, ip: &mut usize) {
    body.push(0xFE);
    write_leb128_u32(body, op.sub() as u32);
    if op == Op::ATOMIC_FENCE {
        let immediate = chunk.code.get(*ip).copied().unwrap_or(0);
        *ip = (*ip).saturating_add(1).min(chunk.code.len());
        body.push(immediate);
        return;
    }
    if op == Op::THREAD_SPAWN || op == Op::THREAD_JOIN {
        return;
    }
    if next_bytes_decode_opcode(chunk, *ip) {
        write_leb128_u32(body, 0);
        write_leb128_u32(body, 0);
        return;
    }
    let align = read_leb_u32(&chunk.code, ip);
    let is_memory64 = align & 0x80 != 0;
    let spec_align = align & !0x80;
    let offset = if is_memory64 {
        read_leb_u64(&chunk.code, ip)
    } else {
        read_leb_u32(&chunk.code, ip) as u64
    };
    let memidx = if spec_align & 0x40 != 0 {
        Some(read_leb_u32(&chunk.code, ip))
    } else {
        None
    };
    write_leb128_u32(body, spec_align);
    if is_memory64 {
        write_leb128_u64(body, offset);
    } else {
        write_leb128_u32(body, offset as u32);
    }
    if let Some(memidx) = memidx {
        write_leb128_u32(body, memidx);
    }
}

fn emit_simd_prefixed_op(body: &mut Vec<u8>, op: Op, chunk: &Chunk, ip: &mut usize) {
    body.push(0xFD);
    write_leb128_u32(body, op.sub() as u32);
    if op == Op::V128_CONST {
        for _ in 0..16 {
            let byte = chunk.code.get(*ip).copied().unwrap_or(0);
            *ip = (*ip).saturating_add(1).min(chunk.code.len());
            body.push(byte);
        }
        return;
    }
    if op == Op::I8X16_SHUFFLE {
        for _ in 0..16 {
            let byte = chunk.code.get(*ip).copied().unwrap_or(0);
            *ip = (*ip).saturating_add(1).min(chunk.code.len());
            body.push(byte);
        }
        return;
    }
    if matches!(op.sub(), 0x00..=0x0B | 0x54..=0x5D) {
        emit_simd_memarg(body, chunk, ip, default_simd_align(op));
        if matches!(op.sub(), 0x54..=0x5B) {
            let lane = chunk.code.get(*ip).copied().unwrap_or(0);
            *ip = (*ip).saturating_add(1).min(chunk.code.len());
            body.push(lane);
        }
        return;
    }
    if op.operand_format() == OperandFormat::U8 {
        let immediate = chunk.code.get(*ip).copied().unwrap_or(0);
        *ip = (*ip).saturating_add(1).min(chunk.code.len());
        body.push(immediate);
    }
}

fn emit_simd_memarg(body: &mut Vec<u8>, chunk: &Chunk, ip: &mut usize, default_align: u32) {
    if next_bytes_decode_opcode(chunk, *ip) {
        write_leb128_u32(body, default_align);
        write_leb128_u32(body, 0);
        return;
    }
    let mut probe = *ip;
    let marker_align = read_leb_u32(&chunk.code, &mut probe);
    if marker_align & 0x80 == 0 {
        write_leb128_u32(body, default_align);
        write_leb128_u32(body, 0);
        return;
    }
    *ip = probe;
    let memory64 = marker_align & 0x100 != 0;
    let align = marker_align & !0x180;
    let offset = if memory64 {
        read_leb_u64(&chunk.code, ip)
    } else {
        read_leb_u32(&chunk.code, ip) as u64
    };
    let memidx = if align & 0x40 != 0 {
        Some(read_leb_u32(&chunk.code, ip))
    } else {
        None
    };
    write_leb128_u32(body, align);
    if memory64 {
        write_leb128_u64(body, offset);
    } else {
        write_leb128_u32(body, offset as u32);
    }
    if let Some(memidx) = memidx {
        write_leb128_u32(body, memidx);
    }
}

fn default_simd_align(op: Op) -> u32 {
    match op.sub() {
        0x00 | 0x0B => 4,
        0x01 | 0x02 | 0x03 | 0x04 | 0x05 | 0x06 | 0x0A | 0x57 | 0x5B | 0x5D => 3,
        0x09 | 0x56 | 0x5A | 0x5C => 2,
        0x08 | 0x55 | 0x59 => 1,
        _ => 0,
    }
}

// ── Binary operation helpers ─────────────────────────────────────────

/// Emit binary f64 op: [externref_a, externref_b] → f64.op → [externref_result]
/// Uses temp local to save b while unboxing a.
fn emit_binary_f64_op(
    body: &mut Vec<u8>,
    wasm_opcode: u8,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    temp_idx: u32,
) {
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
fn emit_binary_f64_cmp(
    body: &mut Vec<u8>,
    wasm_opcode: u8,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    temp_idx: u32,
) {
    body.push(0x21);
    write_leb128_u32(body, temp_idx); // local.set $temp (save b)
    emit_unbox_f64(body, rt_idx); // toF64(a)
    body.push(0x20);
    write_leb128_u32(body, temp_idx); // local.get $temp (restore b)
    emit_unbox_f64(body, rt_idx); // toF64(b)
    body.push(wasm_opcode); // f64.lt/gt/le/ge → i32
    emit_box_i32(body, rt_idx); // fromI32 → externref
}

/// Emit binary i32 op: [externref_a, externref_b] → i32.op → [externref_result]
fn emit_binary_i32_op(
    body: &mut Vec<u8>,
    wasm_opcode: u8,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    temp_idx: u32,
) {
    body.push(0x21);
    write_leb128_u32(body, temp_idx); // local.set $temp (save b)
    emit_unbox_i32(body, rt_idx); // toI32(a)
    body.push(0x20);
    write_leb128_u32(body, temp_idx); // local.get $temp (restore b)
    emit_unbox_i32(body, rt_idx); // toI32(b)
    body.push(wasm_opcode); // i32.op → i32
    emit_box_i32(body, rt_idx); // fromI32 → externref
}

/// Emit binary i32 comparison: [externref_a, externref_b] → i32.cmp → [externref_result]
fn emit_binary_i32_cmp(
    body: &mut Vec<u8>,
    wasm_opcode: u8,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    temp_idx: u32,
) {
    body.push(0x21);
    write_leb128_u32(body, temp_idx); // local.set $temp (save b)
    emit_unbox_i32(body, rt_idx); // toI32(a)
    body.push(0x20);
    write_leb128_u32(body, temp_idx); // local.get $temp (restore b)
    emit_unbox_i32(body, rt_idx); // toI32(b)
    body.push(wasm_opcode); // i32.eq/ne → i32
    emit_box_i32(body, rt_idx); // fromI32 → externref
}

/// Emit a GC op (prefix 0xFB) — emit real WASM GC binary encoding with type indices.
///
/// GC refs (ref $struct, ref $array) are NOT subtypes of externref.
/// We use externref as our universal local type, so:
/// - After GC ops that PRODUCE refs: emit `extern.convert_any` (0xFB 0x1B) → externref
/// - Before GC ops that CONSUME refs: emit `any.convert_extern` (0xFB 0x1A) → anyref,
///   then `ref.cast` to the specific GC type
fn emit_gc_op(
    body: &mut Vec<u8>,
    op: Op,
    chunk: &Chunk,
    ip: &mut usize,
    _rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    type_ctx: &WasmTypeContext,
    temp_idx: u32,
) {
    match op {
        _ if op == Op::STRUCT_NEW => {
            let prop_count = read_u16(&chunk.code, ip);
            let typeidx = wasm_struct_type_matching_field_count(chunk, type_ctx, prop_count);
            body.push(0xFB);
            write_leb128_u32(body, 0x00); // struct.new
            write_leb128_u32(body, typeidx);
            emit_externalize(body); // (ref $struct) → externref
        }
        _ if op == Op::STRUCT_GET => {
            let field_name_idx = read_u16(&chunk.code, ip);
            let (typeidx, fieldidx) = wasm_struct_field_for_name(chunk, type_ctx, field_name_idx);
            emit_internalize(body); // externref → anyref
            emit_ref_cast(body, typeidx); // anyref → (ref $struct)
            body.push(0xFB);
            write_leb128_u32(body, 0x02); // struct.get
            write_leb128_u32(body, typeidx);
            write_leb128_u32(body, fieldidx);
            // Result is externref (field type) — no conversion needed
        }
        _ if op == Op::STRUCT_SET => {
            let field_name_idx = read_u16(&chunk.code, ip);
            let (typeidx, fieldidx) = wasm_struct_field_for_name(chunk, type_ctx, field_name_idx);
            // Stack: [externref_obj, externref_val]. struct.set expects
            // [(ref $struct), externref_val] and returns void, while the
            // VM leaves the assigned value on stack. Save the value, cast
            // the object, set the field, then reload the saved value.
            body.push(0x21); // local.set $temp = val
            write_leb128_u32(body, temp_idx);
            emit_internalize(body); // obj: externref → anyref
            emit_ref_cast(body, typeidx);
            body.push(0x20); // local.get $temp = val
            write_leb128_u32(body, temp_idx);
            body.push(0xFB);
            write_leb128_u32(body, 0x05); // struct.set
            write_leb128_u32(body, typeidx);
            write_leb128_u32(body, fieldidx);
            body.push(0x20); // VM-compatible result: assigned value
            write_leb128_u32(body, temp_idx);
        }
        _ if op == Op::ARRAY_NEW_FIXED => {
            // Spec: `array.new_fixed $t N` (0xFB 0x08), pops N values.
            let elem_count = read_u16(&chunk.code, ip);
            body.push(0xFB);
            write_leb128_u32(body, 0x08);
            write_leb128_u32(body, type_ctx.array_type_idx);
            write_leb128_u32(body, elem_count as u32);
            emit_externalize(body); // (ref $arr) → externref
        }
        _ if op == Op::ARRAY_NEW => {
            // Spec: `array.new $t` (0xFB 0x06), pops [value, length i32].
            // Our bytecode emits the 2-byte type-index immediate like the
            // fixed variant; consume it, drop the type index, and pass
            // through to the engine (which will fill len copies of value).
            let _typeidx = read_u16(&chunk.code, ip);
            body.push(0xFB);
            write_leb128_u32(body, 0x06);
            write_leb128_u32(body, type_ctx.array_type_idx);
            emit_externalize(body);
        }
        _ if op == Op::ARRAY_NEW_DEFAULT => {
            // Spec: `array.new_default $t` (0xFB 0x07), pops [length].
            let _typeidx = read_u16(&chunk.code, ip);
            body.push(0xFB);
            write_leb128_u32(body, 0x07);
            write_leb128_u32(body, type_ctx.array_type_idx);
            emit_externalize(body);
        }
        _ if op == Op::ARRAY_NEW_DATA => {
            // Spec: `array.new_data $t $d`, pops [offset, size].
            let _typeidx = read_u16(&chunk.code, ip);
            let data_idx = read_u16(&chunk.code, ip);
            body.push(0xFB);
            write_leb128_u32(body, 0x09);
            write_leb128_u32(body, type_ctx.array_type_idx);
            write_leb128_u32(body, data_idx as u32);
            emit_externalize(body);
        }
        _ if op == Op::ARRAY_NEW_ELEM => {
            let _typeidx = read_u16(&chunk.code, ip);
            let elem_idx = read_u16(&chunk.code, ip);
            body.push(0xFB);
            write_leb128_u32(body, 0x0A);
            write_leb128_u32(body, type_ctx.array_type_idx);
            write_leb128_u32(body, elem_idx as u32);
            emit_externalize(body);
        }
        _ if op == Op::ARRAY_GET_S || op == Op::ARRAY_GET_U => {
            // Packed variants. Semantics identical to array.get for our
            // externref-only arrays but we must still emit the spec byte.
            let _typeidx = read_u16(&chunk.code, ip);
            body.push(0x21);
            write_leb128_u32(body, temp_idx); // save idx
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0x20);
            write_leb128_u32(body, temp_idx);
            emit_unbox_i32(body, _rt_idx);
            body.push(0xFB);
            write_leb128_u32(body, op.sub() as u32);
            write_leb128_u32(body, type_ctx.array_type_idx);
        }
        _ if op == Op::ARRAY_INIT_DATA => {
            let _typeidx = read_u16(&chunk.code, ip);
            let data_idx = read_u16(&chunk.code, ip);
            body.push(0xFB);
            write_leb128_u32(body, 0x12);
            write_leb128_u32(body, type_ctx.array_type_idx);
            write_leb128_u32(body, data_idx as u32);
        }
        _ if op == Op::ARRAY_INIT_ELEM => {
            let _typeidx = read_u16(&chunk.code, ip);
            let elem_idx = read_u16(&chunk.code, ip);
            body.push(0xFB);
            write_leb128_u32(body, 0x13);
            write_leb128_u32(body, type_ctx.array_type_idx);
            write_leb128_u32(body, elem_idx as u32);
        }
        _ if op == Op::STRUCT_NEW_DEFAULT => {
            let typeidx = read_u16(&chunk.code, ip);
            body.push(0xFB);
            write_leb128_u32(body, 0x01);
            write_leb128_u32(
                body,
                wasm_struct_type_for_chunk_type(chunk, type_ctx, typeidx),
            );
            emit_externalize(body);
        }
        // ── Custom Descriptors proposal emission ─────────────────────────
        // Opcodes from `proposals/custom-descriptors/`.
        // The binary format of the described struct types themselves is
        // already produced by `types.rs` (CD_DESCRIBES / CD_DESCRIPTOR
        // prefixes on each struct); these three opcodes are the operator
        // side of the proposal.
        //
        // Encodings:
        //   struct.new_desc $typeidx          → 0xFB 0x20 typeidx
        //   struct.new_default_desc $typeidx  → 0xFB 0x21 typeidx
        //   ref.get_desc $typeidx             → 0xFB 0x22 typeidx
        //
        // Our VM stores the operand as a u16 chunk-local typeidx; we
        // remap it to the WASM typeidx via `type_ctx` if a real one is
        // known, else emit 0 (conservative — engines that consume our
        // modules with descriptors enabled re-derive the shape from the
        // type section).
        _ if op == Op::STRUCT_NEW_DESC => {
            let typeidx = read_u16(&chunk.code, ip);
            body.push(0xFB);
            write_leb128_u32(body, 0x20);
            write_leb128_u32(
                body,
                wasm_struct_type_for_chunk_type(chunk, type_ctx, typeidx),
            );
        }
        _ if op == Op::STRUCT_NEW_DEFAULT_DESC => {
            let typeidx = read_u16(&chunk.code, ip);
            body.push(0xFB);
            write_leb128_u32(body, 0x21);
            write_leb128_u32(
                body,
                wasm_struct_type_for_chunk_type(chunk, type_ctx, typeidx),
            );
        }
        _ if op == Op::REF_GET_DESC => {
            let typeidx = read_u16(&chunk.code, ip);
            body.push(0xFB);
            write_leb128_u32(body, 0x22);
            write_leb128_u32(
                body,
                wasm_struct_type_for_chunk_type(chunk, type_ctx, typeidx),
            );
        }
        _ if op == Op::STRUCT_GET_S || op == Op::STRUCT_GET_U => {
            // Our struct.get uses a field-name-constant u16 operand;
            // spec packed variants take typeidx + fieldidx. Emit the
            // spec byte with conservative indices for round-trip sanity.
            let field_name_idx = read_u16(&chunk.code, ip);
            let (typeidx, fieldidx) = wasm_struct_field_for_name(chunk, type_ctx, field_name_idx);
            emit_internalize(body);
            body.push(0xFB);
            write_leb128_u32(body, op.sub() as u32);
            write_leb128_u32(body, typeidx);
            write_leb128_u32(body, fieldidx);
        }
        _ if op == Op::REF_TEST_NULL => {
            // ref.test (ref null ht): `0xFB 0x15 <heaptype>`. Our bytecode
            // operand is a constant-pool index for a type *name*; we emit
            // a conservative `anyref` heaptype so the module validates on
            // any GC-capable engine. Precise per-type dispatch still runs
            // in the VM via `test_type`.
            *ip += op.operand_format().size_in(&chunk.code, *ip);
            body.push(0xFB);
            write_leb128_u32(body, 0x15);
            body.push(HT_ANY);
            emit_box_i32(body, _rt_idx);
        }
        _ if op == Op::REF_CAST_NULL => {
            *ip += op.operand_format().size_in(&chunk.code, *ip);
            body.push(0xFB);
            write_leb128_u32(body, 0x17);
            body.push(HT_ANY);
        }
        _ if op == Op::ANY_CONVERT_EXTERN => {
            // Our externref is the universal value carrier — the op is a
            // no-op in our VM but spec-emit for round-trip fidelity.
            body.push(0xFB);
            write_leb128_u32(body, 0x1A);
        }
        _ if op == Op::EXTERN_CONVERT_ANY => {
            body.push(0xFB);
            write_leb128_u32(body, 0x1B);
        }
        _ if op == Op::ARRAY_GET => {
            // Stack: [externref_arr, externref_idx]
            // Need: [(ref null $arr), i32] for array.get
            // Save idx to temp, convert arr, restore idx as i32
            body.push(0x21);
            write_leb128_u32(body, temp_idx); // local.set $temp (save idx)
            emit_internalize(body); // externref_arr → anyref
            emit_ref_cast_array(body, type_ctx.array_type_idx); // anyref → (ref null $arr)
            body.push(0x20);
            write_leb128_u32(body, temp_idx); // local.get $temp (restore idx)
            emit_unbox_i32(body, _rt_idx); // externref_idx → i32
            body.push(0xFB);
            write_leb128_u32(body, 0x0B); // array.get
            write_leb128_u32(body, type_ctx.array_type_idx);
            // Result: externref (element type)
        }
        _ if op == Op::ARRAY_SET => {
            // Stack: [externref_arr, externref_idx, externref_val]
            // Need: [(ref null $arr), i32, externref] for array.set
            // Save val and idx, convert arr, restore idx as i32, restore val
            body.push(0x21);
            write_leb128_u32(body, temp_idx); // local.set $temp (save val)
            // Stack: [externref_arr, externref_idx]
            body.push(0x21);
            write_leb128_u32(body, temp_idx + 1); // local.set $temp2 (save idx) — use next local
            // Stack: [externref_arr]
            emit_internalize(body); // externref_arr → anyref
            emit_ref_cast_array(body, type_ctx.array_type_idx); // anyref → (ref null $arr)
            body.push(0x20);
            write_leb128_u32(body, temp_idx + 1); // local.get $temp2 (restore idx)
            emit_unbox_i32(body, _rt_idx); // externref_idx → i32
            body.push(0x20);
            write_leb128_u32(body, temp_idx); // local.get $temp (restore val)
            // Stack: [(ref null $arr), i32, externref]
            body.push(0xFB);
            write_leb128_u32(body, 0x0E); // array.set
            write_leb128_u32(body, type_ctx.array_type_idx);
            // array.set is void in WASM but our VM leaves a value on stack.
            // Push dummy so subsequent drop has something to consume.
            body.push(0xD0);
            body.push(0x6F); // ref.null extern
        }
        _ if op == Op::ARRAY_LENGTH => {
            // Stack: [externref_arr] → need (ref null array)
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0xFB);
            write_leb128_u32(body, 0x0F); // array.len → i32
            // Result is i32, box to externref
            emit_box_i32(body, _rt_idx);
        }
        _ if op == Op::ARRAY_FILL => {
            emit_internalize(body);
            body.push(0xFB);
            write_leb128_u32(body, 0x10);
            write_leb128_u32(body, type_ctx.array_type_idx);
        }
        _ if op == Op::ARRAY_COPY => {
            body.push(0xFB);
            write_leb128_u32(body, 0x11);
            write_leb128_u32(body, type_ctx.array_type_idx);
            write_leb128_u32(body, type_ctx.array_type_idx);
        }
        _ if op == Op::ARRAY_NEW_DEFAULT => {
            body.push(0xFB);
            write_leb128_u32(body, 0x07);
            write_leb128_u32(body, type_ctx.array_type_idx);
            emit_externalize(body);
        }
        _ if op == Op::REF_TEST || op == Op::REF_CAST => {
            // ref.test / ref.cast (non-null): `0xFB {0x14|0x16} <heaptype>`.
            // Our bytecode carries a constant-pool string index (the type
            // name) rather than a WASM typeidx. We resolve the name via
            // `type_ctx.struct_type(name)` when available — otherwise fall
            // back to `anyref` so the module still validates. VM-side the
            // check runs through `test_type` which handles the precise
            // name-based lookup regardless of what we encode here.
            let name_idx = read_u16(&chunk.code, ip) as usize;
            let ht_bytes = resolve_heaptype_from_name(chunk, name_idx, type_ctx);
            body.push(0xFB);
            write_leb128_u32(body, op.sub() as u32);
            body.extend_from_slice(&ht_bytes);
        }
        // `ref.eq` moved to the core prefix (0xD3) — emit path is in
        // `emit_core_op`. This branch is kept intentionally empty as a
        // breadcrumb for anyone searching for it.
        _ if op == Op::BR_ON_CAST || op == Op::BR_ON_CAST_FAIL => {
            // br_on_cast  flags l ht1 ht2 : `0xFB 0x18 <flags> <label> <ht1> <ht2>`
            // br_on_cast_fail              : `0xFB 0x19 <flags> <label> <ht1> <ht2>`
            //
            // `flags` is a 2-bit bitfield: bit 0 = null source, bit 1 =
            // null target. Our bytecode operand is (u16 type-name-idx,
            // u8 label-depth). The u8 depth maps 1:1 to WASM's labelidx
            // (structured block depth), so we can emit it directly.
            //   * source heaptype → `anyref` (we accept any non-null ref)
            //   * target heaptype → resolved from the type-name, fall back
            //     to `anyref` when the name isn't in the type table
            let name_idx = read_u16(&chunk.code, ip) as usize;
            let depth = chunk.code[*ip];
            *ip += 1;
            let ht_bytes = resolve_heaptype_from_name(chunk, name_idx, type_ctx);
            body.push(0xFB);
            write_leb128_u32(body, op.sub() as u32);
            body.push(0x00); // flags: non-null source, non-null target
            write_leb128_u32(body, depth as u32);
            body.push(HT_ANY); // ht1: source
            body.extend_from_slice(&ht_bytes); // ht2: target
        }
        _ if op == Op::I31_NEW => {
            // i31.new expects i32 — unbox externref first
            emit_unbox_i32(body, _rt_idx);
            body.push(0xFB);
            write_leb128_u32(body, 0x1C); // ref.i31
            emit_externalize(body); // (ref i31) → externref
        }
        _ if op == Op::I31_GET_S || op == Op::I31_GET_U => {
            emit_internalize(body); // externref → anyref
            body.push(0xFB);
            write_leb128_u32(body, op.sub() as u32);
            emit_box_i32(body, _rt_idx); // i32 → externref
        }
        _ => {
            // Other GC ops: emit directly
            body.push(0xFB);
            write_leb128_u32(body, op.sub() as u32);
            *ip += op.operand_format().size_in(&chunk.code, *ip);
        }
    }
}

/// Emit inline dyn_add: type check both operands, f64 arithmetic if numbers, string concat if not.
/// Uses wasm:js-number and wasm:js-string builtins (standard WASM proposals).
#[allow(dead_code)]
fn emit_dyn_binary_numeric(
    body: &mut Vec<u8>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    temp_idx: u32,
    f64_opcode: u8,
) {
    // Stack: [externref_a, externref_b]
    // Simple approach: always treat as f64 (matches how stdlib uses dyn_add for numbers)
    // Save b, unbox a, restore b, unbox b, operate, rebox
    body.push(0x21);
    write_leb128_u32(body, temp_idx); // local.set $temp (save b)
    emit_unbox_f64(body, rt_idx); // toF64(a)
    body.push(0x20);
    write_leb128_u32(body, temp_idx); // local.get $temp (restore b)
    emit_unbox_f64(body, rt_idx); // toF64(b)
    body.push(f64_opcode); // f64.add/sub/mul/div
    emit_box_f64(body, rt_idx); // fromF64 → externref
}

/// Emit inline dyn comparison: unbox both as f64, compare, box i32 result.
#[allow(dead_code)]
fn emit_dyn_binary_cmp(
    body: &mut Vec<u8>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    temp_idx: u32,
    f64_cmp_opcode: u8,
) {
    body.push(0x21);
    write_leb128_u32(body, temp_idx); // local.set $temp (save b)
    emit_unbox_f64(body, rt_idx); // toF64(a)
    body.push(0x20);
    write_leb128_u32(body, temp_idx); // local.get $temp (restore b)
    emit_unbox_f64(body, rt_idx); // toF64(b)
    body.push(f64_cmp_opcode); // f64.lt/gt/le/ge/eq/ne → i32
    emit_box_i32(body, rt_idx); // fromI32 → externref
}

/// Emit a string constant from the chunk's constant pool.
/// Builds the string char by char using wasm:js-string fromCharCode + concat.
fn emit_string_const(
    body: &mut Vec<u8>,
    chunk: &Chunk,
    const_idx: usize,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
) {
    if let Some(Value::String(s)) = chunk.constants.get(const_idx) {
        let chars: Vec<char> = s.chars().collect();
        if chars.is_empty() {
            body.push(0xD0);
            body.push(0x6F); // ref.null extern
            return;
        }
        // First char
        body.push(0x41);
        write_leb128_i32(body, chars[0] as i32);
        emit_import_call(body, rt_idx, "wasm:js-string", "fromCharCode");
        // Concat remaining chars
        for &ch in &chars[1..] {
            body.push(0x41);
            write_leb128_i32(body, ch as i32);
            emit_import_call(body, rt_idx, "wasm:js-string", "fromCharCode");
            emit_import_call(body, rt_idx, "wasm:js-string", "concat");
        }
    } else {
        body.push(0xD0);
        body.push(0x6F);
    }
}

/// Resolve the heaptype bytes for a `ref.test / ref.cast / br_on_cast`
/// operand. The compiler stores a constant-pool index pointing at a
/// string type name. We look up the name in `type_ctx.struct_type_indices`
/// and emit `(signed LEB128) typeidx` when found; otherwise fall back to
/// the abstract `anyref` single-byte heaptype so the binary still validates.
fn resolve_heaptype_from_name(
    chunk: &Chunk,
    name_idx: usize,
    type_ctx: &WasmTypeContext,
) -> Vec<u8> {
    if let Some(Value::String(s)) = chunk.constants.get(name_idx) {
        if let Some(idx) = type_ctx.struct_type(s) {
            let mut buf = Vec::new();
            write_leb128_i32(&mut buf, idx as i32);
            return buf;
        }
    }
    vec![HT_ANY]
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

fn emit_memory64_op(body: &mut Vec<u8>, chunk: &Chunk, ip: &mut usize, wasm_op: u8, align: u32) {
    body.push(wasm_op);
    let (align, offset, memidx) = read_optional_memarg64(chunk, ip, align);
    encode_memarg_with_memidx(body, align, offset, memidx);
}

fn emit_table64_core_op(body: &mut Vec<u8>, chunk: &Chunk, ip: &mut usize, wasm_op: u8) {
    body.push(wasm_op);
    let table_idx = chunk.code.get(*ip).copied().unwrap_or(0);
    *ip += 1;
    write_leb128_u32(body, table_idx as u32);
}

fn emit_table64_fc_op(
    body: &mut Vec<u8>,
    chunk: &Chunk,
    ip: &mut usize,
    subopcode: u8,
    has_two_indices: bool,
) {
    body.push(0xFC);
    write_leb128_u32(body, subopcode as u32);
    let table_idx = chunk.code.get(*ip).copied().unwrap_or(0);
    *ip += 1;
    write_leb128_u32(body, table_idx as u32);
    if has_two_indices {
        let second_idx = chunk.code.get(*ip).copied().unwrap_or(table_idx);
        *ip += 1;
        write_leb128_u32(body, second_idx as u32);
    }
}

/// Emit a VM-internal op (prefix 0xFF) — lowered to WASM equivalents or runtime calls.
fn emit_vm_internal_op(
    body: &mut Vec<u8>,
    op: Op,
    chunk: &Chunk,
    ip: &mut usize,
    op_start: usize,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    temp_idx: u32,
    type_ctx: &WasmTypeContext,
) {
    match op {
        _ if op == Op::CONST => {
            let idx = read_u16(&chunk.code, ip);
            if let Some(val) = chunk.constants.get(idx as usize) {
                match val {
                    Value::F64(n) => {
                        body.push(0x44);
                        body.extend_from_slice(&n.to_le_bytes());
                        emit_box_f64(body, rt_idx);
                    }
                    Value::I32(n) => {
                        body.push(0x41);
                        write_leb128_i32(body, *n);
                        emit_box_i32(body, rt_idx);
                    }
                    Value::I64(n) => {
                        body.push(0x42);
                        write_leb128_i64(body, *n);
                        emit_box_i32(body, rt_idx);
                    }
                    _ => {
                        body.push(0xD0);
                        body.push(0x6F);
                    } // ref.null externref
                }
            }
        }
        _ if op == Op::NULL => {
            body.push(0xD0);
            body.push(HT_EXTERN);
        }
        _ if op == Op::NULL_FUNC => {
            body.push(0xD0);
            body.push(HT_FUNC);
        }
        _ if op == Op::NULL_ANY => {
            body.push(0xD0);
            body.push(HT_ANY);
        }
        _ if op == Op::NULL_NONE => {
            body.push(0xD0);
            body.push(HT_NONE);
        }
        _ if op == Op::UNDEFINED => {
            // `global.get $js_undefined` — returns the JS host's `undefined`
            // singleton (distinct from `ref.null extern`). The import is
            // declared in sections.rs; its index is fixed by JS_GLOBAL_UNDEFINED.
            body.push(0x23);
            write_leb128_u32(body, super::sections::JS_GLOBAL_UNDEFINED);
        }
        _ if op == Op::SYMBOL => {
            // Skip the const index; emit a stub externref. Symbol identity
            // is a VM-only concept without a `wasm:js-symbol` constructor
            // import, so we lose identity across the boundary.
            let _ = read_u16(&chunk.code, ip);
            body.push(0xD0);
            body.push(0x6F);
        }
        _ if op == Op::BIGINT => {
            // Skip the const index; emit a boxed i64 so host JS sees a
            // number (closest standard value without a bigint constructor).
            let idx = read_u16(&chunk.code, ip);
            let n = chunk
                .constants
                .get(idx as usize)
                .map(|v| match v {
                    Value::I64(n) => *n,
                    Value::I32(n) => *n as i64,
                    _ => 0,
                })
                .unwrap_or(0);
            body.push(0x42);
            write_leb128_i64(body, n);
            // box i64 as i32 (truncate high bits) — acceptable for typical bigint uses
            body.push(0xA7); // i32.wrap_i64
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::REF_IS_UNDEFINED => {
            emit_import_call(body, rt_idx, "wasm:js-undefined", "test");
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::REF_IS_SYMBOL => {
            emit_import_call(body, rt_idx, "wasm:js-symbol", "test");
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::REF_IS_BIGINT => {
            emit_import_call(body, rt_idx, "wasm:js-bigint", "test");
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::REF_IS_I32 => {
            super::sections::emit_test_i32(body, rt_idx);
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::REF_IS_U32 => {
            super::sections::emit_test_u32(body, rt_idx);
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::NUM_BOX_U32 => {
            // Top of stack is a boxed i32 — unbox, rebox as u32 for the host.
            emit_unbox_i32(body, rt_idx);
            super::sections::emit_box_u32(body, rt_idx);
        }
        _ if op == Op::NUM_UNBOX_U32 => {
            super::sections::emit_unbox_u32(body, rt_idx);
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::BOOL_CAST => {
            super::sections::emit_unbox_bool(body, rt_idx);
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::STR_CAST => {
            super::sections::emit_str_cast(body, rt_idx);
        }
        _ if op == Op::STR_FROM_I32 => {
            emit_unbox_i32(body, rt_idx);
            super::sections::emit_str_from_i32(body, rt_idx);
        }
        _ if op == Op::STR_FROM_U32 => {
            emit_unbox_i32(body, rt_idx);
            super::sections::emit_str_from_u32(body, rt_idx);
        }
        _ if op == Op::STR_FROM_I64 => {
            // i64 is stored boxed as i32 in our ABI; widen via extend_s.
            emit_unbox_i32(body, rt_idx);
            body.push(0xAC); // i64.extend_i32_s
            super::sections::emit_str_from_i64(body, rt_idx);
        }
        _ if op == Op::STR_FROM_U64 => {
            emit_unbox_i32(body, rt_idx);
            body.push(0xAD); // i64.extend_i32_u
            super::sections::emit_str_from_u64(body, rt_idx);
        }
        _ if op == Op::STR_FROM_F64 => {
            emit_unbox_f64(body, rt_idx);
            super::sections::emit_str_from_f64(body, rt_idx);
        }
        _ if op == Op::SYMBOL_EQ => {
            super::sections::emit_symbol_equals(body, rt_idx);
            emit_box_i32(body, rt_idx);
        }

        // ── Typed stack-switching helpers ────────────────────────────────
        // These VM-internal typed helpers lower to the real proposal bytes.
        // The untyped stack-switching opcodes themselves live in core 0xE0..=0xE6
        // and are emitted by `emit_core_op` above.
        _ if op == Op::CONT_NEW_TYPED => {
            let _ = read_u16(&chunk.code, ip);
            body.push(super::stack_switching::OP_CONT_NEW);
            write_leb128_u32(body, type_ctx.continuation_type_idx);
        }
        _ if op == Op::SUSPEND_TYPED => {
            let tag_idx = read_u16(&chunk.code, ip);
            body.push(super::stack_switching::OP_SUSPEND);
            let wasm_tag_idx = chunk
                .continuation_tags
                .get(tag_idx as usize)
                .and_then(|tag| {
                    type_ctx
                        .continuation_tag_indices
                        .get(&super::types::continuation_tag_key(tag))
                        .copied()
                })
                .unwrap_or(1);
            write_leb128_u32(body, wasm_tag_idx);
        }
        _ if op == Op::RESUME_TYPED => {
            let _ = read_u16(&chunk.code, ip);
            body.push(super::stack_switching::OP_RESUME);
            write_leb128_u32(body, type_ctx.continuation_type_idx);
            emit_stack_switch_handlers(body, chunk, op_start);
        }
        _ if op == Op::TRUE => {
            // `global.get $js_true` — produces an actual JS `true` boolean,
            // not a boxed `1` (which previously confused `typeof` on the
            // host side). See js-primitive-builtins proposal for globals.
            body.push(0x23);
            write_leb128_u32(body, super::sections::JS_GLOBAL_TRUE);
        }
        _ if op == Op::FALSE => {
            body.push(0x23);
            write_leb128_u32(body, super::sections::JS_GLOBAL_FALSE);
        }
        _ if op == Op::I32_CONST_0 => {
            body.push(0x41);
            write_leb128_i32(body, 0);
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::I32_CONST_1 => {
            body.push(0x41);
            write_leb128_i32(body, 1);
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::F64_CONST_0 => {
            body.push(0x44);
            body.extend_from_slice(&0.0f64.to_le_bytes());
            emit_box_f64(body, rt_idx);
        }
        _ if op == Op::CALL_IMPORT => {
            let import_idx = read_u16(&chunk.code, ip);
            let _argc = chunk.code[*ip];
            *ip += 1;
            body.push(0x10);
            write_leb128_u32(body, import_idx as u32);
        }
        _ if op == Op::BR_IF_FALSE => {
            // Legacy flat jump — skip operand, emit nop
            // (compiler should use structured BR_IF instead)
            let _offset = read_i16(&chunk.code, ip);
            body.push(0x01); // nop
        }
        _ if op == Op::BR_IF_NULL => {
            let _offset = read_i16(&chunk.code, ip);
            body.push(0x01); // nop
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
        _ if op == Op::STR_FROM_CODE_POINT => {
            // (externref codepoint) → (i32)
            emit_unbox_i32(body, rt_idx);
            emit_import_call(body, rt_idx, "wasm:js-string", "fromCodePoint");
        }
        _ if op == Op::STR_CODE_POINT_AT => {
            // Stack: [externref_str, externref_idx] → (externref, i32)
            emit_unbox_i32(body, rt_idx);
            emit_import_call(body, rt_idx, "wasm:js-string", "codePointAt");
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::STR_INTO_CHAR_CODES => {
            // Into-array variant: we drop the target-array arg and simply
            // call intoCharCodeArray on (str, str itself, 0) as a placeholder
            // — the common JS toolchains expect the caller to pass a
            // preallocated Int16Array. In our runtime the receiving array
            // is created on the VM side (STR_INTO_CHAR_CODES opcode); the
            // emitted .wasm is equivalent to a noop that returns the string
            // length. Using `length` keeps a valid signature without needing
            // a real array reference here.
            emit_import_call(body, rt_idx, "wasm:js-string", "length");
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::STR_FROM_CHAR_CODES => {
            // Expect: [externref_array]. We forward to fromCharCodeArray
            // with (array, 0, -1) — the host will interpret -1 as "to end".
            body.push(0x41);
            write_leb128_i32(body, 0);
            body.push(0x41);
            write_leb128_i32(body, -1);
            emit_import_call(body, rt_idx, "wasm:js-string", "fromCharCodeArray");
        }
        _ if op == Op::STR_SUBSTRING => {
            // Stack: [externref_str, externref_start, externref_end]
            // Need: [externref_str, i32_start, i32_end]
            // Save end, unbox start (but start is below end on stack)
            // Simpler: save end to temp, save start to temp+1, keep str,
            // restore start as i32, restore end as i32
            body.push(0x21);
            write_leb128_u32(body, temp_idx); // save end
            body.push(0x21);
            write_leb128_u32(body, temp_idx + 1); // save start
            // Stack: [externref_str]
            body.push(0x20);
            write_leb128_u32(body, temp_idx + 1); // restore start
            emit_unbox_i32(body, rt_idx); // → i32
            body.push(0x20);
            write_leb128_u32(body, temp_idx); // restore end
            emit_unbox_i32(body, rt_idx); // → i32
            emit_import_call(body, rt_idx, "wasm:js-string", "substring");
        }
        // String ops not in wasm:js-string → stub: drop args, return first arg or null
        _ if op == Op::STR_INDEX_OF => {
            // Stack: [externref_str, externref_substr]
            // Inline indexOf: for i=0 to len-sublen, check substring(i, i+sublen) == substr
            body.push(0x21);
            write_leb128_u32(body, temp_idx); // $t0 = substr
            body.push(0x21);
            write_leb128_u32(body, temp_idx + 1); // $t1 = str
            // Get substr length
            body.push(0x20);
            write_leb128_u32(body, temp_idx); // substr
            emit_import_call(body, rt_idx, "wasm:js-string", "length"); // → i32
            emit_box_i32(body, rt_idx);
            body.push(0x21);
            write_leb128_u32(body, temp_idx + 2); // $t2 = sublen (boxed)
            // Get str length
            body.push(0x20);
            write_leb128_u32(body, temp_idx + 1); // str
            emit_import_call(body, rt_idx, "wasm:js-string", "length"); // → i32
            emit_box_i32(body, rt_idx);
            body.push(0x21);
            write_leb128_u32(body, temp_idx + 3); // $t3 = strlen (boxed)
            // i = 0 (boxed)
            body.push(0x41);
            write_leb128_i32(body, 0);
            emit_box_i32(body, rt_idx);
            body.push(0x21);
            write_leb128_u32(body, temp_idx + 4); // $t4 = i (boxed)
            // block $exit (result externref) { loop $search (void) { ... } }
            body.push(0x02);
            body.push(TYPE_EXTERNREF); // block (result externref)
            body.push(0x03);
            body.push(0x40); // loop (void)
            // if i > limit, push -1 and break
            body.push(0x20);
            write_leb128_u32(body, temp_idx + 4);
            emit_unbox_i32(body, rt_idx);
            body.push(0x20);
            write_leb128_u32(body, temp_idx + 3);
            emit_unbox_i32(body, rt_idx);
            body.push(0x4A); // i32.gt_s
            body.push(0x04);
            body.push(0x40); // if (void) — past limit
            body.push(0x41);
            write_leb128_i32(body, -1);
            emit_box_i32(body, rt_idx);
            body.push(0x0C);
            write_leb128_u32(body, 2); // br block (if=0, loop=1, block=2)
            body.push(0x0B); // end if
            // substring(str, i, i + sublen)
            body.push(0x20);
            write_leb128_u32(body, temp_idx + 1); // str
            body.push(0x20);
            write_leb128_u32(body, temp_idx + 4);
            emit_unbox_i32(body, rt_idx); // i → i32
            body.push(0x20);
            write_leb128_u32(body, temp_idx + 4);
            emit_unbox_i32(body, rt_idx); // i → i32
            body.push(0x20);
            write_leb128_u32(body, temp_idx + 2);
            emit_unbox_i32(body, rt_idx); // sublen → i32
            body.push(0x6A); // i32.add → i + sublen
            emit_import_call(body, rt_idx, "wasm:js-string", "substring"); // (str, i, i+sublen) → string
            // Compare with substr
            body.push(0x20);
            write_leb128_u32(body, temp_idx); // substr
            emit_import_call(body, rt_idx, "wasm:js-string", "equals"); // → i32
            body.push(0x04);
            body.push(0x40); // if (void)
            // Found: push i (boxed) and break to $exit
            body.push(0x20);
            write_leb128_u32(body, temp_idx + 4); // i (already boxed)
            body.push(0x0C);
            write_leb128_u32(body, 2); // br $exit (depth: if=0, loop=1, block=2)
            body.push(0x0B); // end if
            // i++
            body.push(0x20);
            write_leb128_u32(body, temp_idx + 4);
            emit_unbox_i32(body, rt_idx);
            body.push(0x41);
            write_leb128_i32(body, 1);
            body.push(0x6A); // i32.add
            emit_box_i32(body, rt_idx);
            body.push(0x21);
            write_leb128_u32(body, temp_idx + 4); // save i
            body.push(0x0C);
            write_leb128_u32(body, 0); // br $search (loop)
            body.push(0x0B); // end loop
            // Not found: push -1 (boxed)
            body.push(0x41);
            write_leb128_i32(body, -1);
            emit_box_i32(body, rt_idx);
            body.push(0x0B); // end block
            // Result: externref (i or -1, boxed)
        }
        // Binary string ops (2 args → 1 result): drop second, keep first
        _ if op == Op::STR_LAST_INDEX_OF
            || op == Op::STR_STARTS_WITH
            || op == Op::STR_ENDS_WITH
            || op == Op::STR_CONTAINS
            || op == Op::STR_SPLIT
            || op == Op::STR_REPEAT =>
        {
            body.push(0x1A); // drop second arg, keep first as result
        }
        // Ternary string ops (3 args → 1 result): drop two, keep first
        _ if op == Op::STR_SLICE
            || op == Op::STR_REPLACE
            || op == Op::STR_PAD_START
            || op == Op::STR_PAD_END =>
        {
            body.push(0x1A);
            body.push(0x1A); // drop 2, keep first
        }
        // Unary string ops (1 arg → 1 result): keep as-is
        _ if op == Op::STR_CHAR_AT
            || op == Op::STR_TO_UPPER
            || op == Op::STR_TO_LOWER
            || op == Op::STR_TRIM
            || op == Op::STR_TRIM_START
            || op == Op::STR_TRIM_END
            || op == Op::STR_REVERSE =>
        {
            // pass through — input externref becomes output externref
            body.push(0x01); // nop
        }
        _ if op == Op::STR_CONCAT_N => {
            let n = chunk.code[*ip];
            *ip += 1;
            // Concat N strings: chain wasm:js-string concat calls
            for _ in 1..n {
                emit_import_call(body, rt_idx, "wasm:js-string", "concat");
            }
        }

        // Phase E: the 9  ARRAY_* opcodes (PUSH/POP/SLICE/
        // JOIN/REVERSE/CONTAINS/INDEX_OF/CONCAT/SHIFT) were removed
        // along with their inline WASM GC emit sequences. Callers
        // compile to  CALL_IMPORTs directly via
        // .

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
            body.push(0x22);
            write_leb128_u32(body, temp_idx); // tee value to temp
            // Check null first
            body.push(0xD1); // ref.is_null → i32
            body.push(0x04);
            body.push(TYPE_EXTERNREF); // if null (result externref)
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
            for (ci, c) in chunk.constants.iter().enumerate() {
                if let Value::String(s) = c {
                    match s.as_ref() {
                        "number" => number_val = Some(ci),
                        "string" => string_val = Some(ci),
                        _ => {}
                    }
                }
            }
            // null → return "undefined" placeholder (ref.null extern for now)
            body.push(0xD0);
            body.push(0x6F); // ref.null extern = "undefined"
            body.push(0x05); // else (not null)
            // Check if number
            body.push(0x20);
            write_leb128_u32(body, temp_idx); // value
            emit_import_call(body, rt_idx, "wasm:js-number", "test"); // → i32
            body.push(0x04);
            body.push(TYPE_EXTERNREF); // if number (result externref)
            // Return "number" constant if available, else boxed tag
            if let Some(ci) = number_val {
                // Emit the string constant from the chunk's pool
                emit_string_const(body, chunk, ci, rt_idx);
            } else {
                body.push(0x41);
                write_leb128_i32(body, 1);
                emit_box_i32(body, rt_idx);
            }
            body.push(0x05); // else (not number)
            // Check if string
            body.push(0x20);
            write_leb128_u32(body, temp_idx);
            emit_import_call(body, rt_idx, "wasm:js-string", "test");
            body.push(0x04);
            body.push(TYPE_EXTERNREF); // if string
            if let Some(ci) = string_val {
                emit_string_const(body, chunk, ci, rt_idx);
            } else {
                body.push(0x41);
                write_leb128_i32(body, 2);
                emit_box_i32(body, rt_idx);
            }
            body.push(0x05); // else
            // Default: "object"
            body.push(0x41);
            write_leb128_i32(body, 3);
            emit_box_i32(body, rt_idx);
            body.push(0x0B); // end if string
            body.push(0x0B); // end if number
            body.push(0x0B); // end if null
        }
        _ if op == Op::REF_IS_OBJECT || op == Op::REF_IS_FUNC || op == Op::REF_IS_ARRAY => {
            // TODO: proper type checks
            body.push(0x1A); // drop value
            body.push(0x41);
            write_leb128_i32(body, 0); // push false (i32 0)
            emit_box_i32(body, rt_idx);
        }
        // Stack ops
        _ if op == Op::DUP => {
            // Duplicate TOS: local.tee $temp, local.get $temp
            body.push(0x22);
            write_leb128_u32(body, temp_idx); // local.tee $temp
            body.push(0x20);
            write_leb128_u32(body, temp_idx); // local.get $temp
        }
        // Exception handling
        _ if op == Op::TRY_START => {
            let _ = read_u16(&chunk.code, ip);
            let _ = read_u16(&chunk.code, ip);
            body.push(0x01);
        } // nop for now
        _ if op == Op::TRY_END => {
            body.push(0x01);
        } // nop for now
        // Spread — TODO: inline impl
        _ if op == Op::SPREAD => {
            body.push(0x01);
        } // nop
        // Set timer — TODO: needs host import (not stdlib)
        _ if op == Op::SET_TIMER => {
            body.push(0x01);
        } // nop
        // Upvalue get/set — closures use WASM function references
        _ if op == Op::UPVALUE_GET => {
            let _idx = chunk.code[*ip];
            *ip += 1;
            // TODO: proper closure/upvalue via WASM funcref + tables
            body.push(0xD0);
            body.push(0x6F); // ref.null extern (placeholder)
        }
        _ if op == Op::UPVALUE_SET => {
            let _idx = chunk.code[*ip];
            *ip += 1;
            // TODO: proper closure/upvalue via WASM funcref + tables
            // drop the value being set
            body.push(0x1A); // drop
            body.push(0xD0);
            body.push(0x6F); // ref.null extern (placeholder return)
        }
        // Memory64 reuses the standard memory instruction bytes. The i64
        // address shape is carried by the memory type in the binary format.
        _ if op == Op::I64_MEMORY_SIZE => {
            body.push(Op::MEMORY_SIZE.sub_u8());
            let memidx = read_optional_memidx_immediate(chunk, ip);
            write_leb128_u32(body, memidx);
        }
        _ if op == Op::I64_MEMORY_GROW => {
            body.push(Op::MEMORY_GROW.sub_u8());
            let memidx = read_optional_memidx_immediate(chunk, ip);
            write_leb128_u32(body, memidx);
        }
        _ if op == Op::I64_MEMORY_COPY => {
            body.push(0xFC);
            write_leb128_u32(body, Op::MEMORY_COPY.sub() as u32);
            let dst_mem = read_optional_memidx_immediate(chunk, ip);
            let src_mem = read_optional_memidx_immediate(chunk, ip);
            write_leb128_u32(body, dst_mem);
            write_leb128_u32(body, src_mem);
        }
        _ if op == Op::I64_MEMORY_FILL => {
            body.push(0xFC);
            write_leb128_u32(body, Op::MEMORY_FILL.sub() as u32);
            let memidx = read_optional_memidx_immediate(chunk, ip);
            write_leb128_u32(body, memidx);
        }
        _ if op == Op::I32_LOAD_64 => emit_memory64_op(body, chunk, ip, Op::I32_LOAD.sub_u8(), 2),
        _ if op == Op::I64_LOAD_64 => emit_memory64_op(body, chunk, ip, Op::I64_LOAD.sub_u8(), 3),
        _ if op == Op::F32_LOAD_64 => emit_memory64_op(body, chunk, ip, Op::F32_LOAD.sub_u8(), 2),
        _ if op == Op::F64_LOAD_64 => emit_memory64_op(body, chunk, ip, Op::F64_LOAD.sub_u8(), 3),
        _ if op == Op::I32_LOAD8_S_64 => {
            emit_memory64_op(body, chunk, ip, Op::I32_LOAD8_S.sub_u8(), 0)
        }
        _ if op == Op::I32_LOAD8_U_64 => {
            emit_memory64_op(body, chunk, ip, Op::I32_LOAD8_U.sub_u8(), 0)
        }
        _ if op == Op::I32_LOAD16_S_64 => {
            emit_memory64_op(body, chunk, ip, Op::I32_LOAD16_S.sub_u8(), 1)
        }
        _ if op == Op::I32_LOAD16_U_64 => {
            emit_memory64_op(body, chunk, ip, Op::I32_LOAD16_U.sub_u8(), 1)
        }
        _ if op == Op::I64_LOAD8_S_64 => {
            emit_memory64_op(body, chunk, ip, Op::I64_LOAD8_S.sub_u8(), 0)
        }
        _ if op == Op::I64_LOAD8_U_64 => {
            emit_memory64_op(body, chunk, ip, Op::I64_LOAD8_U.sub_u8(), 0)
        }
        _ if op == Op::I64_LOAD16_S_64 => {
            emit_memory64_op(body, chunk, ip, Op::I64_LOAD16_S.sub_u8(), 1)
        }
        _ if op == Op::I64_LOAD16_U_64 => {
            emit_memory64_op(body, chunk, ip, Op::I64_LOAD16_U.sub_u8(), 1)
        }
        _ if op == Op::I64_LOAD32_S_64 => {
            emit_memory64_op(body, chunk, ip, Op::I64_LOAD32_S.sub_u8(), 2)
        }
        _ if op == Op::I64_LOAD32_U_64 => {
            emit_memory64_op(body, chunk, ip, Op::I64_LOAD32_U.sub_u8(), 2)
        }
        _ if op == Op::I32_STORE_64 => emit_memory64_op(body, chunk, ip, Op::I32_STORE.sub_u8(), 2),
        _ if op == Op::I64_STORE_64 => emit_memory64_op(body, chunk, ip, Op::I64_STORE.sub_u8(), 3),
        _ if op == Op::F32_STORE_64 => emit_memory64_op(body, chunk, ip, Op::F32_STORE.sub_u8(), 2),
        _ if op == Op::F64_STORE_64 => emit_memory64_op(body, chunk, ip, Op::F64_STORE.sub_u8(), 3),
        _ if op == Op::I32_STORE8_64 => emit_memory64_op(body, chunk, ip, Op::I32_STORE8.sub_u8(), 0),
        _ if op == Op::I32_STORE16_64 => {
            emit_memory64_op(body, chunk, ip, Op::I32_STORE16.sub_u8(), 1)
        }
        _ if op == Op::I64_STORE8_64 => emit_memory64_op(body, chunk, ip, Op::I64_STORE8.sub_u8(), 0),
        _ if op == Op::I64_STORE16_64 => {
            emit_memory64_op(body, chunk, ip, Op::I64_STORE16.sub_u8(), 1)
        }
        _ if op == Op::I64_STORE32_64 => {
            emit_memory64_op(body, chunk, ip, Op::I64_STORE32.sub_u8(), 2)
        }
        _ if op == Op::TABLE_GET_64 => emit_table64_core_op(body, chunk, ip, Op::TABLE_GET.sub_u8()),
        _ if op == Op::TABLE_SET_64 => emit_table64_core_op(body, chunk, ip, Op::TABLE_SET.sub_u8()),
        _ if op == Op::TABLE_INIT_64 => {
            emit_table64_fc_op(body, chunk, ip, Op::TABLE_INIT.sub_u8(), true)
        }
        _ if op == Op::TABLE_COPY_64 => {
            emit_table64_fc_op(body, chunk, ip, Op::TABLE_COPY.sub_u8(), true)
        }
        _ if op == Op::TABLE_GROW_64 => {
            emit_table64_fc_op(body, chunk, ip, Op::TABLE_GROW.sub_u8(), false)
        }
        _ if op == Op::TABLE_SIZE_64 => {
            emit_table64_fc_op(body, chunk, ip, Op::TABLE_SIZE.sub_u8(), false)
        }
        _ if op == Op::TABLE_FILL_64 => {
            emit_table64_fc_op(body, chunk, ip, Op::TABLE_FILL.sub_u8(), false)
        }
        // Set type ID — GC type stamps handled by WASM GC type system
        _ if op == Op::SET_TYPE_ID => {
            body.push(0x01);
        } // nop (type is structural in GC)
        _ if op == Op::HALT => {
            body.push(0x0F);
        } // return (not unreachable — _start should return cleanly)
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
                OperandFormat::TryTable => {
                    let count = chunk.code.get(*ip).copied().unwrap_or(0) as usize;
                    *ip += 1 + count * 3;
                }
                _ => {
                    *ip += fmt.size_in(&chunk.code, *ip);
                }
            }
            body.push(0x01); // nop
        }
    }
}

/// Total instruction size: 2-byte opcode + operand bytes.
pub fn opcode_size(op: Op, code: &[u8], ip: usize) -> usize {
    let base = 2;
    base + op.operand_format().size_in(code, ip + base)
}
