//! WASM code section encoding.
//! Translates internal bytecode to WASM binary format.
//!
//! Type strategy: externref is the universal value representation.
//! All locals, params, and function results are externref.
//! Typed WASM ops (f64.add, i32.mul, etc.) require unboxing via
//! wasm:js-number builtins (toF64/toI32) before the op and reboxing
//! (fromF64/fromI32) after. Binary ops need a temp externref local
//! to save TOS while unboxing the second operand.

use crate::encoding::*;
use crate::writer::sections::{emit_box_f64, emit_box_i32, emit_unbox_f64, emit_unbox_i32};
use crate::writer::types::WasmTypeContext;
use vybe_runtime::opcode::{OperandFormat, read_leb_u32};
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op};

// The `TryRegion` machinery that used to live here is DELETED.
//
// It reconstructed `block $after; block $catch; … end` around a `try_table` by
// reading the clause's third field as a signed BYTE OFFSET and doing pointer
// arithmetic with it. That field is a spec `labelidx` — a block depth — so the
// arithmetic produced a garbage `catch_ip`, and because the width is unchanged
// nothing errored: the emitted `.wasm` was simply wrong.
//
// The reconstruction is also no longer needed. The compiler emits the handler
// block for real (see `errors::emit_try_start`), so the structure the writer
// used to invent is already present in the bytecode and the ordinary
// BLOCK/END/BR paths translate it. Synthesizing it again would double-wrap.

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

/// The locals declaration for a chunk, as `(count, valtype)` groups in index
/// order — THE single source of truth for both the bytes and their size.
///
/// ⛔ THE EMITTER AND `locals_prefix_size` MUST NOT COMPUTE THIS TWICE. The
/// spec makes branch-hint offsets relative to the end of this prefix, so a
/// disagreement of one byte silently misplaces every hint in the function.
/// They were two constants agreeing by luck while every local was externref;
/// `i64_locals` makes the shape depend on the chunk, and two readers of one
/// fact is the exact split that has cost this session all day.
fn locals_groups(chunk: &Chunk) -> Vec<(u32, u8)> {
    let wasm_params = chunk.arity as u32;
    let extra_locals = (chunk.local_count as u32).saturating_sub(wasm_params);
    let temp_count = count_temp_locals(chunk);
    let mut groups: Vec<(u32, u8)> = Vec::new();
    let mut push = |ty: u8, groups: &mut Vec<(u32, u8)>| match groups.last_mut() {
        Some((n, t)) if *t == ty => *n += 1,
        _ => groups.push((1, ty)),
    };
    // Declared locals are indices `wasm_params .. wasm_params+extra_locals`,
    // and slot k IS local index k (slots below arity are the params).
    for i in 0..extra_locals {
        let slot = wasm_params + i;
        let ty = if u16::try_from(slot).is_ok_and(|s| chunk.i64_locals.contains(&s)) {
            TYPE_I64
        } else {
            TYPE_EXTERNREF
        };
        push(ty, &mut groups);
    }
    for _ in 0..temp_count {
        push(TYPE_EXTERNREF, &mut groups);
    }
    groups
}

/// Write the locals prefix from [`locals_groups`].
fn write_locals_prefix(body: &mut Vec<u8>, chunk: &Chunk) {
    let groups = locals_groups(chunk);
    write_leb128_u32(body, groups.len() as u32);
    for (count, ty) in groups {
        write_leb128_u32(body, count);
        body.push(ty);
    }
}

/// Size in bytes of the locals declaration prefix written at the start of a
/// function body. The spec says branch hint offsets are relative to this
/// position, so the hint scanner adds this value to every `chunk.code` offset.
pub(crate) fn locals_prefix_size(chunk: &Chunk) -> u32 {
    let groups = locals_groups(chunk);
    leb128_u32_size(groups.len() as u32)
        + groups
            .iter()
            .map(|&(count, _)| leb128_u32_size(count) + 1)
            .sum::<u32>()
}

/// Count how many temp locals a chunk needs for stack manipulation.
/// Returns 0, 1, or 2 depending on which ops are used.
/// Read a big-endian `u16` immediate at a fixed position, without moving an
/// instruction pointer — for scans that only need to peek at an operand.
fn read_u16_at(code: &[u8], at: usize) -> u16 {
    match (code.get(at), code.get(at + 1)) {
        (Some(hi), Some(lo)) => ((*hi as u16) << 8) | *lo as u16,
        _ => 0,
    }
}

fn count_temp_locals(chunk: &Chunk) -> u32 {
    let mut need = 0u32;
    let mut ip = 0;
    while ip < chunk.code.len() {
        if ip + 3 >= chunk.code.len() {
            break;
        }
        if let Some(op) = Op::decode(
            ((chunk.code[ip] as u16) << 8) | chunk.code[ip + 1] as u16,
            ((chunk.code[ip + 2] as u16) << 8) | chunk.code[ip + 3] as u16,
        ) {
            if op == Op::CALL_REF || op == Op::RETURN_CALL || op == Op::RETURN_CALL_REF {
                // Dynamic callee-on-stack calls need argc+1 temps
                // (save args + the funcref) to reorder into the spec
                // `call_indirect`/`return_call_indirect` stack shape.
                let call_argc = chunk.code.get(ip + 4).copied().unwrap_or(0) as u32;
                need = need.max(call_argc + 1);
            } else if op == Op::STRUCT_NEW
                && read_u16_at(&chunk.code, ip + 4) == 0
                && read_u16_at(&chunk.code, ip + 6) != 0
            {
                // Dynamic object built from key/value pairs already on the
                // stack: the object has to exist before any of them can be
                // stored, so each pair is unstacked through temps.
                need = need.max(3);
            } else if memory_store_shape(op).is_some() {
                // One temp to lift the value off the address underneath it.
                need = need.max(1);
            } else if op == Op::CALL {
                // Spill temps for `emit_unbox_raw_params` — an import with
                // a raw param under another argument. Reserved for every
                // static call because the import tables are not built yet
                // here; two externref locals are cheaper than threading them in.
                need = need.max(2);
            } else if op == Op::ARRAY_SET || op == Op::STRUCT_SET {
                need = need.max(2); // need 2 temps for 3-operand reorder
            } else if op == Op::DESC_SET_PROTO {
                // Expands to local.set / global.get / local.get / struct.set —
                // one temp to get the descriptor ref UNDER the prototype value.
                need = need.max(1);
            } else if is_binary_typed_op(op)
                || op == Op::ARRAY_GET
                || op == Op::ARRAY_LENGTH
                || op == Op::REF_IS_NULL
                || op == Op::REF_EQ
            {
                need = need.max(1);
            }
            // Phase E: the `0xFF` ARRAY_* (PUSH/POP/SLICE/JOIN/REVERSE/
            // CONTAINS/INDEX_OF/CONTAINS/CONCAT/SHIFT) no longer exist.
            // Their call-site temps are no longer needed here either.
            ip += opcode_size(op, &chunk.code, ip);
        } else {
            ip += 4;
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

// (The 0xEE selector reader is deleted: memory.size/grow/fill/copy/init
// carry fixed u16 memidx immediates declared in their OperandFormat.)

fn read_optional_memarg(chunk: &Chunk, ip: &mut usize, default_align: u32) -> (u32, u64, u32) {
    // OPTIONAL marker-tagged memarg (`SimdMemArg` treatment) — must mirror
    // the VM's `read_optional_memarg` exactly: present iff the first LEB
    // carries 0x80 (instruction group-hi bytes are always 0x00, so the peek
    // is unambiguous — no opcode-decode guessing); 0x100 = memory64 offset;
    // 0x40 = memidx LEB follows. Absent → the op's natural alignment,
    // offset 0, memory 0 — the spec binary always writes a memarg.
    let mut probe = *ip;
    let marker_align = read_leb_u32(&chunk.code, &mut probe);
    if marker_align & 0x80 == 0 {
        return (default_align, 0, 0);
    }
    *ip = probe;
    let offset = if marker_align & 0x100 != 0 {
        read_leb_u64(&chunk.code, ip)
    } else {
        read_leb_u32(&chunk.code, ip) as u64
    };
    let memidx = if marker_align & 0x40 != 0 {
        read_leb_u32(&chunk.code, ip)
    } else {
        0
    };
    (marker_align & !0x1C0, offset, memidx)
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
    // ⚠ THE INDEX MAP FIRST, and it is not an optimisation.
    //
    // The name path below can only work while encoding chunk 0: the type
    // table is written to `chunks[0].types` exclusively, and
    // `encode_code_section` walks every chunk. In a constructor — a non-zero
    // chunk — `chunk.types` is empty, the lookup misses, and the
    // `unwrap_or` hands back the CHUNK type index as if it were a WASM type
    // index. Two different spaces, no diagnostic.
    //
    // `struct_type_by_module_index` answers the same question without a name
    // and without a chunk, which is why it exists.
    type_ctx
        .struct_type_by_index(typeidx as u32)
        .or_else(|| {
            chunk
                .types
                .get(typeidx as usize)
                .and_then(|ty| type_ctx.struct_type(&ty.name))
        })
        .unwrap_or(typeidx as u32)
}

/// The descriptor singleton global for the class a chunk typeidx names, or
/// `None` when the typeidx is the dynamic form (0) or names nothing.
///
/// A `None` here means "allocate with the plain instruction", which is only
/// correct for a type with no `(descriptor …)` clause.
fn descriptor_global_for_chunk_type(
    chunk: &Chunk,
    type_ctx: &WasmTypeContext,
    typeidx: u16,
) -> Option<u32> {
    if typeidx == 0 {
        return None;
    }
    // ⚠ BY INDEX, NOT BY NAME. The type table lives on `chunks[0]` alone, and
    // `encode_code_section` walks every chunk — so `chunk.types` is EMPTY in
    // every chunk but the first, which is exactly where constructors live. A
    // name lookup here resolved for nothing and the descriptor was never
    // found. The module type index space is chunk-independent, so this asks
    // the question that has an answer in every chunk.
    let _ = chunk;
    type_ctx.desc_global_by_index(typeidx as u32)
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

/// Resolve a struct field from the TYPE THE BYTECODE ALREADY NAMES.
///
/// ⛔ THE TYPEIDX WAS BEING READ AND DISCARDED. Both `STRUCT_GET` and
/// `STRUCT_SET` carry `(typeidx, field-name-constant)`, and the writer threw the
/// first away and searched EVERY type in the module for a field of that name —
/// `wasm_struct_field_for_name` below. Since `constructor`, `__type`, `__types`
/// and `name` exist on many classes, ambiguity is the ordinary case, and its
/// answer is the `(0, 0)` sentinel: struct type 0, field 0. Measured on
/// `class A { constructor(x){ this.x = x; } m(){ return this.x; } }` — **897 of
/// 902** struct immediates in the emitted binary decoded to `(0, 0)`.
///
/// `(0, 0)` is not a safe default: index 0 is a REAL type, so a miss silently
/// names something valid and the module is invalid rather than rejected. V8:
/// `invalid field index: 0`, or `struct.set: Field 2 of type 3 is immutable`
/// when the index runs past the described struct into the paired descriptor's
/// `GC_IMMUT` vtable slots.
///
/// A typeidx of 0 means genuinely dynamic (a JS object literal, a lua table) and
/// has no declared type to resolve against, so it returns `None` and the caller
/// falls back. `None` is also returned when the named field is not in that
/// type's list — which today is most synthesized slots, because
/// `TypeEntry.fields` carries declared fields and auto-property backing fields
/// only. That gap is the slot census (`class_slots.rs`), and when it closes this
/// path answers for everything and the name search can go.
/// The GLOBAL index of the string constant naming a dynamic property, if the
/// module imports it. `None` when the constant is not a string or was not
/// collected, in which case the caller falls back to the struct path rather
/// than emitting a `global.get` at an index that means something else.
fn dynamic_prop_name_global(
    chunk: &Chunk,
    type_ctx: &WasmTypeContext,
    field_name_idx: u16,
) -> Option<u32> {
    let value = chunk.constants.get(field_name_idx as usize)?;
    let text = format!("{value}");
    type_ctx.string_const_global.get(&text).copied()
}

fn wasm_struct_field_by_typeidx(
    chunk: &Chunk,
    type_ctx: &WasmTypeContext,
    typeidx: u16,
    field_name_idx: u16,
) -> Option<(u32, u32)> {
    if typeidx == 0 {
        return None;
    }
    // The operand is 1-based over `chunk.types`, matching `TypeEntry.parent_index`.
    let ty = chunk.types.get(typeidx as usize - 1)?;
    let value = chunk.constants.get(field_name_idx as usize)?;
    let field_name = format!("{value}");
    let field_idx = ty.fields.iter().position(|field| field == &field_name)?;
    Some((type_ctx.struct_type(&ty.name)?, field_idx as u32))
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

/// Write a blocktype immediate. The shorthands only spell "no results" and "one
/// value"; a block that takes operands or yields several must use the s33
/// typeidx form, which `build_type_context` pre-declared for every
/// (params, results) shape the bytecode contains.
fn write_blocktype(body: &mut Vec<u8>, params: u8, results: u8, type_ctx: &WasmTypeContext) {
    if params == 0 && results == 0 {
        body.push(TYPE_VOID);
        return;
    }
    if params == 0 && results == 1 {
        body.push(TYPE_EXTERNREF);
        return;
    }
    match type_ctx.block_type_by_results.get(&(params, results)) {
        // s33, the same signed-LEB form the BLOCK path writes.
        Some(&tidx) => write_leb128_i32(body, tidx as i32),
        // Unreachable for bytecode this writer scanned; degrade to the closest
        // shorthand rather than emitting a dangling typeidx.
        None => body.push(if results == 0 { TYPE_VOID } else { TYPE_EXTERNREF }),
    }
}

pub fn encode_code_section(
    chunks: &[Chunk],
    rt_imports: &[(&str, &str)],
    type_ctx: &WasmTypeContext,
    tag_plan: &crate::writer::proposals::exception_handling::ModuleTagPlan,
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

    for (ci, chunk) in chunks.iter().enumerate() {
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

        // Locals declaration (only declare locals beyond params) — grouped by
        // type, because `i64_locals` makes them heterogeneous.
        write_locals_prefix(&mut body, chunk);

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
        // Translate opcodes
        let mut ip = 0;
        let mut i32_block_stack: Vec<bool> = Vec::new();
        while ip < chunk.code.len() {
            if ip + 3 >= chunk.code.len() {
                break;
            }
            let op = match Op::decode(
                ((chunk.code[ip] as u16) << 8) | chunk.code[ip + 1] as u16,
                ((chunk.code[ip + 2] as u16) << 8) | chunk.code[ip + 3] as u16,
            ) {
                Some(op) => op,
                None => {
                    ip += 4;
                    continue;
                }
            };

            // TRY_TABLE → a real spec try_table, clause for clause.
            //
            // This used to SYNTHESIZE `block $after; block $catch; …` around
            // the instruction, because the internal form had no handler block
            // — the clause carried a patched byte offset and the structure had
            // to be reconstructed from it. The compiler now emits that shape
            // natively (the clause names a `labelidx`, and the block whose
            // `end` is the handler is really there), so synthesizing it again
            // would DOUBLE-WRAP the region. The surrounding BLOCK/END/BR
            // opcodes translate through the ordinary paths below.
            //
            // Writing the clauses that are actually present also lifts the old
            // single-clause restriction: a multi-clause `try_table` from wast
            // used to fall through to the nop-skip and emit `0x01`.
            if op == Op::TRY_TABLE {
                // [params, results, u16 clause_count, N×(kind, tag, label)]
                let params = chunk.code[ip + 4];
                let results = chunk.code[ip + 5];
                let clause_count =
                    (((chunk.code[ip + 6] as usize) << 8) | chunk.code[ip + 7] as usize) as usize;
                body.push(0x1F); // try_table
                // The real blocktype. The shorthands only cover void and a
                // single value; anything that takes operands or yields several
                // needs the s33 typeidx form, and writing `externref` for those
                // silently changed the block's type.
                write_blocktype(&mut body, params, results, type_ctx);
                write_leb128_u32(&mut body, clause_count as u32);
                for i in 0..clause_count {
                    let base = ip + 8 + i * 5;
                    let kind = chunk.code[base];
                    let tag = ((chunk.code[base + 1] as u16) << 8) | chunk.code[base + 2] as u16;
                    let label = ((chunk.code[base + 3] as u32) << 8) | chunk.code[base + 4] as u32;
                    body.push(kind);
                    // catch / catch_ref carry a tagidx; the catch_all kinds do
                    // not. The clause's own tag, mapped into the module's tag
                    // index space — writing the shared exception tag here fused
                    // every tag into one and broke the matching rule.
                    if kind == 0x00 || kind == 0x01 {
                        write_leb128_u32(&mut body, tag_plan.module_tag(ci, tag));
                    }
                    write_leb128_u32(&mut body, label);
                }
                // opcode(4) + params(1) + results(1) + clause_count(2)
                //   + N·[kind, tag(2), label(2)]
                ip += 8 + clause_count * 5;
                continue;
            }
            // The `end` closing a try_table and the success-path `br` that
            // skips the handler need no special case any more: both are
            // ORDINARY structured instructions in the emitted bytecode now, so
            // the generic paths below translate them. The old code rewrote the
            // `br` to a hardcoded `br 1` because the writer had invented the
            // blocks it was branching past; the compiler emits the real depth.
            let op_start = ip;
            ip += 4;

            // ⛔ RAW-REGION TRACKING, NOT A TYPE CHECKER. `block_i32_results`
            // is the compiler naming the blocks whose result is a bare i32 —
            // the truthiness ladder and nothing else. All this stack answers is
            // "is the value I am about to emit an arm result of one of those",
            // so the i32 producers inside a ladder are left unboxed.
            let in_i32_block = i32_block_stack.last().copied().unwrap_or(false);
            let mut closed_i32_region = false;
            if op == Op::BLOCK || op == Op::IF || op == Op::LOOP {
                i32_block_stack.push(chunk.block_i32_results.contains(&op_start));
            } else if op == Op::END {
                // ⛔ THE RAW REGION HAS TO REJOIN THE VALUE ABI. A truthiness
                // ladder is a bare i32 throughout — that is what
                // `block_i32_results` declares — but at its OUTERMOST `end` the
                // answer stops being an arm result and becomes an ordinary
                // value: stored to a local, passed to a call, returned. All of
                // those are externref, and the i32 was reaching a `local.set`
                // on an externref local.
                //
                // Only the outermost `end` boxes. An inner one is still inside
                // the ladder, and a result feeding `if`/`br_if` is wanted raw —
                // the same question `box_i32_unless_condition` answers.
                let closed = i32_block_stack.pop().unwrap_or(false);
                closed_i32_region = closed && !i32_block_stack.last().copied().unwrap_or(false);
            }

            if op.group() == 0x00 && !op.is_vm_internal() {
                emit_core_op(
                    &mut body,
                    op,
                    chunk,
                    ci,
                    &mut ip,
                    op_start,
                    &rt_idx,
                    temp_local_idx,
                    has_temp,
                    type_ctx,
                    host_import_count,
                    tag_plan,
                    in_i32_block,
                );
            } else if op.group() == 0xFB {
                emit_gc_op(
                    &mut body,
                    op,
                    chunk,
                    &mut ip,
                    &rt_idx,
                    type_ctx,
                    temp_local_idx,
                    in_i32_block,
                );
            } else if op.group() == 0xFC {
                // 0xFC-prefix ops per the bulk-memory / reference-types spec
                // need specific trailing immediates that aren't captured by
                // our `operand_format` in bytecode. Translate each case.
                body.push(op.group() as u8);
                write_leb128_u32(&mut body, op.sub() as u32);
                match op {
                    Op::MEMORY_INIT => {
                        // spec: data_idx, memory_idx (internal: u16 BE + u16 BE)
                        let data_idx = read_u16(&chunk.code, &mut ip);
                        let memidx = read_u16(&chunk.code, &mut ip);
                        write_leb128_u32(&mut body, data_idx as u32);
                        write_leb128_u32(&mut body, memidx as u32);
                    }
                    Op::DATA_DROP => {
                        let data_idx = read_u16(&chunk.code, &mut ip);
                        write_leb128_u32(&mut body, data_idx as u32);
                    }
                    Op::MEMORY_COPY => {
                        // spec: dst_mem, src_mem
                        let dst_mem = read_u16(&chunk.code, &mut ip);
                        let src_mem = read_u16(&chunk.code, &mut ip);
                        write_leb128_u32(&mut body, dst_mem as u32);
                        write_leb128_u32(&mut body, src_mem as u32);
                    }
                    Op::MEMORY_FILL => {
                        let memidx = read_u16(&chunk.code, &mut ip);
                        write_leb128_u32(&mut body, memidx as u32);
                    }
                    Op::TABLE_INIT => {
                        let elem_idx = read_u16(&chunk.code, &mut ip);
                        let table_idx = read_u16(&chunk.code, &mut ip);
                        write_leb128_u32(&mut body, elem_idx as u32);
                        write_leb128_u32(&mut body, table_idx as u32);
                    }
                    Op::ELEM_DROP => {
                        let elem_idx = read_u16(&chunk.code, &mut ip);
                        write_leb128_u32(&mut body, elem_idx as u32);
                    }
                    Op::TABLE_COPY => {
                        let dst_table = read_u16(&chunk.code, &mut ip);
                        let src_table = read_u16(&chunk.code, &mut ip);
                        write_leb128_u32(&mut body, dst_table as u32);
                        write_leb128_u32(&mut body, src_table as u32);
                    }
                    Op::TABLE_GROW | Op::TABLE_SIZE | Op::TABLE_FILL => {
                        let table_idx = read_u16(&chunk.code, &mut ip);
                        write_leb128_u32(&mut body, table_idx as u32);
                    }
                    _ => {
                        ip += op.operand_format().size_in(&chunk.code, ip);
                    }
                }
            } else if op.group() == 0xFD {
                emit_simd_prefixed_op(&mut body, op, chunk, &mut ip);
            } else if op.group() == 0xFE {
                emit_thread_prefixed_op(
                    &mut body,
                    op,
                    chunk,
                    &mut ip,
                    &rt_idx,
                    in_i32_block,
                );
            } else {
                emit_vm_internal_op(&mut body, op, chunk, &mut ip);
            }

            if closed_i32_region {
                box_i32_unless_condition(&mut body, &rt_idx, chunk, ip, false);
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
#[allow(clippy::too_many_arguments)]
fn emit_core_op(
    body: &mut Vec<u8>,
    op: Op,
    chunk: &Chunk,
    ci: usize,
    ip: &mut usize,
    op_start: usize,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    temp_idx: u32,
    _has_temp: bool,
    type_ctx: &WasmTypeContext,
    _host_import_count: usize,
    tag_plan: &crate::writer::proposals::exception_handling::ModuleTagPlan,
    in_i32_block: bool,
) {
    match op {
        _ if op == Op::LOCAL_GET => {
            body.push(op.sub() as u8);
            write_leb128_u32(body, read_u16(&chunk.code, ip) as u32);
        }
        _ if op == Op::LOCAL_SET => {
            // ⛔ `local.set` IS NOT `local.tee`. This emitted `0x22` (tee) on
            // the stated grounds that "our VM keeps the value for both" — it
            // does not: `dispatch.rs` `LOCAL_SET` is `let val = self.pop()`
            // while `LOCAL_TEE` is `self.peek(0).clone()`. Every `local.set` we
            // emitted therefore left a value the bytecode had consumed, and the
            // imbalance surfaced at whichever block END came next
            // ("values remaining on stack at end of block") — never at the
            // instruction that caused it. The VM could not see it: it runs
            // Chunks, and in the Chunks the set really does pop.
            body.push(0x21);
            write_leb128_u32(body, read_u16(&chunk.code, ip) as u32);
        }
        // ⛔⛔ `local.tee` HAD NO ARM AT ALL and fell through to the catch-all,
        // which pushes the opcode byte and DISCARDS the operand. The emitted
        // `0x22` then swallowed the FOLLOWING instruction's first byte as its
        // own local index, and every byte after that decoded one position out.
        //
        // Measured on `(module)` — the empty wast module, which needs nothing
        // but the scaffold: `22 | 22 37 | 24 08` where `22 37 24 08` was meant
        // to be `local.tee 55; global.set 8`. V8 read `local.tee 34`, then took
        // `0x37` as `i64.store` and its align byte as a memory ordering. That
        // is the whole of `invalid memory ordering: acquire-release requires
        // --experimental-wasm-acquire-release`, and it is why NO module we
        // emitted — in ANY language — could be decoded past its first
        // `local.tee`.
        //
        // `LOCAL_TEE` keeps its value in the VM and in the spec alike, so this
        // one really is `local.tee`. See the arm above for why `LOCAL_SET` is
        // not.
        _ if op == Op::LOCAL_TEE => {
            body.push(0x22);
            write_leb128_u32(body, read_u16(&chunk.code, ip) as u32);
        }
        // Spec `call`: u16 funcidx + VM-internal u8 argc. The argc byte is
        // dropped — the .wasm binary carries only LEB(funcidx). Imports
        // occupy the front of the module's function index space, so the
        // chunk-scoped import index is already the module-level funcidx.
        _ if op == Op::CALL => {
            let funcidx = read_u16(&chunk.code, ip);
            let argc = chunk.code[*ip];
            *ip += 1;
            emit_unbox_raw_params(body, rt_idx, type_ctx, funcidx as u32, argc, temp_idx);
            body.push(0x10);
            write_leb128_u32(body, funcidx as u32);
            // ⛔ A TRUTHFUL SIGNATURE HAS TWO HALVES. Declaring
            // `wasm:js-string.test` as `-> i32` is only half of it: its result
            // then sits on a stack whose every other value is externref, and
            // the first consumer that is not an `if` rejects it. Box here, and
            // let the peephole drop the box again when the consumer really
            // does want the raw i32.
            match type_ctx.raw_result_funcs.get(&(funcidx as u32)) {
                Some(&crate::encoding::TYPE_I32) => {
                    box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block)
                }
                // ⛔ f64 RESULTS NEED IT TOO. `wasm:js-number.toF64` really
                // returns f64, and the truthiness ladder stores that straight
                // into a local — a local this writer declares `externref`,
                // like every other. Boxing here is what keeps the ladder's
                // `f64.ne` reachable through `emit_binary_f64_cmp`, which
                // unboxes its operands anyway.
                Some(&crate::encoding::TYPE_F64) => emit_box_f64(body, rt_idx),
                _ => {}
            }
        }
        // `call_ref`: one `u8` argc, callee on the stack, lowered to
        // call_indirect through the function table. (The old `Op::CALL`
        // alias for this shape is retired; spec `call` is a static import
        // call — see callimportretirement.md.)
        //
        // `return_call` / `return_call_ref` share the internal shape
        // (u8 argc, callee below the args) and the exact same staging;
        // being tail calls they lower to spec `return_call_indirect`
        // (0x13, tail-call proposal) instead of `call_indirect`.
        _ if op == Op::CALL_REF || op == Op::RETURN_CALL || op == Op::RETURN_CALL_REF => {
            let spec_byte: u8 = if op == Op::CALL_REF { 0x11 } else { 0x13 };
            let argc = chunk.code[*ip];
            *ip += 1;
            let results = chunk.code[*ip];
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
            // 5. call_indirect / return_call_indirect with the EXACT
            // functype (argc externrefs -> results externrefs, from the
            // op's own immediates); first-seen-arity is the fallback for
            // pre-registry chunks.
            // ⛔ DECLARED FUNCTYPE FIRST. The three immediates are counts, and
            // a count cannot choose between two same-arity types — with real
            // functypes emitted, `func_type_by_arity` (`.or_insert`, first seen
            // wins) would hand an argc-1 call site whichever argc-1 type came
            // first, and `call_indirect`'s immediate must EQUAL the callee's
            // functype or a conforming engine traps. The side table carries the
            // call site's own `(type $t)` / `(param…)(result…)`.
            if let Some(&type_idx) = chunk
                .call_indirect_sigs
                .get(&op_start)
                .and_then(|sig| type_ctx.func_type_by_signature.get(sig))
                .or_else(|| type_ctx.block_type_by_results.get(&(argc, results)))
                .or_else(|| type_ctx.func_type_by_arity.get(&argc))
            {
                body.push(spec_byte);
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
        _ if op == Op::CALL_INDIRECT || op == Op::RETURN_CALL_INDIRECT => {
            let spec_byte: u8 = if op == Op::CALL_INDIRECT { 0x11 } else { 0x13 };
            let argc = chunk.code[*ip];
            *ip += 1;
            let table_idx = chunk.code[*ip];
            *ip += 1;
            // Third immediate: the expected result count. The spec
            // `(type $sig)` annotation must carry it exactly — a
            // first-seen-arity functype with a different result count is a
            // structural mismatch (traps in a conforming engine).
            let results = chunk.code[*ip];
            *ip += 1;
            // Declared functype first — see the sibling site above.
            if let Some(&type_idx) = chunk
                .call_indirect_sigs
                .get(&op_start)
                .and_then(|sig| type_ctx.func_type_by_signature.get(sig))
                .or_else(|| type_ctx.block_type_by_results.get(&(argc, results)))
                .or_else(|| type_ctx.func_type_by_arity.get(&argc))
            {
                body.push(spec_byte);
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
        // BLOCK/LOOP/IF carry (param_count, result_count) bytes. Translate
        // to the spec blocktype: 0x40 void / one valtype / positive s33
        // typeidx into the pre-registered externref^M -> externref^N types.
        _ if op == Op::BLOCK || op == Op::LOOP || op == Op::IF => {
            let param_count = chunk.code[*ip];
            let result_count = chunk.code[*ip + 1];
            *ip += 2;
            body.push(op.sub() as u8); // 0x02 / 0x03 / 0x04
            // ⛔ A 1-RESULT BLOCK IS NOT ALWAYS EXTERNREF. This arm used to
            // push `TYPE_EXTERNREF` unconditionally, because the block
            // immediate carries `(param_count, result_count)` — two COUNTS and
            // no type — so there was nothing to choose from. Every truthiness
            // and comparison chain therefore came out declared `externref`
            // while its arms pushed i32, and V8 rejected the consumer:
            // "if[0] expected type i32, found call of type externref".
            // Our VM never caught it because it does not check block result
            // types.
            //
            // `block_i32_results` is the compiler saying which ones are i32,
            // keyed by opcode offset (`*ip` is the operand start, so the
            // opcode began 4 bytes earlier).
            let opcode_offset = ip.saturating_sub(2).saturating_sub(4);
            match (param_count, result_count) {
                (0, 0) => body.push(TYPE_VOID),
                (0, 1) if chunk.block_i32_results.contains(&opcode_offset) => {
                    body.push(TYPE_I32)
                }
                (0, 1) => body.push(TYPE_EXTERNREF),
                key => {
                    let tidx = *type_ctx
                        .block_type_by_results
                        .get(&key)
                        .expect("block functype was not pre-registered");
                    write_leb128_i32(body, tidx as i32);
                }
            }
        }
        // ELSE: no operands.
        _ if op == Op::ELSE => {
            body.push(0x05); // else
        }
        // `i32.wrap_i64` is the ONLY way an i64 leaves the raw world here, and
        // its i32 result goes straight into an externref local. It had no arm,
        // so it fell to the default and stayed raw.
        _ if op == Op::I32_WRAP_I64 => {
            body.push(op.sub() as u8);
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block);
        }
        _ if op == Op::MEMORY_SIZE || op == Op::MEMORY_GROW => {
            // `memory.grow` takes a raw i32 delta and both answer a raw i32.
            if op == Op::MEMORY_GROW {
                emit_unbox_i32(body, rt_idx);
            }
            body.push(op.sub() as u8);
            let memidx = read_u16(&chunk.code, ip);
            write_leb128_u32(body, memidx as u32);
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block);
        }
        // Memory load: the address is the ONLY operand, so it is on top and
        // unboxes in place; the loaded value rejoins the value ABI.
        _ if memory_load_shape(op).is_some() => {
            let (default_align, result) = memory_load_shape(op).unwrap_or((0, 0));
            emit_unbox_i32(body, rt_idx); // externref addr → i32
            body.push(op.sub() as u8);
            let (align, offset, memidx) = read_optional_memarg(chunk, ip, default_align);
            encode_memarg_with_memidx(body, align, offset, memidx);
            match result {
                crate::encoding::TYPE_I32 => {
                    box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block)
                }
                crate::encoding::TYPE_F64 => emit_box_f64(body, rt_idx),
                crate::encoding::TYPE_F32 => {
                    // The reader widens f32 to `Value::F64` and there is no f32
                    // box — the same widening `f32.const` already does here.
                    body.push(0xBB); // f64.promote_f32
                    emit_box_f64(body, rt_idx);
                }
                // i64: raw, and it stays raw. There is no i64 box (see the
                // `i64.const` arm) and its consumers are i64 instructions.
                _ => {}
            }
        }
        // Memory store: [addr, value]. The address is UNDER the value, so the
        // value goes to a temp exactly as `ARRAY_SET` does.
        _ if memory_store_shape(op).is_some() => {
            let (default_align, value) = memory_store_shape(op).unwrap_or((0, 0));
            let liftable = value != crate::encoding::TYPE_I64;
            if liftable {
                body.push(0x21); // local.set $temp (save value)
                write_leb128_u32(body, temp_idx);
                emit_unbox_i32(body, rt_idx); // externref addr → i32
                body.push(0x20); // local.get $temp (restore value)
                write_leb128_u32(body, temp_idx);
                match value {
                    crate::encoding::TYPE_I32 => emit_unbox_i32(body, rt_idx),
                    crate::encoding::TYPE_F64 => emit_unbox_f64(body, rt_idx),
                    crate::encoding::TYPE_F32 => {
                        emit_unbox_f64(body, rt_idx);
                        body.push(0xB6); // f32.demote_f64
                    }
                    _ => {}
                }
            }
            // ⛔ AN i64 STORE IS LEFT ALONE, DELIBERATELY. Its value operand is
            // already a raw i64 (only an i64 instruction can have produced it)
            // and the temps this writer declares are externref, so the value
            // cannot be lifted out of the way to reach the address underneath.
            // Reconciling the address alone would reorder a stack it cannot put
            // back. The wall is the missing i64 unbox, not this arm.
            body.push(op.sub() as u8);
            let (align, offset, memidx) = read_optional_memarg(chunk, ip, default_align);
            encode_memarg_with_memidx(body, align, offset, memidx);
        }
        // WASM global.get/set — the operand IS the global index.
        //
        // It used to be a constant index naming the global, which this resolved
        // through `global_map`. The compiler now assigns a real `globalidx`
        // (`globals::normalize_global_table`) over the SAME ordering this
        // writer uses — `chunk::global_index_space`: the `rt_globals()`
        // js-primitive singletons, then one imported global per string
        // constant, then host globals, then the module's own. One definition,
        // consumed at both ends, so no remapping is needed or wanted here.
        _ if op == Op::GLOBAL_GET => {
            let gidx = read_u16(&chunk.code, ip) as u32;
            body.push(0x23); // global.get
            write_leb128_u32(body, gidx);
        }
        _ if op == Op::GLOBAL_SET => {
            // ⛔ THE VM DOES NOT KEEP IT. This emitted
            // `local.tee $temp; global.set; local.get $temp` on the stated
            // grounds that "global.set consumes it — but our VM keeps it".
            // `dispatch.rs` `GLOBAL_SET` is `let val = self.pop()`: it consumes
            // the value exactly as the spec instruction does. The restore
            // therefore left a value the bytecode had already spent, and the
            // surplus surfaced at the next block boundary as "values remaining
            // on stack at end of block" — never at the global write itself.
            //
            // Same false premise as the `LOCAL_SET` arm above, in a second
            // place: a claim about the VM that the VM does not make.
            let wasm_gidx = read_u16(&chunk.code, ip) as u32;
            body.push(0x24); // global.set
            write_leb128_u32(body, wasm_gidx);
        }
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
            // Internal fixed-width u16 tag immediate → LEB tagidx, mapped into
            // the module's tag index space. Writing the shared exception tag
            // for every throw discarded which tag was raised, so a catch on a
            // different tag matched it.
            let tag = read_u16(&chunk.code, ip);
            body.push(0x08); // throw
            write_leb128_u32(body, tag_plan.module_tag(ci, tag));
        }
        // `throw_ref` re-raises the exception an exnref refers to: opcode 0x0A,
        // and it takes NO tag immediate — the tag comes from the exception
        // itself. This used to emit 0x08 (`throw`) plus a tagidx, which is a
        // DIFFERENT instruction: it left the exnref on the stack and raised a
        // fresh exception under the shared tag.
        _ if op == Op::THROW_REF => {
            body.push(0x0A);
        }
        // Reference-types `table.get tbl` / `table.set tbl` (core prefix).
        // Bytecode carries a u16 BE table index; WASM binary uses a
        // LEB128 tableidx, so we re-serialize on the way out.
        _ if op == Op::TABLE_GET => {
            let tbl = ((chunk.code[*ip] as u32) << 8) | chunk.code[*ip + 1] as u32;
            *ip += 2;
            body.push(0x25);
            write_leb128_u32(body, tbl);
        }
        _ if op == Op::TABLE_SET => {
            let tbl = ((chunk.code[*ip] as u32) << 8) | chunk.code[*ip + 1] as u32;
            *ip += 2;
            body.push(0x26);
            write_leb128_u32(body, tbl);
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
        // `ref.null <heaptype>` — the heaptype is carried as the instruction's
        // own immediate and written straight through, so `ref.null extern` and
        // `ref.null none` round-trip as themselves. This used to hardcode
        // `HT_EXTERN` here, with a separate `0xFF` opcode standing in for the
        // GC-heap case.
        _ if op == Op::NULL => {
            let ht = chunk.code.get(*ip).copied().unwrap_or(HT_EXTERN);
            *ip += 1;
            body.push(0xD0);
            body.push(ht);
        }
        // ref.is_null produces i32 — box it since our value representation is externref
        _ if op == Op::REF_IS_NULL => {
            body.push(0xD1); // ref.is_null → i32
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block);
        }
        // GC proposal (core prefix): `ref.eq` compares two references in the
        // ANY hierarchy and produces i32.
        //
        // ⛔ IT CANNOT SEE AN `externref`. `eqref` bottoms out in ANY; our
        // value ABI is EXTERN. V8: `ref.eq[0] expected either eqref or
        // (ref null shared eq), found local.get of type externref`. Both
        // operands have to be internalised, and the second one sits ON TOP of
        // the first — the same temp dance `ARRAY_SET` uses to reach under.
        _ if op == Op::REF_EQ => {
            body.push(0x21); // local.set $temp (save b)
            write_leb128_u32(body, temp_idx);
            emit_internalize(body); // a: externref → anyref
            emit_ref_cast_eq(body); // anyref → eqref
            body.push(0x20); // local.get $temp (restore b)
            write_leb128_u32(body, temp_idx);
            emit_internalize(body); // b: externref → anyref
            emit_ref_cast_eq(body); // anyref → eqref
            body.push(0xD3);
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block);
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
            emit_binary_f64_op(body, op.sub() as u8, rt_idx, temp_idx);
        }

        // ── f64 comparisons: unbox both → compare → rebox i32 result ──
        // ⛔ `F64_EQ`/`F64_NE` WERE ABSENT and fell to the catch-all, which
        // emits the bare opcode onto two BOXED operands: "expected f64, found
        // externref". Every comparison that unboxes has to be listed here —
        // an omission is not a missing optimisation, it is invalid wasm.
        _ if op == Op::F64_LT
            || op == Op::F64_GT
            || op == Op::F64_LE
            || op == Op::F64_GE
            || op == Op::F64_EQ
            || op == Op::F64_NE =>
        {
            emit_binary_f64_cmp(
                !(is_i32_block_result(chunk, *ip, in_i32_block)
                    || i32_condition_follows(chunk, *ip)),
                body,
                op.sub() as u8,
                rt_idx,
                temp_idx,
            );
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
            emit_binary_i32_op(
                !(is_i32_block_result(chunk, *ip, in_i32_block)
                    || i32_condition_follows(chunk, *ip)),
                body,
                op.sub() as u8,
                rt_idx,
                temp_idx,
            );
        }

        // ── i32 comparisons (eq, ne): unbox both → compare → rebox ──
        _ if op == Op::EQ
            || op == Op::NE
            || op == Op::I32_LT_S
            || op == Op::I32_LT_U
            || op == Op::I32_GT_S
            || op == Op::I32_GT_U
            || op == Op::I32_LE_S
            || op == Op::I32_LE_U
            || op == Op::I32_GE_S
            || op == Op::I32_GE_U =>
        {
            emit_binary_i32_cmp(
                !(is_i32_block_result(chunk, *ip, in_i32_block)
                    || i32_condition_follows(chunk, *ip)),
                body,
                op.sub() as u8,
                rt_idx,
                temp_idx,
            );
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
            body.push(op.sub() as u8);
            emit_box_f64(body, rt_idx);
        }

        // ── i32 unary ops ──
        _ if op == Op::I32_EQZ => {
            emit_unbox_i32(body, rt_idx);
            body.push(op.sub() as u8);
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block);
        }
        _ if op == Op::I32_CLZ || op == Op::I32_CTZ || op == Op::I32_POPCNT => {
            emit_unbox_i32(body, rt_idx);
            body.push(op.sub() as u8);
            emit_box_i32(body, rt_idx);
        }

        // ── Conversions: unbox source type → convert → rebox target type ──
        _ if op == Op::I32_FROM_F64 => {
            // externref → f64 → i32.trunc_f64_s → externref
            emit_unbox_f64(body, rt_idx);
            body.push(op.sub() as u8);
            emit_box_i32(body, rt_idx);
        }
        _ if op == Op::F64_FROM_I32 => {
            // externref → i32 → f64.convert_i32_s → externref
            emit_unbox_i32(body, rt_idx);
            body.push(op.sub() as u8);
            emit_box_f64(body, rt_idx);
        }

        // ref.func (Closure format): emit ref.func with WASM function index
        _ if op == Op::REF_FUNC => {
            let chunk_idx = read_u16(&chunk.code, ip);
            let uv_count = (chunk.code[*ip] & 0x7f) as usize; // mask 0x80 no-intern flag
            *ip += 1;
            *ip += uv_count * 3; // skip upvalue descriptors (u8 is_local + u16 index)
            // Store as table index (i32) for call_indirect — box as externref.
            // chunk_idx is the table index because the element section maps chunks 0..N to table slots.
            body.push(0x41); // i32.const
            write_leb128_i32(body, chunk_idx as i32);
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block);
        }
        _ if op == Op::RETHROW || op == Op::DELEGATE => {
            body.push(op.sub() as u8);
            let depth = read_leb_u32(&chunk.code, ip);
            write_leb128_u32(body, depth);
        }
        // ── Stack-switching proposal: real core opcodes 0xE0..=0xE6 ──
        _ if op == Op::CONT_NEW => {
            body.push(crate::writer::proposals::stack_switching::OP_CONT_NEW);
            write_leb128_u32(body, type_ctx.continuation_type_idx);
        }
        _ if op == Op::SUSPEND => {
            let tag_idx = read_u16(&chunk.code, ip);
            body.push(crate::writer::proposals::stack_switching::OP_SUSPEND);
            write_leb128_u32(body, tag_idx as u32);
        }
        _ if op == Op::RESUME => {
            let _ = read_u16(&chunk.code, ip);
            body.push(crate::writer::proposals::stack_switching::OP_RESUME);
            write_leb128_u32(body, type_ctx.continuation_type_idx);
            emit_stack_switch_handlers(body, chunk, op_start);
        }
        _ if op == Op::SWITCH => {
            let tag_idx = read_u16(&chunk.code, ip);
            body.push(crate::writer::proposals::stack_switching::OP_SWITCH);
            write_leb128_u32(body, type_ctx.continuation_type_idx);
            write_leb128_u32(body, tag_idx as u32);
        }
        _ if op == Op::CONT_BIND => {
            let _argc = chunk.code[*ip];
            *ip += 1;
            body.push(crate::writer::proposals::stack_switching::OP_CONT_BIND);
            write_leb128_u32(body, type_ctx.continuation_type_idx);
            write_leb128_u32(body, type_ctx.continuation_type_idx);
        }
        _ if op == Op::RESUME_THROW => {
            let tag_idx = read_u16(&chunk.code, ip);
            body.push(crate::writer::proposals::stack_switching::OP_RESUME_THROW);
            write_leb128_u32(body, type_ctx.continuation_type_idx);
            write_leb128_u32(body, tag_idx as u32);
            emit_stack_switch_handlers(body, chunk, op_start);
        }
        _ if op == Op::RESUME_THROW_REF => {
            body.push(crate::writer::proposals::stack_switching::OP_RESUME_THROW_REF);
            write_leb128_u32(body, type_ctx.continuation_type_idx);
            emit_stack_switch_handlers(body, chunk, op_start);
        }

        // ── Spec consts ─────────────────────────────────────────────────
        // The VM-internal immediates already use the spec encodings (signed
        // LEB128 for i32/i64, raw LE bytes for f32/f64), so they copy through
        // verbatim. The value is then boxed to the externref stack ABI —
        // exactly what the retired constant-pool CONST arm did (i64 rides
        // box_i32, its precedent). Falling into the generic default here
        // dropped the immediate entirely: `i32.const 42` serialized as a
        // bare 0x41 — malformed WASM.
        _ if op == Op::I32_CONST || op == Op::I64_CONST || op == Op::F64_CONST => {
            let sz = op.operand_format().size_in(&chunk.code, *ip);
            body.push(op.sub() as u8);
            body.extend_from_slice(&chunk.code[*ip..*ip + sz]);
            *ip += sz;
            if op == Op::F64_CONST {
                emit_box_f64(body, rt_idx);
            } else if op == Op::I32_CONST {
                box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block);
            }
            // ⛔ `i64.const` IS NOT BOXED, AND ITS OLD "i64 rides box_i32"
            // PRECEDENT WAS A TYPE ERROR: it emitted
            // `wasm:js-number.fromI32(i64)`. There is no i64 box to use
            // instead — our value ABI is externref and js-primitive-builtins
            // exposes only `wasm:js-bigint.test`, the i64 conversions having
            // been removed from the proposal — and there does not need to be:
            // both producers in the tree feed a real i64 instruction.
            // `canon stream.new`'s packed handle is shifted by `i64.shr_u`
            // (`io.rs:140`) and `memory.atomic.wait32`'s timeout is consumed
            // as the i64 the spec declares (`threading.rs`). Raw is what they
            // want. The remaining half — an i64 STORED INTO A LOCAL, which
            // this writer declares externref — is not fixable from this arm
            // and is not fixed here.

        }
        _ if op == Op::F32_CONST => {
            let sz = op.operand_format().size_in(&chunk.code, *ip);
            body.push(op.sub() as u8);
            body.extend_from_slice(&chunk.code[*ip..*ip + sz]);
            *ip += sz;
            // Widen to f64 before boxing — the reader deliberately widens
            // f32 constants to Value::F64, and there is no f32 box.
            body.push(0xBB); // f64.promote_f32
            emit_box_f64(body, rt_idx);
        }

        _ => {
            // Other core ops: emit WASM byte directly
            body.push(op.sub() as u8);
            *ip += op.operand_format().size_in(&chunk.code, *ip);
        }
    }
}

/// `(natural alignment, result valtype)` for a plain memory LOAD, or `None`
/// when `op` is not one. The alignments are the spec's naturals, unchanged
/// from the arms this replaced.
fn memory_load_shape(op: Op) -> Option<(u32, u8)> {
    use crate::encoding::{TYPE_F32, TYPE_F64, TYPE_I32, TYPE_I64};
    Some(match op {
        _ if op == Op::I32_LOAD => (2, TYPE_I32),
        _ if op == Op::F32_LOAD => (2, TYPE_F32),
        _ if op == Op::I64_LOAD => (3, TYPE_I64),
        _ if op == Op::F64_LOAD => (3, TYPE_F64),
        _ if op == Op::I32_LOAD8_S || op == Op::I32_LOAD8_U => (0, TYPE_I32),
        _ if op == Op::I64_LOAD8_S || op == Op::I64_LOAD8_U => (0, TYPE_I64),
        _ if op == Op::I32_LOAD16_S || op == Op::I32_LOAD16_U => (1, TYPE_I32),
        _ if op == Op::I64_LOAD16_S || op == Op::I64_LOAD16_U => (1, TYPE_I64),
        _ if op == Op::I64_LOAD32_S || op == Op::I64_LOAD32_U => (2, TYPE_I64),
        _ => return None,
    })
}

/// `(natural alignment, value valtype)` for a plain memory STORE, or `None`.
fn memory_store_shape(op: Op) -> Option<(u32, u8)> {
    use crate::encoding::{TYPE_F32, TYPE_F64, TYPE_I32, TYPE_I64};
    Some(match op {
        _ if op == Op::I32_STORE => (2, TYPE_I32),
        _ if op == Op::F32_STORE => (2, TYPE_F32),
        _ if op == Op::I64_STORE => (3, TYPE_I64),
        _ if op == Op::F64_STORE => (3, TYPE_F64),
        _ if op == Op::I32_STORE8 => (0, TYPE_I32),
        _ if op == Op::I64_STORE8 => (0, TYPE_I64),
        _ if op == Op::I32_STORE16 => (1, TYPE_I32),
        _ if op == Op::I64_STORE16 => (1, TYPE_I64),
        _ if op == Op::I64_STORE32 => (2, TYPE_I64),
        _ => return None,
    })
}

/// An atomic whose ONLY operand is the address and whose result is an i32.
///
/// These are the three the whole tree emits with a single operand; every other
/// atomic takes a value under/over the address and needs the `temp_idx` dance
/// `ARRAY_SET` uses. See the ⛔ note on `emit_thread_prefixed_op`.
fn atomic_is_i32_address_only_load(op: Op) -> bool {
    op == Op::I32_ATOMIC_LOAD || op == Op::I32_ATOMIC_LOAD8_U || op == Op::I32_ATOMIC_LOAD16_U
}

/// ⛔ AN ATOMIC IS A REAL WASM INSTRUCTION, NOT A DYNAMIC OPERATION.
///
/// The VM hides this: `pop_atomic_addr` (`threads.rs:263`) does
/// `self.pop().as_i32()`, a COERCION, so a boxed address works there and the
/// corpus never sees the defect. WASM does not coerce — `i32.atomic.load` fed
/// the externref that `ecma:object.get(obj, "__vybe_shared_addr")` returns is
/// `type mismatch: expected i32, found externref`, and every emitted module
/// carrying one is invalid. Same defect class as the `I64_*` ops used as
/// dynamic BigInt ops, and the same fix shape: reconcile at the ABI boundary
/// the writer already owns (`ARRAY_GET` unboxes its index here for exactly
/// this reason).
///
/// ⛔ WHAT IS NOT DONE, AND WHY IT CANNOT BE FROM HERE:
/// `memory.atomic.wait32`/`wait64` and the whole `i64.atomic.*` family take or
/// return an i64 OPERAND. There is no i64 unbox: our value ABI is externref,
/// and `wasm:js-bigint` in js-primitive-builtins exposes ONLY `test` — the
/// i64 conversions were deliberately removed from the proposal. Closing them
/// needs a host conversion declared `(result i64)`, relying on
/// JS-BigInt-integration to do the marshalling — a NEW host function, which is
/// not mine to add. Those arms are left emitting exactly what they emitted
/// before; they are no more wrong than they were, and the boundary is here in
/// writing rather than silent.
fn emit_thread_prefixed_op(
    body: &mut Vec<u8>,
    op: Op,
    chunk: &Chunk,
    ip: &mut usize,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    in_i32_block: bool,
) {
    let address_only_load = atomic_is_i32_address_only_load(op);
    if address_only_load {
        emit_unbox_i32(body, rt_idx); // externref addr → i32
    }
    body.push(0xFE);
    write_leb128_u32(body, op.sub() as u32);
    if op == Op::ATOMIC_FENCE {
        let immediate = chunk.code.get(*ip).copied().unwrap_or(0);
        *ip = (*ip).saturating_add(1).min(chunk.code.len());
        body.push(immediate);
        return;
    }
    // Every 0xFE atomic (notify/wait included) carries an explicit memarg in
    // the internal bytecode — declared MemArg, emitted by every emitter. The
    // old "does the next 4-byte block decode as an opcode?" guess is gone:
    // memarg bytes could legitimately decode as an opcode, silently dropping
    // a real memarg.
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
    if address_only_load {
        // i32 result → back onto the externref value ABI, unless the consumer
        // is a raw-i32 region. `box_i32_unless_condition`, not a bare box: an
        // atomic feeding an `if` directly must stay raw, exactly as for every
        // other i32 producer in this writer.
        box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block);
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
    // The internal SIMD memarg is self-describing (`SimdMemArg`): present iff
    // the first LEB carries the 0x80 marker. Instruction group-hi bytes are
    // always 0x00, so peeking the marker is unambiguous — no opcode-decode
    // guessing needed.
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

/// Allocate a fresh dynamic object, consuming `prop_count` key/value pairs
/// that are already on the stack — the lowering of `struct.new` with typeidx 0.
///
/// Stack in:  `[k1, v1, … kN, vN]`
/// Stack out: `[obj]`
///
/// The object cannot be allocated before the pairs, because they are already
/// below it; and nothing can be stored into it until it exists. So the object
/// is allocated on top and each pair is then unstacked through temps.
///
/// ⛔ PAIRS ARE CONSUMED TOP-DOWN, so keys are inserted in REVERSE source
/// order. That is visible only through key ORDER (`Object.keys`) or a repeated
/// key, and no front end emits this form at all — `emit_class_alloc` is the
/// only producer of typeidx 0 and it always passes `count = 0`, building the
/// object empty and assigning each property with its own `STRUCT_SET`. Fixing
/// the order needs 2N locals rather than 3; that is worth doing the day a
/// producer appears, and not before.
fn emit_dynamic_object_new(
    body: &mut Vec<u8>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    new_idx: usize,
    prop_count: u16,
    temp_idx: u32,
) {
    body.push(0x10); // call ecma:object.new() -> externref
    write_leb128_u32(body, new_idx as u32);
    let Some(&set_idx) = rt_idx.get(&("ecma:object", "set")) else {
        return;
    };
    if prop_count == 0 {
        return;
    }
    let (obj, val, key) = (temp_idx, temp_idx + 1, temp_idx + 2);
    body.push(0x21); // local.set $obj
    write_leb128_u32(body, obj);
    for _ in 0..prop_count {
        body.push(0x21); // local.set $val
        write_leb128_u32(body, val);
        body.push(0x21); // local.set $key
        write_leb128_u32(body, key);
        for slot in [obj, key, val] {
            body.push(0x20); // local.get
            write_leb128_u32(body, slot);
        }
        body.push(0x10); // call ecma:object.set(obj, key, val) -> boolean
        write_leb128_u32(body, set_idx as u32);
        body.push(0x1A); // drop the §10.1.9 success flag
    }
    body.push(0x20); // local.get $obj
    write_leb128_u32(body, obj);
}

/// Does a `ref.test` / `ref.cast` on this heaptype operate in the ANY
/// hierarchy — i.e. does its operand have to be internalized first?
///
/// ⛔ OUR VALUE ABI IS EXTERNREF; THESE INSTRUCTIONS TAKE ANYREF. `struct.get`
/// already internalized before its `ref.cast`, but the standalone `ref.test` /
/// `ref.cast` arms emitted the instruction straight onto an externref operand:
/// "expected anyref, found externref". The `extern`/`func` hierarchies are the
/// exception — an operand already in the right hierarchy must not be converted.
/// A concrete type index (a non-negative LEB, so below the abstract encodings)
/// is always a GC type and always needs the conversion.
fn heaptype_is_internal(ht_bytes: &[u8]) -> bool {
    !matches!(
        ht_bytes.first().copied(),
        Some(HT_EXTERN) | Some(HT_NOEXTERN) | Some(HT_FUNC) | Some(HT_NOFUNC) | None
    )
}

/// True when the NEXT bytecode instruction pops a RAW i32 rather than a boxed
/// value. Spec `if` and `br_if` both take an i32 condition — nothing else in
/// the emitted stream does.
///
/// ⛔ THE VALUE ABI IS NOT UNIFORM. Everything an arm produces is boxed to
/// externref, but the builtin calls (`wasm:js-string.test` and friends) return
/// a bare i32 by their declared signature, and `if`/`br_if` consume that i32
/// directly. So a comparison that boxed its result and was then branched on
/// came out as `i32.lt_s; fromI32; if` — V8: "if[0] expected type i32, found
/// call of type externref". The box was pure waste at exactly the sites where
/// it was also wrong.
///
/// `ip` must point at the next opcode header, which is where every arm leaves
/// it once its own operands are consumed. Bounded like the arity scanner: a
/// producer at the very end of a chunk has no successor to read.
fn i32_condition_follows(chunk: &Chunk, ip: usize) -> bool {
    matches!(next_op(chunk, ip), Some(next) if next == Op::IF || next == Op::BR_IF)
}

/// Decode the instruction at `ip` without consuming it. Bounded like the arity
/// scanner: a producer at the very end of a chunk has no successor to read.
fn next_op(chunk: &Chunk, ip: usize) -> Option<Op> {
    if ip + 3 >= chunk.code.len() {
        return None;
    }
    let g = ((chunk.code[ip] as u16) << 8) | chunk.code[ip + 1] as u16;
    let sub = ((chunk.code[ip + 2] as u16) << 8) | chunk.code[ip + 3] as u16;
    Op::decode(g, sub)
}

/// Is the value about to be emitted the RESULT of an arm of an i32-typed
/// block — the last thing pushed before its `else` or `end`?
///
/// ⛔ "INSIDE AN i32 BLOCK" IS NOT THE SAME QUESTION, and answering that one
/// instead left every i32 producer in the ladder raw — including the operands
/// of `emit_binary_i32_cmp`, which unboxes what it is given and so parked a
/// bare i32 in an externref temp. Only the arm's RESULT is the block's i32;
/// everything else inside is an ordinary boxed value.
fn is_i32_block_result(chunk: &Chunk, ip: usize, in_i32_block: bool) -> bool {
    in_i32_block
        && matches!(next_op(chunk, ip), Some(next) if next == Op::ELSE || next == Op::END)
}

/// Box an i32 to the externref value ABI unless the very next instruction is
/// going to want it raw. See `i32_condition_follows`.
fn box_i32_unless_condition(
    body: &mut Vec<u8>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    chunk: &Chunk,
    ip: usize,
    in_i32_block: bool,
) {
    // ⛔ AN i32-RESULT BLOCK IS A RAW REGION. `emit_if_i32`/`emit_block_i32`
    // are the compiler saying "this block's result is a bare i32, not a
    // value" — that is what the whole truthiness ladder is built from, and
    // every arm of it ends in an i32 producer. Boxing those made each arm
    // yield an externref out of a block declared `(result i32)`.
    if is_i32_block_result(chunk, ip, in_i32_block) || i32_condition_follows(chunk, ip) {
        return;
    }
    emit_box_i32(body, rt_idx);
}

/// Unbox the arguments of a `call` to an import declared with raw numeric
/// params — the operand half of the ABI reconciliation `raw_result_funcs`
/// does for results.
///
/// ⛔ NOTHING ON THIS STACK IS EVER A RAW NUMBER. Every f64 and i32 this
/// writer produces — `f64.const`, `f64.add`, a `-> i32` builtin result — is
/// boxed onto the externref value ABI immediately. So a compiler that pushes
/// "an f64" and calls `wasm:js-string.fromF64` (declared `(param f64)`) is
/// pushing an externref, and V8 rejects it. The compiler is not wrong: raw is
/// not expressible at a call boundary. The unbox belongs here, where the
/// declared signature is known.
///
/// The deepest raw param may sit under other arguments, so the ones above it
/// are spilled to temps and restored — the same reorder `ARRAY_SET` uses.
/// Only ONE spill temp is taken, which covers every signature table the
/// writer consults (`substring` and `fromCharCodeArray` are the deepest, at
/// one). A signature needing more is left untransformed rather than half-done.
fn emit_unbox_raw_params(
    body: &mut Vec<u8>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    type_ctx: &WasmTypeContext,
    funcidx: u32,
    argc: u8,
    temp_idx: u32,
) {
    let Some(params) = type_ctx.raw_param_funcs.get(&funcidx) else {
        return;
    };
    // ⛔ A DECLARED ARITY IS NOT THE CALL SITE'S ARITY. Reordering a stack
    // that holds a different number of arguments than the table describes
    // would spill operands that are not there. A call site that disagrees
    // with its own import's signature is already broken; do not make it
    // worse, and do not hide it behind a partial reorder.
    if argc as usize != params.len() {
        return;
    }
    let Some(deepest_raw) = params.iter().position(|&t| t != crate::encoding::TYPE_EXTERNREF)
    else {
        return;
    };
    let spill = params.len() - 1 - deepest_raw;
    // ⛔ THIS LIMIT MUST COVER EVERY TABLE THE WRITER CONSULTS, or it silently
    // declines to reconcile a call it was written for. `canon stream.read` and
    // `stream.write` are `(i32 i32 i32)` — deepest raw param at index 0, so two
    // arguments sit above it. Two is now the deepest any consulted signature
    // needs; a `return` here is invisible, so raising it is not optional when a
    // table grows.
    if spill > 2 {
        return;
    }
    let unbox = |body: &mut Vec<u8>, vt: u8| match vt {
        crate::encoding::TYPE_I32 => emit_unbox_i32(body, rt_idx),
        crate::encoding::TYPE_F64 => emit_unbox_f64(body, rt_idx),
        _ => {}
    };
    // Spill the arguments sitting above the deepest raw one; temp_idx + i
    // holds param n-1-i.
    for i in 0..spill {
        body.push(0x21); // local.set
        write_leb128_u32(body, temp_idx + i as u32);
    }
    unbox(body, params[deepest_raw]);
    for i in (0..spill).rev() {
        body.push(0x20); // local.get
        write_leb128_u32(body, temp_idx + i as u32);
        unbox(body, params[params.len() - 1 - i]);
    }
}

/// Emit binary f64 comparison: [externref_a, externref_b] → f64.cmp → [externref_result(i32)]
fn emit_binary_f64_cmp(
    box_result: bool,
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
    if box_result {
        emit_box_i32(body, rt_idx); // fromI32 → externref
    }
}

/// Emit binary i32 op: [externref_a, externref_b] → i32.op → [externref_result]
fn emit_binary_i32_op(
    box_result: bool,
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
    if box_result {
        emit_box_i32(body, rt_idx); // fromI32 → externref
    }
}

/// Emit binary i32 comparison: [externref_a, externref_b] → i32.cmp → [externref_result]
fn emit_binary_i32_cmp(
    box_result: bool,
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
    if box_result {
        emit_box_i32(body, rt_idx); // fromI32 → externref
    }
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
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    type_ctx: &WasmTypeContext,
    temp_idx: u32,
    in_i32_block: bool,
) {
    match op {
        _ if op == Op::STRUCT_NEW => {
            // `(typeidx, count)`. A real typeidx names the type directly; the
            // dynamic form (0) still has to guess a wasm struct type from the
            // key/value pair count, as before.
            let chunk_typeidx = read_u16(&chunk.code, ip);
            let prop_count = read_u16(&chunk.code, ip);
            // ⛔ TYPEIDX 0 IS "NO TYPE", NOT "TYPE 0" — the same convention the
            // read and write sides already honour. An untyped allocation is a
            // fresh DYNAMIC object (`emit_class_alloc` emits `(0, 0)` for every
            // one of them), and there is no wasm struct type that describes it.
            //
            // This used to GUESS a struct type by field count. The guess landed
            // on the user's own class, so `{}` came out as
            // `struct.new_desc $C` — an allocation that pops as many operands
            // as `$C` has fields. In a plain `class C { constructor(){…} }` it
            // ate the two operands sitting under it, and the failure surfaced
            // three instructions later on an unrelated call:
            // "expected externref but nothing on stack". The allocation was
            // never the thing that looked wrong.
            if chunk_typeidx == 0
                && let Some(&new_idx) = rt_idx.get(&("ecma:object", "new"))
            {
                emit_dynamic_object_new(body, rt_idx, new_idx, prop_count, temp_idx);
                return;
            }
            let typeidx = if chunk_typeidx != 0 {
                wasm_struct_type_for_chunk_type(chunk, type_ctx, chunk_typeidx)
            } else {
                wasm_struct_type_matching_field_count(chunk, type_ctx, prop_count)
            };
            // Same prohibition as `struct.new_default`: a type carrying a
            // `(descriptor …)` clause cannot be allocated by `struct.new`.
            //
            // ⚠ The DYNAMIC form reaches it too. `wasm_struct_type_matching_field_count`
            // guesses a type from the field count, and that guess can land on a
            // class — which is descriptor-carrying like every other class — so
            // resolving the descriptor from the RESULT rather than from
            // `chunk_typeidx` is what keeps both paths valid. (The guess itself
            // is a separate defect; this only stops it emitting an invalid
            // module on top of a wrong one.)
            //
            // The descriptor is the LAST operand, so it is pushed after the
            // field values the stack already holds.
            match type_ctx.desc_global_for_described(typeidx) {
                Some(global_idx) => {
                    body.push(0x23); // global.get
                    write_leb128_u32(body, global_idx);
                    body.push(0xFB);
                    write_leb128_u32(body, 0x20); // struct.new_desc
                    write_leb128_u32(body, typeidx);
                }
                None => {
                    body.push(0xFB);
                    write_leb128_u32(body, 0x00); // struct.new
                    write_leb128_u32(body, typeidx);
                }
            }
            emit_externalize(body); // (ref $struct) → externref
        }
        _ if op == Op::STRUCT_GET => {
            let named_type = read_u16(&chunk.code, ip);
            let field_name_idx = read_u16(&chunk.code, ip);
            // ⛔ TYPEIDX 0 IS "NO TYPE", NOT "TYPE 0". A dynamic property read
            // by NAME is not a typed struct access, and lowering it to one
            // emitted `struct.get 0 0` — a well-formed instruction addressing a
            // field that does not exist, which is exactly what V8 reports as
            // `invalid field index: 0`. Core wasm has no by-name property
            // access, so this is a host call.
            if named_type == 0
                && let Some(gidx) = dynamic_prop_name_global(chunk, type_ctx, field_name_idx)
                && let Some(&fidx) = rt_idx.get(&("ecma:object", "get"))
            {
                body.push(0x23); // global.get <name string constant>
                write_leb128_u32(body, gidx);
                body.push(0x10); // call ecma:object.get(obj, name)
                write_leb128_u32(body, fidx as u32);
                return;
            }
            let (typeidx, fieldidx) =
                wasm_struct_field_by_typeidx(chunk, type_ctx, named_type, field_name_idx)
                    .unwrap_or_else(|| {
                        wasm_struct_field_for_name(chunk, type_ctx, field_name_idx)
                    });
            emit_internalize(body); // externref → anyref
            emit_ref_cast(body, typeidx); // anyref → (ref $struct)
            body.push(0xFB);
            write_leb128_u32(body, 0x02); // struct.get
            write_leb128_u32(body, typeidx);
            write_leb128_u32(body, fieldidx);
            // Result is externref (field type) — no conversion needed
        }
        _ if op == Op::STRUCT_SET => {
            let named_type = read_u16(&chunk.code, ip);
            let field_name_idx = read_u16(&chunk.code, ip);
            // See the read side above: typeidx 0 is an untyped by-name write.
            // Stack is `[obj, val]` and the host fn wants `(obj, name, val)`,
            // so the value is parked in the temp while the name is pushed.
            if named_type == 0
                && let Some(gidx) = dynamic_prop_name_global(chunk, type_ctx, field_name_idx)
                && let Some(&fidx) = rt_idx.get(&("ecma:object", "set"))
            {
                body.push(0x21); // local.set $temp = val
                write_leb128_u32(body, temp_idx);
                body.push(0x23); // global.get <name>
                write_leb128_u32(body, gidx);
                body.push(0x20); // local.get $temp = val
                write_leb128_u32(body, temp_idx);
                body.push(0x10); // call ecma:object.set(obj, name, val)
                write_leb128_u32(body, fidx as u32);
                // `[[Set]]` answers a Boolean (§10.1.9) and this lowering is a
                // STATEMENT — discard the flag. Declared `-> boolean` now, so
                // omitting this leaves a value the enclosing block never took.
                body.push(0x1A); // drop
                return;
            }
            let (typeidx, fieldidx) =
                wasm_struct_field_by_typeidx(chunk, type_ctx, named_type, field_name_idx)
                    .unwrap_or_else(|| {
                        wasm_struct_field_for_name(chunk, type_ctx, field_name_idx)
                    });
            // Stack: [externref_obj, externref_val]. struct.set expects
            // [(ref $struct), externref_val] and pushes NOTHING — the VM's
            // internal op now has the same spec shape, so no compensation:
            // save the value, cast the object, reload, set.
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
        }
        _ if op == Op::ARRAY_NEW_FIXED => {
            // Spec: `array.new_fixed $t N` (0xFB 0x08), pops N values.
            // Our bytecode now carries both immediates. A 0 type index means a
            // dynamic-language array literal, which maps to the module's
            // generic `(array (mut externref))`; a stamped index names its own
            // array type.
            let typeidx = read_u16(&chunk.code, ip);
            let elem_count = read_u16(&chunk.code, ip);
            body.push(0xFB);
            write_leb128_u32(body, 0x08);
            write_leb128_u32(
                body,
                if typeidx == 0 {
                    type_ctx.array_type_idx
                } else {
                    typeidx as u32
                },
            );
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
            emit_unbox_i32(body, rt_idx);
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
            let wasm_typeidx = wasm_struct_type_for_chunk_type(chunk, type_ctx, typeidx);
            // ⛔ `struct.new_default` CANNOT allocate a type that carries a
            // `(descriptor …)` clause, and `types.rs` stamps one on every
            // class — so this arm emitting 0xFB 0x01 against a class type is
            // an invalid module, not a missing feature. Every class-bearing
            // module we emitted was rejected on exactly this.
            //
            // The descriptor is the LAST operand, and `struct.new_default`
            // takes none, so pushing the singleton immediately before the
            // allocation is the whole sequence.
            match descriptor_global_for_chunk_type(chunk, type_ctx, typeidx) {
                Some(global_idx) => {
                    body.push(0x23); // global.get
                    write_leb128_u32(body, global_idx);
                    body.push(0xFB);
                    write_leb128_u32(body, 0x21); // struct.new_default_desc
                    write_leb128_u32(body, wasm_typeidx);
                }
                None => {
                    body.push(0xFB);
                    write_leb128_u32(body, 0x01); // struct.new_default
                    write_leb128_u32(body, wasm_typeidx);
                }
            }
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
            // Our bytecode carries a field COUNT alongside the typeidx, the
            // same as `struct.new`. The binary format does not — the count is
            // implied by the declared type — so it is consumed and dropped.
            let _count = read_u16(&chunk.code, ip);
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
        // The descriptor-comparing casts:
        //   ref.cast_desc_eq (ref ht)         → 0xFB 0x23 <heaptype>
        //   ref.cast_desc_eq (ref null ht)    → 0xFB 0x24 <heaptype>
        //   br_on_cast_desc_eq $l ht1 ht2     → 0xFB 0x25 castflags labelidx ht ht
        //   br_on_cast_desc_eq_fail ...       → 0xFB 0x26 castflags labelidx ht ht
        //
        // Our bytecode operand for the target type is a constant-pool index
        // for a type NAME, which has no typeidx to map onto — the same
        // situation as `ref.cast`, so we emit the same conservative `any`
        // heaptype it does.
        // `desc.set_proto $classname` — store the class's JS prototype into
        // field 0 of that class's descriptor singleton.
        //
        // VM-internal, expanded here; nothing of it reaches the binary. The
        // immediate names the CLASS rather than a type index because descriptor
        // type indices and descriptor global indices are both derived HERE
        // (`desc_type` / `desc_global`, off the `(2i, 2i+1)` layout) and the
        // compiler has no access to either. A name also survives the
        // per-chunk problem `desc_global_by_index` warns about: the type table
        // lives on `chunks[0]` only, while this walks every chunk, and
        // `desc_type_indices` is on the shared context.
        //
        // The prototype is on the stack and `struct.set` wants the ref BELOW
        // it, so the value is parked in the scratch local the writer already
        // reserves for exactly this kind of reorder (`count_temp_locals`).
        _ if op == Op::DESC_SET_PROTO => {
            let name_idx = read_u16(&chunk.code, ip) as usize;
            let resolved = match chunk.constants.get(name_idx) {
                Some(Value::String(s)) => type_ctx
                    .desc_type(s)
                    .and_then(|t| type_ctx.desc_global(s).map(|g| (t, g))),
                _ => None,
            };
            match resolved {
                Some((desc_type_idx, desc_global_idx)) => {
                    body.push(0x21); // local.set $tmp
                    write_leb128_u32(body, temp_idx);
                    body.push(0x23); // global.get $C.desc
                    write_leb128_u32(body, desc_global_idx);
                    body.push(0x20); // local.get $tmp
                    write_leb128_u32(body, temp_idx);
                    body.push(0xFB);
                    write_leb128_u32(body, 0x05); // struct.set
                    write_leb128_u32(body, desc_type_idx);
                    write_leb128_u32(body, 0); // field 0
                }
                // Unresolvable class name: drop the prototype rather than
                // leaving it on the stack and desynchronising the frame.
                None => body.push(0x1A), // drop
            }
        }
        _ if op == Op::REF_CAST_DESC_EQ || op == Op::REF_CAST_DESC_EQ_NULL => {
            let name_idx = read_u16(&chunk.code, ip) as usize;
            let ht_bytes = resolve_heaptype_from_name(chunk, name_idx, type_ctx);
            body.push(0xFB);
            write_leb128_u32(body, op.sub() as u32);
            body.extend_from_slice(&ht_bytes);
        }
        _ if op == Op::BR_ON_CAST_DESC_EQ || op == Op::BR_ON_CAST_DESC_EQ_FAIL => {
            // Spec: `0xFB 37|38 (null_1?, null_2?):castflags l:labelidx ht_1 ht_2`.
            //
            // ⛔ This used to emit castflags hardcoded `0x00` and `HT_ANY` for
            // `ht_1`, because the bytecode carried only ONE name — so the
            // emitted instruction did not say what the source said, and a
            // nullable operand or a narrower source type was silently lost.
            // Logged once as "a pre-existing GC MVP gap, not a Custom
            // Descriptors one"; it is a Custom Descriptors gap, both flags and
            // both heaptypes being this instruction's own immediates.
            let to_idx = read_u16(&chunk.code, ip) as usize;
            let from_idx = read_u16(&chunk.code, ip) as usize;
            let depth = chunk.code[*ip];
            *ip += 1;
            let (to_bytes, to_null) = resolve_heaptype_nullable(chunk, to_idx, type_ctx);
            let (from_bytes, from_null) = resolve_heaptype_nullable(chunk, from_idx, type_ctx);
            body.push(0xFB);
            write_leb128_u32(body, op.sub() as u32);
            // bit 0 = ht_1 (source) nullable, bit 1 = ht_2 (target) nullable
            body.push((u8::from(from_null)) | (u8::from(to_null) << 1));
            write_leb128_u32(body, depth as u32);
            body.extend_from_slice(&from_bytes); // ht_1: source
            body.extend_from_slice(&to_bytes); // ht_2: target
        }
        _ if op == Op::STRUCT_GET_S || op == Op::STRUCT_GET_U => {
            // Our struct.get uses a field-name-constant u16 operand;
            // spec packed variants take typeidx + fieldidx. Emit the
            // spec byte with conservative indices for round-trip sanity.
            let _typeidx = read_u16(&chunk.code, ip);
            let field_name_idx = read_u16(&chunk.code, ip);
            let (typeidx, fieldidx) = wasm_struct_field_for_name(chunk, type_ctx, field_name_idx);
            emit_internalize(body);
            body.push(0xFB);
            write_leb128_u32(body, op.sub() as u32);
            write_leb128_u32(body, typeidx);
            write_leb128_u32(body, fieldidx);
        }
        _ if op == Op::REF_TEST_NULL => {
            // `ref.test (ref null ht)`: `0xFB 0x15 <heaptype>`. The bytecode
            // immediate IS a heaptype now, so it passes straight through —
            // only a concrete index needs translating into this module's
            // numbering.
            let ht_bytes = read_heaptype_operand(chunk, ip, type_ctx);
            if heaptype_is_internal(&ht_bytes) {
                emit_internalize(body); // externref → anyref
            }
            body.push(0xFB);
            write_leb128_u32(body, 0x15);
            body.extend_from_slice(&ht_bytes);
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block);
        }
        _ if op == Op::REF_CAST_NULL => {
            let ht_bytes = read_heaptype_operand(chunk, ip, type_ctx);
            let internal = heaptype_is_internal(&ht_bytes);
            if internal {
                emit_internalize(body); // externref → anyref
            }
            body.push(0xFB);
            write_leb128_u32(body, 0x17);
            body.extend_from_slice(&ht_bytes);
            if internal {
                emit_externalize(body); // the cast result rejoins the value ABI
            }
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
            emit_unbox_i32(body, rt_idx); // externref_idx → i32
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
            emit_unbox_i32(body, rt_idx); // externref_idx → i32
            body.push(0x20);
            write_leb128_u32(body, temp_idx); // local.get $temp (restore val)
            // Stack: [(ref null $arr), i32, externref]
            body.push(0xFB);
            write_leb128_u32(body, 0x0E); // array.set
            write_leb128_u32(body, type_ctx.array_type_idx);
            // Spec `array.set` is void — and so is the VM's op now; no dummy.
        }
        _ if op == Op::ARRAY_LENGTH => {
            // Stack: [externref_arr] → need (ref null array)
            emit_internalize(body);
            emit_ref_cast_array(body, type_ctx.array_type_idx);
            body.push(0xFB);
            write_leb128_u32(body, 0x0F); // array.len → i32
            // Result is i32, box to externref
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block);
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
        // ⛔⛔ EXACT CASTS ADD NO OPCODE. Custom Descriptors, Overview
        // "Instructions": *"All existing instructions that take heap type
        // immediates work without modification with the encoding of exact heap
        // types."* Exactness rides in the HEAPTYPE — `0x62 x` — not in a new
        // instruction. `REF_TEST_EXACT` … `REF_CAST_EXACT_NULL` (0xFB 0x27-0x2A)
        // are VM-INTERNAL numbers with no spec counterpart, so they must be
        // rewritten here to the ordinary `ref.test` / `ref.cast` opcode plus an
        // exact heaptype.
        //
        // They had no arm at all and fell to `emit_gc_op`'s catch-all, which
        // emits `0xFB <sub>` and DISCARDS the operand — so they reached the
        // binary as an opcode the spec does not define, carrying no heaptype,
        // desyncing everything after. Invisible to the descriptor suite:
        // `exact.wast` and `exact-casts.wast` pass 11/11 because the VM runs
        // Chunks and never the emitted module. Same shape as the `local.tee`
        // hole, one layer down.
        _ if op == Op::REF_TEST_EXACT
            || op == Op::REF_TEST_EXACT_NULL
            || op == Op::REF_CAST_EXACT
            || op == Op::REF_CAST_EXACT_NULL =>
        {
            let spec_sub: u32 = if op == Op::REF_TEST_EXACT {
                0x14 // ref.test (ref ht)
            } else if op == Op::REF_TEST_EXACT_NULL {
                0x15 // ref.test (ref null ht)
            } else if op == Op::REF_CAST_EXACT {
                0x16 // ref.cast (ref ht)
            } else {
                0x17 // ref.cast (ref null ht)
            };
            let ht_bytes = read_exact_heaptype_operand(chunk, ip, type_ctx);
            let internal = heaptype_is_internal(&ht_bytes);
            if internal {
                emit_internalize(body); // externref → anyref
            }
            body.push(0xFB);
            write_leb128_u32(body, spec_sub);
            body.extend_from_slice(&ht_bytes);
            let is_test = op == Op::REF_TEST_EXACT || op == Op::REF_TEST_EXACT_NULL;
            if is_test {
                box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block);
            } else if internal {
                emit_externalize(body); // the cast result rejoins the value ABI
            }
        }
        _ if op == Op::REF_TEST || op == Op::REF_CAST => {
            // `ref.test` / `ref.cast` (non-null): `0xFB {0x14|0x16} <heaptype>`.
            let ht_bytes = read_heaptype_operand(chunk, ip, type_ctx);
            let internal = heaptype_is_internal(&ht_bytes);
            if internal {
                emit_internalize(body); // externref → anyref
            }
            body.push(0xFB);
            write_leb128_u32(body, op.sub() as u32);
            body.extend_from_slice(&ht_bytes);
            if op == Op::REF_TEST {
                box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block);
            } else if internal {
                emit_externalize(body); // the cast result rejoins the value ABI
            }
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
            emit_unbox_i32(body, rt_idx);
            body.push(0xFB);
            write_leb128_u32(body, 0x1C); // ref.i31
            emit_externalize(body); // (ref i31) → externref
        }
        _ if op == Op::I31_GET_S || op == Op::I31_GET_U => {
            emit_internalize(body); // externref → anyref
            body.push(0xFB);
            write_leb128_u32(body, op.sub() as u32);
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block);
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
    box_result: bool,
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
    if box_result {
        emit_box_i32(body, rt_idx); // fromI32 → externref
    }
}

/// Emit a string constant from the chunk's constant pool.
/// Builds the string char by char using wasm:js-string fromCharCode + concat.
/// Resolve the heaptype bytes for a `ref.test / ref.cast / br_on_cast`
/// operand. The compiler stores a constant-pool index pointing at a
/// string type name. We look up the name in `type_ctx.struct_type_indices`
/// and emit `(signed LEB128) typeidx` when found; otherwise fall back to
/// the abstract `anyref` single-byte heaptype so the binary still validates.
/// Resolve a type NAME held in the constant pool to a heaptype.
///
/// Still used by `br_on_cast` and the Custom Descriptors casts, whose operands
/// have not been converted to heaptype immediates yet — they carry a label
/// depth alongside the type, so their operand format changes with them.
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

/// Read a bytecode heaptype immediate and re-encode it for the binary.
///
/// The bytecode already carries the spec's signed-LEB heaptype, so an abstract
/// type passes through byte-for-byte. A concrete one is an index into the
/// MODULE's type space and has to be translated into the writer's own
/// numbering (described/descriptor pairs), which `struct_type_by_index` does
/// without going near a name. An index with no emitted type — a class declared
/// only so a `ref.test` could name it, and never defined here — widens to
/// `anyref` so the module still validates; the VM's declared-name fallback is
/// what answers that test until the platforms allocate typed objects.
/// The same heaptype operand as [`read_heaptype_operand`], encoded EXACT.
///
/// `heaptype ::= … | 0x62 x:u32 => exact x`.
///
/// ⛔ An ABSTRACT heap type is returned unprefixed. The proposal
/// (Overview, "Exact heap types") states that the encoding *"intentionally
/// makes it impossible to encode an exact abstract heap type"* — `0x62` takes
/// a typeidx, so prefixing an abstract byte would produce a typeidx out of a
/// tag. Only a concrete type can be exact.
fn read_exact_heaptype_operand(
    chunk: &Chunk,
    ip: &mut usize,
    type_ctx: &WasmTypeContext,
) -> Vec<u8> {
    let (value, len) = read_leb128_i32(&chunk.code[*ip..]);
    *ip += len;
    let mut buf = Vec::new();
    match vybe_runtime::opcode::heaptype::HeapType::from_sleb(value) {
        vybe_runtime::opcode::heaptype::HeapType::Abstract(byte) => buf.push(byte),
        vybe_runtime::opcode::heaptype::HeapType::Concrete(index) => {
            match type_ctx.struct_type_by_index(index) {
                Some(wasm_idx) => {
                    buf.push(0x62); // exact
                    write_leb128_u32(&mut buf, wasm_idx);
                }
                None => buf.push(HT_ANY),
            }
        }
    }
    buf
}

/// A name-constant heaptype plus the nullability its SPELLING declared.
///
/// The bytecode carries reftypes as names (`$t`, `ref null $t`, `any`), which
/// is what lets a two-heaptype instruction fit an existing operand format —
/// see the `br_on_cast_desc_eq` row in `opcode/gc.rs`. `castflags` needs the
/// nullability separately from the heaptype bytes, so it is returned rather
/// than folded in.
/// Split a REFTYPE spelling into its heap type bytes and its nullability.
///
/// ⛔ THE NAME HAS TO BE PEELED BEFORE IT IS RESOLVED. This read the
/// nullability off the spelling and then handed the WHOLE spelling to
/// `resolve_heaptype_from_name`, which looks the string up as a type name —
/// so `"ref null $a"` missed the type table and widened to `anyref`, losing
/// exactly the concrete type the immediate exists to name. Latent only
/// because nothing emitted the nullable spelling yet; the first producer
/// would have inherited it silently, since a too-wide heap type still
/// validates and still passes every cast the narrow one passes.
///
/// Accepts `(ref null $a)`, `ref $a`, and the §2.3.4 abbreviations
/// (`anyref`, `eqref`, …), which are shorthand for a NULLABLE reference —
/// `funcref ≡ (ref null func)`.
fn resolve_heaptype_nullable(
    chunk: &Chunk,
    name_idx: usize,
    type_ctx: &WasmTypeContext,
) -> (Vec<u8>, bool) {
    use vybe_runtime::opcode::heaptype::HeapType;
    let raw = match chunk.constants.get(name_idx) {
        Some(v) => v.to_string(),
        None => String::new(),
    };
    let spelled = raw
        .trim()
        .trim_start_matches('(')
        .trim_end_matches(')')
        .trim();
    // `ref null ht` → nullable; `ref ht` → not.
    //
    // ⛔ A BARE NAME IS NOT AUTOMATICALLY NULLABLE. Only the §2.3.4
    // ABBREVIATIONS are nullable shorthand (`funcref ≡ (ref null func)`); a
    // bare user type name is what the instructions that encode nullability in
    // the OPCODE store — `ref.cast_desc_eq` and friends — and it says nothing
    // about nullability at all. Treating every bare name as nullable set the
    // castflag on those, which `gc_test`'s encoding assertion caught.
    let (bare, nullable) = match spelled.strip_prefix("ref ") {
        Some(rest) => match rest.trim().strip_prefix("null ") {
            Some(ht) => (ht.trim(), true),
            None => (rest.trim(), false),
        },
        None => (
            spelled,
            HeapType::from_spec_reftype_name(spelled).is_some(),
        ),
    };
    // A user type wins over an abstract spelling: a module may legitimately
    // declare a type called `any`, and the type table is the authority on
    // what this module actually emitted.
    if let Some(idx) = type_ctx.struct_type(bare) {
        let mut buf = Vec::new();
        write_leb128_i32(&mut buf, idx as i32);
        return (buf, nullable);
    }
    let abstract_ht = HeapType::from_spec_reftype_name(bare)
        .and_then(|(heap, _)| HeapType::from_spec_name(heap))
        .or_else(|| HeapType::from_spec_name(bare));
    match abstract_ht {
        Some(HeapType::Abstract(byte)) => (vec![byte], nullable),
        _ => (vec![HT_ANY], nullable),
    }
}

fn read_heaptype_operand(chunk: &Chunk, ip: &mut usize, type_ctx: &WasmTypeContext) -> Vec<u8> {
    let (value, len) = read_leb128_i32(&chunk.code[*ip..]);
    *ip += len;
    let mut buf = Vec::new();
    match vybe_runtime::opcode::heaptype::HeapType::from_sleb(value) {
        vybe_runtime::opcode::heaptype::HeapType::Abstract(byte) => buf.push(byte),
        vybe_runtime::opcode::heaptype::HeapType::Concrete(index) => {
            match type_ctx.struct_type_by_index(index) {
                Some(wasm_idx) => write_leb128_i32(&mut buf, wasm_idx as i32),
                None => buf.push(HT_ANY),
            }
        }
    }
    buf
}

/// Emit `ref.cast null eq` — narrows `anyref` to `eqref`.
///
/// ⚠ `any.convert_extern` yields `anyref`, which is a SUPERtype of `eqref`, so
/// internalising alone does not satisfy `ref.eq`. Everything our value ABI puts
/// on the ANY side is a struct, array, i31 or null — all of them eq — so the
/// cast is total for values this writer produces. A non-eq reference reaching
/// it would trap, which is the spec's answer for a reference that cannot be
/// compared rather than a silent `false`.
fn emit_ref_cast_eq(body: &mut Vec<u8>) {
    body.push(0xFB);
    write_leb128_u32(body, 0x17); // ref.cast null
    body.push(HT_EQ);
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

/// Emit an op with no dedicated core-arm lowering. Prefix 0xFF holds ZERO
/// opcodes now (CONST, CALL_IMPORT and HALT all retired to spec encodings),
/// so this is only the TRY_TABLE nop-skip and the unknown-op fallback.
fn emit_vm_internal_op(body: &mut Vec<u8>, op: Op, chunk: &Chunk, ip: &mut usize) {
    match op {
        // Exception handling — a TRY_TABLE not recognised as a structural
        // single-clause region (multi-clause try_tables, e.g. from wast, are
        // not yet serialized here) is skipped as a nop. Skip the full variable
        // immediate (1 + 5·clause_count) so multi-clause never mis-parses into
        // malformed bytes. Unreached for current compiler output (the OO path
        // and wast-via-VM both stay single-clause).
        _ if op == Op::TRY_TABLE => {
            let clause_count = chunk.code[*ip] as usize;
            *ip += 1 + 5 * clause_count; // clause_count + N·[kind,tag(2),offset(2)]
            body.push(0x01);
        }
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
                    // Route through the canonical size (`OperandFormat::size_in`,
                    // `opcode/mod.rs`) rather than re-deriving it. This said
                    // `1 + count * 3`, but a clause is FIVE bytes — kind(1) +
                    // tag(2) + label(2) — so it under-skipped by 2 per clause
                    // and resumed decoding mid-operand.
                    *ip += fmt.size_in(&chunk.code, *ip);
                }
                _ => {
                    *ip += fmt.size_in(&chunk.code, *ip);
                }
            }
            body.push(0x01); // nop
        }
    }
}

/// Total instruction size: 4-byte opcode + operand bytes.
pub fn opcode_size(op: Op, code: &[u8], ip: usize) -> usize {
    let base = 4;
    base + op.operand_format().size_in(code, ip + base)
}
