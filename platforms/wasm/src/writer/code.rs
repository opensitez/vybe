//! WASM code section encoding.
//! Translates internal bytecode to WASM binary format.
//!
//! Type strategy: externref is the universal value representation.
//! Ordinary values use externref; explicitly typed i64 temporaries stay native.
//! Typed WASM ops (f64.add, i32.mul, etc.) require unboxing via
//! wasm:js-number builtins (toF64/toI32) before the op and reboxing
//! (fromF64/fromI32) after. Binary ops need a temp externref local
//! to save TOS while unboxing the second operand.

use crate::encoding::*;
use crate::writer::sections::HostImport;
use crate::writer::sections::{
    JS_GLOBAL_UNDEFINED, emit_box_f64, emit_box_i32, emit_import_call,
    emit_unbox_f64, emit_unbox_i32,
};
use crate::writer::types::WasmTypeContext;
use vybe_runtime::opcode::{OperandFormat, read_leb_u32};
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op};

#[path = "source_i32_regions.rs"]
mod source_i32_regions;
use source_i32_regions::SourceI32Regions;

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
fn wasm_param_count(chunk: &Chunk) -> u32 {
    chunk
        .func_sig
        .as_ref()
        .map_or(chunk.arity as u32, |(params, _)| {
            params.split(',').filter(|ty| !ty.trim().is_empty()).count() as u32
        })
}

fn boxed_numeric_param_type(
    chunk: &Chunk,
    slot: u16,
    forwarded_types: &[Option<u8>],
) -> Option<u8> {
    if forwarded_types
        .get(slot as usize)
        .copied()
        .flatten()
        .is_some()
    {
        return None;
    }
    chunk
        .func_sig
        .as_ref()?
        .0
        .split(',')
        .filter(|ty| !ty.trim().is_empty())
        .nth(slot as usize)
        .and_then(|ty| match ty.trim() {
            "i32" => Some(TYPE_I32),
            "f32" => Some(TYPE_F32),
            "f64" => Some(TYPE_F64),
            _ => None,
        })
}

fn emit_box_f32(body: &mut Vec<u8>, rt_idx: &std::collections::HashMap<(&str, &str), usize>) {
    body.push(0xBB); // f64.promote_f32
    emit_box_f64(body, rt_idx);
}

fn emit_unbox_f32(body: &mut Vec<u8>, rt_idx: &std::collections::HashMap<(&str, &str), usize>) {
    emit_unbox_f64(body, rt_idx);
    body.push(0xB6); // f32.demote_f64
}

// A straight-line numeric forwarder can retain the core signature's raw
// value throughout compiler-introduced local copies. Do not infer through
// other instructions: those use the boxed value ABI and need an explicit
// typed boundary rather than a guessed local type.
fn forwarded_numeric_local_types(chunk: &Chunk) -> Vec<Option<u8>> {
    let Some((params, result)) = &chunk.func_sig else {
        return Vec::new();
    };
    let numeric_type = |name: &str| match name.trim() {
        "i32" => Some(TYPE_I32),
        "i64" => Some(TYPE_I64),
        "f32" => Some(TYPE_F32),
        "f64" => Some(TYPE_F64),
        _ => None,
    };
    let Some(result_type) = numeric_type(result) else {
        return Vec::new();
    };
    let mut types = vec![None; chunk.local_count as usize];
    for (slot, name) in params
        .split(',')
        .filter(|name| !name.trim().is_empty())
        .enumerate()
    {
        let Some(entry) = types.get_mut(slot) else {
            return Vec::new();
        };
        *entry = numeric_type(name);
    }
    let mut ip = 0;
    let mut pending_get = None;
    let mut saw_return = false;
    while ip < chunk.code.len() {
        if ip + 4 > chunk.code.len() {
            return Vec::new();
        }
        let Some(op) = Op::decode(
            u16::from_be_bytes([chunk.code[ip], chunk.code[ip + 1]]),
            u16::from_be_bytes([chunk.code[ip + 2], chunk.code[ip + 3]]),
        ) else {
            return Vec::new();
        };
        match op {
            Op::LOCAL_GET if pending_get.is_none() => {
                pending_get = Some(read_u16_at(&chunk.code, ip + 4) as usize);
            }
            Op::LOCAL_SET => {
                let Some(source) = pending_get.take() else {
                    return Vec::new();
                };
                let target = read_u16_at(&chunk.code, ip + 4) as usize;
                let Some(source_type) = types.get(source).copied().flatten() else {
                    return Vec::new();
                };
                let Some(target_type) = types.get_mut(target) else {
                    return Vec::new();
                };
                if target_type.is_some_and(|existing| existing != source_type) {
                    return Vec::new();
                }
                *target_type = Some(source_type);
            }
            Op::RETURN => {
                let Some(source) = pending_get.take() else {
                    return Vec::new();
                };
                if types.get(source).copied().flatten() != Some(result_type) {
                    return Vec::new();
                }
                saw_return = true;
            }
            Op::NULL
                if saw_return
                    && pending_get.is_none()
                    && chunk.code.get(ip + 4).copied() == Some(HT_EXTERN)
                    && next_op(chunk, ip + 5) == Some(Op::RETURN)
                    && ip + 9 == chunk.code.len() =>
            {
                break;
            }
            _ => return Vec::new(),
        }
        ip += opcode_size(op, &chunk.code, ip);
    }
    if saw_return && pending_get.is_none() {
        types
    } else {
        Vec::new()
    }
}

// Core scalar signatures are shared by raw-body verification and emission.
// These are standard Wasm operand/result types, not source-language guesses.
fn raw_scalar_shape(op: Op) -> Option<(u8, usize, u8)> {
    if op.group() == 0xFC {
        return Some(match op.sub() {
            0 | 1 => (TYPE_F32, 1, TYPE_I32),
            2 | 3 => (TYPE_F64, 1, TYPE_I32),
            4 | 5 => (TYPE_F32, 1, TYPE_I64),
            6 | 7 => (TYPE_F64, 1, TYPE_I64),
            _ => return None,
        });
    }
    if op.group() != 0 { return None; }
    Some(match op.sub() {
        0x45 => (TYPE_I32, 1, TYPE_I32),
        0x46..=0x4F => (TYPE_I32, 2, TYPE_I32),
        0x50 => (TYPE_I64, 1, TYPE_I32),
        0x51..=0x5A => (TYPE_I64, 2, TYPE_I32),
        0x5B..=0x60 => (TYPE_F32, 2, TYPE_I32),
        0x61..=0x66 => (TYPE_F64, 2, TYPE_I32),
        0x67..=0x69 => (TYPE_I32, 1, TYPE_I32),
        0x6A..=0x78 => (TYPE_I32, 2, TYPE_I32),
        0x79..=0x7B => (TYPE_I64, 1, TYPE_I64),
        0x7C..=0x8A => (TYPE_I64, 2, TYPE_I64),
        0x8B..=0x91 => (TYPE_F32, 1, TYPE_F32),
        0x92..=0x98 => (TYPE_F32, 2, TYPE_F32),
        0x99..=0x9F => (TYPE_F64, 1, TYPE_F64),
        0xA0..=0xA6 => (TYPE_F64, 2, TYPE_F64),
        0xA7 => (TYPE_I64, 1, TYPE_I32),
        0xA8 | 0xA9 => (TYPE_F32, 1, TYPE_I32),
        0xAA | 0xAB => (TYPE_F64, 1, TYPE_I32),
        0xAC | 0xAD => (TYPE_I32, 1, TYPE_I64),
        0xAE | 0xAF => (TYPE_F32, 1, TYPE_I64),
        0xB0 | 0xB1 => (TYPE_F64, 1, TYPE_I64),
        0xB2 | 0xB3 => (TYPE_I32, 1, TYPE_F32),
        0xB4 | 0xB5 => (TYPE_I64, 1, TYPE_F32),
        0xB6 => (TYPE_F64, 1, TYPE_F32),
        0xB7 | 0xB8 => (TYPE_I32, 1, TYPE_F64),
        0xB9 | 0xBA => (TYPE_I64, 1, TYPE_F64),
        0xBB => (TYPE_F32, 1, TYPE_F64),
        0xBC => (TYPE_F32, 1, TYPE_I32),
        0xBD => (TYPE_F64, 1, TYPE_I64),
        0xBE => (TYPE_I32, 1, TYPE_F32),
        0xBF => (TYPE_I64, 1, TYPE_F64),
        0xC0 | 0xC1 => (TYPE_I32, 1, TYPE_I32),
        0xC2..=0xC4 => (TYPE_I64, 1, TYPE_I64),
        _ => return None,
    })
}

struct RawNumericBody {
    local_types: Vec<Option<u8>>,
}

#[cfg(test)]
#[path = "raw_memory_tests.rs"]
mod raw_memory_tests;

#[cfg(test)]
#[path = "simd_memory_tests.rs"]
mod simd_memory_tests;

// Standard SIMD stack types. Memory addressing is handled separately through
// the shared memarg decoder, not through a special-case byte sequence.
fn raw_simd_shape(op: Op) -> Option<(&'static [u8], u8)> {
    if op.group() != 0xfd { return None; }
    Some(match op.sub() {
        0x0f..=0x11 => (&[TYPE_I32], 0x7b),
        0x12 => (&[TYPE_I64], 0x7b),
        0x13 => (&[TYPE_F32], 0x7b),
        0x14 => (&[TYPE_F64], 0x7b),
        0x15 | 0x16 | 0x18 | 0x19 | 0x1b => (&[0x7b], TYPE_I32),
        0x1d => (&[0x7b], TYPE_I64),
        0x1f => (&[0x7b], TYPE_F32),
        0x21 => (&[0x7b], TYPE_F64),
        0x17 | 0x1a | 0x1c => (&[0x7b, TYPE_I32], 0x7b),
        0x1e => (&[0x7b, TYPE_I64], 0x7b),
        0x20 => (&[0x7b, TYPE_F32], 0x7b),
        0x22 => (&[0x7b, TYPE_F64], 0x7b),
        0x53 | 0x63 | 0x64 | 0x83 | 0x84 | 0xa3 | 0xa4 | 0xc3 | 0xc4 => (&[0x7b], TYPE_I32),
        0x6b..=0x6d | 0x8b..=0x8d | 0xab..=0xad | 0xcb..=0xcd => (&[0x7b, TYPE_I32], 0x7b),
        0x52 => (&[0x7b, 0x7b, 0x7b], 0x7b),
        0x0d | 0x0e | 0x23..=0x4c | 0x4e..=0x51 | 0x65..=0x66
        | 0x6e..=0x73 | 0x76..=0x79 | 0x7b | 0x82 | 0x85..=0x86
        | 0x8e..=0x93 | 0x95..=0x99 | 0x9b..=0x9f
        | 0xae | 0xb1 | 0xb5..=0xba | 0xbc..=0xbf
        | 0xce | 0xd1 | 0xd5..=0xdf | 0xe4..=0xeb | 0xf0..=0xf7 => (&[0x7b, 0x7b], 0x7b),
        0x4d | 0x5e..=0x62 | 0x67..=0x6a | 0x74..=0x75 | 0x7a | 0x7c..=0x81
        | 0x87..=0x8a | 0x94 | 0xa0..=0xa1 | 0xa7..=0xaa
        | 0xc0..=0xc1 | 0xc7..=0xca | 0xe0..=0xe1 | 0xe3 | 0xec..=0xed | 0xef | 0xf8..=0xff => (&[0x7b], 0x7b),
        _ => return None,
    })
}

// In a straight-line core numeric body, the declared signature and every
// producer establish the Wasm types. Retain raw values through local copies
// and arithmetic, without guessing a type from an f64 constant.
fn raw_numeric_body(chunk: &Chunk, type_ctx: &WasmTypeContext) -> Option<RawNumericBody> {
    let (params, result) = chunk.func_sig.as_ref()?;
    let result_types: Vec<_> = result
        .split(',')
        .filter(|ty| !ty.trim().is_empty())
        .map(|ty| match ty.trim() {
            "i32" => Some(TYPE_I32),
            "i64" => Some(TYPE_I64),
            "f32" => Some(TYPE_F32),
            "f64" => Some(TYPE_F64),
            "v128" => Some(0x7b),
            _ => None,
        })
        .collect::<Option<_>>()?;
    let params: Vec<_> = params
        .split(',')
        .filter(|ty| !ty.trim().is_empty())
        .collect();
    let mut local_types = vec![None; chunk.local_count as usize];
    for (slot, param) in params.iter().enumerate() {
        *local_types.get_mut(slot)? = Some(match param.trim() {
            "i32" => TYPE_I32,
            "i64" => TYPE_I64,
            "f32" => TYPE_F32,
            "f64" => TYPE_F64,
            "v128" => 0x7b,
            _ => return None,
        });
    }
    let mut stack: Vec<u8> = Vec::new();
    let mut saw_return = false;
    let mut ip = 0;
    while ip < chunk.code.len() {
        let op = Op::decode(
            u16::from_be_bytes([*chunk.code.get(ip)?, *chunk.code.get(ip + 1)?]),
            u16::from_be_bytes([*chunk.code.get(ip + 2)?, *chunk.code.get(ip + 3)?]),
        )?;
        if let Some((input, arity, result)) = raw_scalar_shape(op) {
            for _ in 0..arity {
                if stack.pop()? != input { return None; }
            }
            stack.push(result);
            ip += opcode_size(op, &chunk.code, ip);
            continue;
        }
        if let Some((inputs, result)) = raw_simd_shape(op) {
            for &input in inputs.iter().rev() {
                if stack.pop()? != input { return None; }
            }
            stack.push(result);
            ip += opcode_size(op, &chunk.code, ip);
            continue;
        }
        match op {
            Op::LOCAL_GET => {
                let slot = read_u16_at(&chunk.code, ip + 4) as usize;
                stack.push(local_types.get(slot).copied().flatten()?);
            }
            Op::LOCAL_SET | Op::LOCAL_TEE => {
                let slot = read_u16_at(&chunk.code, ip + 4) as usize;
                let value_type = stack.pop()?;
                let ty = local_types.get_mut(slot)?;
                if ty.is_some_and(|existing| existing != value_type) {
                    return None;
                }
                *ty = Some(value_type);
                if op == Op::LOCAL_TEE { stack.push(value_type); }
            }
            op if memory_load_shape(op).is_some() =>
            {
                let (alignment, result) = memory_load_shape(op)?;
                let mut operand = ip + 4;
                let (_, _, memory) = read_optional_memarg(chunk, &mut operand, alignment);
                let address = if chunk.memory_is_64.get(memory as usize).copied().unwrap_or(false) {
                    TYPE_I64
                } else { TYPE_I32 };
                if stack.pop()? != address { return None; }
                stack.push(result);
            }
            op if memory_store_shape(op).is_some() =>
            {
                let (alignment, value) = memory_store_shape(op)?;
                let mut operand = ip + 4;
                let (_, _, memory) = read_optional_memarg(chunk, &mut operand, alignment);
                let address = if chunk.memory_is_64.get(memory as usize).copied().unwrap_or(false) {
                    TYPE_I64
                } else { TYPE_I32 };
                if stack.pop()? != value || stack.pop()? != address { return None; }
            }
            Op::MEMORY_SIZE | Op::MEMORY_GROW => {
                let mut operand = ip + 4;
                let memory = read_leb_u32(&chunk.code, &mut operand) as usize;
                let address = if chunk.memory_is_64.get(memory).copied().unwrap_or(false) {
                    TYPE_I64
                } else { TYPE_I32 };
                if op == Op::MEMORY_GROW && stack.pop()? != address { return None; }
                stack.push(address);
            }
            Op::MEMORY_COPY | Op::MEMORY_FILL | Op::MEMORY_INIT | Op::TABLE_COPY | Op::TABLE_INIT => {
                let (destination, middle, length) = bulk_memory_operand_widths(chunk, op, ip + 4, Some(type_ctx));
                for wide in [length, middle, destination] {
                    if stack.pop()? != if wide { TYPE_I64 } else { TYPE_I32 } {
                        return None;
                    }
                }
            }
            Op::DATA_DROP | Op::ELEM_DROP | Op::NOP => {}
            Op::DROP => { stack.pop()?; }
            Op::SELECT | Op::SELECT_T => {
                if stack.pop()? != TYPE_I32 { return None; }
                let right = stack.pop()?;
                if stack.pop()? != right || !matches!(right, TYPE_I32 | TYPE_I64 | TYPE_F32 | TYPE_F64 | 0x7b) {
                    return None;
                }
                stack.push(right);
            }
            Op::CALL_INDIRECT => {
                let scalar_types = |types: &str| -> Option<Vec<u8>> {
                    types.split(',').map(str::trim).filter(|ty| !ty.is_empty())
                        .map(|ty| match ty {
                            "i32" => Some(TYPE_I32), "i64" => Some(TYPE_I64),
                            "f32" => Some(TYPE_F32), "f64" => Some(TYPE_F64),
                            _ => None,
                        }).collect()
                };
                let (params, results) = chunk.call_indirect_sigs.get(&ip)?.split_once("->")?;
                let params = scalar_types(params)?;
                let results = scalar_types(results)?;
                if params.len() != usize::from(*chunk.code.get(ip + 4)?)
                    || results.len() != usize::from(*chunk.code.get(ip + 6)?) {
                    return None;
                }
                let table = usize::from(*chunk.code.get(ip + 5)?);
                if stack.pop()? != type_ctx.table_operands.get(table)?.0 { return None; }
                for param in params.into_iter().rev() {
                    if stack.pop()? != param { return None; }
                }
                stack.extend(results);
            }
            op if op.group() == 0xfd && matches!(op.sub(), 0x00..=0x0b | 0x54..=0x5d) => {
                let mut operand = ip + 4;
                let (_, _, memory) = read_optional_memarg(chunk, &mut operand, default_simd_align(op));
                let address = if *chunk.memory_is_64.get(memory as usize)? { TYPE_I64 } else { TYPE_I32 };
                let store = matches!(op.sub(), 0x0b | 0x58..=0x5b);
                let lane = matches!(op.sub(), 0x54..=0x5b);
                if store || lane {
                    if stack.pop()? != 0x7b { return None; }
                }
                if stack.pop()? != address { return None; }
                if !store { stack.push(0x7b); }
            }
            Op::V128_CONST => stack.push(0x7B),
            Op::I32_CONST => stack.push(TYPE_I32),
            Op::F32_CONST => stack.push(TYPE_F32),
            Op::F64_CONST => stack.push(TYPE_F64),
            Op::TABLE_SIZE | Op::TABLE_GET | Op::TABLE_SET | Op::TABLE_GROW | Op::TABLE_FILL => {
                let mut operand = ip + 4;
                let index = read_leb_u32(&chunk.code, &mut operand) as usize;
                let &(address, reference) = type_ctx.table_operands.get(index)?;
                match op {
                    Op::TABLE_SIZE => stack.push(address),
                    Op::TABLE_GET => {
                        if stack.pop()? != address { return None; }
                        stack.push(reference?);
                    }
                    Op::TABLE_SET => {
                        if stack.pop()? != reference? || stack.pop()? != address { return None; }
                    }
                    Op::TABLE_GROW => {
                        if stack.pop()? != address || stack.pop()? != reference? { return None; }
                        stack.push(address);
                    }
                    Op::TABLE_FILL => {
                        if stack.pop()? != address || stack.pop()? != reference?
                            || stack.pop()? != address { return None; }
                    }
                    _ => unreachable!(),
                }
            }
            Op::NULL if !saw_return => {
                let reference = *chunk.code.get(ip + 4)?;
                if !matches!(reference, 0x70 | 0x6f) { return None; }
                stack.push(reference);
            }
            Op::REF_IS_NULL => {
                if !matches!(stack.pop()?, 0x70 | 0x6f) { return None; }
                stack.push(TYPE_I32);
            }
            Op::MEMORY_ATOMIC_NOTIFY => {
                let mut operand = ip + 4;
                let align = read_leb_u32(&chunk.code, &mut operand);
                let address_type = if align & 0x80 != 0 { TYPE_I64 } else { TYPE_I32 };
                if stack.pop()? != TYPE_I32 || stack.pop()? != address_type { return None; }
                stack.push(TYPE_I32);
            }
            Op::I64_CONST => stack.push(TYPE_I64),
            Op::RETURN => {
                if stack.as_slice() != result_types.as_slice() {
                    return None;
                }
                stack.clear();
                saw_return = true;
            }
            Op::NULL
                if saw_return
                    && stack.is_empty()
                    && chunk.code.get(ip + 4).copied() == Some(HT_EXTERN)
                    && next_op(chunk, ip + 5) == Some(Op::RETURN)
                    && ip + 9 == chunk.code.len() =>
            {
                break;
            }
            Op::UNREACHABLE
                if saw_return
                    && stack.is_empty()
                    && ip + opcode_size(op, &chunk.code, ip) == chunk.code.len() =>
            {
                break;
            }
            _ => return None,
        }
        ip += opcode_size(op, &chunk.code, ip);
    }
    saw_return.then_some(RawNumericBody { local_types })
}

// Compiler-proven emit_next regions. Decoded ordinary Wasm has no marker and
// keeps its existing structured labels; VM handler instruction pointers remain
// untouched in every input Chunk.
#[derive(Clone, Copy)]
struct NativeNext {
    start: usize,
    resume: usize,
    handler: usize,
    end: usize,
    slot: u16,
    operands: usize,
    done_get: usize,
    completion_get: usize,
    completion_br: usize,
}
pub(crate) fn has_native_next(chunk: &Chunk) -> bool {
    !native_next_regions(chunk).is_empty() || !native_direct_resumes(chunk).is_empty()
}
fn native_growable_arrays(chunk:&Chunk)->std::collections::BTreeSet<usize> {
    let names:Vec<_>=chunk.local_names.iter().filter(|n|n.origin==vybe_runtime::chunk::LocalOrigin::Compiler && n.name.starts_with("<native-growable-array:")).collect();
    let mut found=std::collections::BTreeSet::new();
    if names.is_empty(){return found;}
    let mut boundaries=std::collections::BTreeSet::new();let mut ip=0;
    while ip<chunk.code.len(){boundaries.insert(ip);ip+=opcode_size(Op::from_code_at(&chunk.code,ip),&chunk.code,ip);}
    for n in names {
        let value=n.name.strip_prefix("<native-growable-array:").unwrap().strip_suffix('>').expect("closed growable marker");
        let at:usize=value.parse().expect("growable allocation offset");
        assert!(n.slot<chunk.local_count && boundaries.contains(&at));
        assert_eq!(Op::from_code_at(&chunk.code,at),Op::ARRAY_NEW_FIXED);
        assert_eq!(read_u16_at(&chunk.code,at+4),0);assert_eq!(read_u16_at(&chunk.code,at+6),0);
        assert_eq!(Op::from_code_at(&chunk.code,at+8),Op::LOCAL_SET);
        assert_eq!(read_u16_at(&chunk.code,at+12),n.slot,"same retained private buffer root");
        assert!(found.insert(at),"duplicate growable marker");
    }
    found
}
fn native_direct_resumes(chunk: &Chunk) -> Vec<usize> {
    // Compiler markers are the only source of native generator regions.
    // Avoid decoding marker-free bodies; marker validation below is unchanged.
    if !chunk.local_names.iter().any(|local| {
        local.origin == vybe_runtime::chunk::LocalOrigin::Compiler
            && local
                .name
                .strip_prefix("<native-generator-resume:")
                .and_then(|s| s.strip_suffix('>'))
                .is_some()
    }) {
        return Vec::new();
    }

    let mut boundaries = std::collections::BTreeSet::new();
    let mut ip = 0;
    while ip < chunk.code.len() {
        boundaries.insert(ip);
        ip += opcode_size(Op::from_code_at(&chunk.code, ip), &chunk.code, ip);
    }
    boundaries.insert(ip);
    let mut found = std::collections::BTreeSet::new();
    for local in &chunk.local_names {
        if local.origin != vybe_runtime::chunk::LocalOrigin::Compiler {
            continue;
        }
        let Some(fields) = local
            .name
            .strip_prefix("<native-generator-resume:")
            .and_then(|s| s.strip_suffix('>'))
        else {
            continue;
        };
        let fields: Vec<usize> = fields
            .split(':')
            .map(|x| x.parse().expect("compiler resume offset"))
            .collect();
        assert_eq!(fields.len(), 2);
        let (start, end) = (fields[0], fields[1]);
        assert!(boundaries.contains(&start) && boundaries.contains(&end));
        assert_eq!(end, start + 6);
        assert_eq!(Op::from_code_at(&chunk.code, start), Op::RESUME);
        assert_eq!(read_u16_at(&chunk.code, start + 4), 0);
        assert!(
            chunk
                .stack_switch_handlers
                .get(&start)
                .is_none_or(|h| h.is_empty()),
            "direct resume has no VM branch handlers"
        );
        assert!(found.insert(start), "duplicate direct resume marker");
    }
    found.into_iter().collect()
}
fn native_next_regions(chunk: &Chunk) -> Vec<NativeNext> {
    // Compiler markers are the only source of native generator regions.
    // Avoid decoding marker-free bodies; marker validation below is unchanged.
    if !chunk.local_names.iter().any(|local| {
        local.origin == vybe_runtime::chunk::LocalOrigin::Compiler
            && local
                .name
                .strip_prefix("<native-generator-next:")
                .and_then(|s| s.strip_suffix('>'))
                .is_some()
    }) {
        return Vec::new();
    }

    let mut instructions = Vec::new();
    let mut p = 0;
    while p < chunk.code.len() {
        let op = Op::from_code_at(&chunk.code, p);
        instructions.push((p, op));
        p += opcode_size(op, &chunk.code, p);
    }
    chunk
        .local_names
        .iter()
        .filter_map(|local| {
            if local.origin != vybe_runtime::chunk::LocalOrigin::Compiler {
                return None;
            }
            let s = local
                .name
                .strip_prefix("<native-generator-next:")?
                .strip_suffix('>')?;
            let fields: Vec<usize> = s
                .split(':')
                .map(|x| x.parse().expect("compiler next marker integer"))
                .collect();
            assert_eq!(fields.len(), 4, "compiler next marker arity");
            let (start, resume, handler, end) = (fields[0], fields[1], fields[2], fields[3]);
            let at = |offset| {
                instructions
                    .iter()
                    .position(|&(p, _)| p == offset)
                    .expect("next marker instruction boundary")
            };
            let a = at(start);
            let r = at(resume);
            let h = at(handler);
            assert!(start < resume && resume < handler && handler < end && end <= chunk.code.len());
            assert_eq!(instructions[a].1, Op::LOCAL_SET);
            assert_eq!(read_u16_at(&chunk.code, start + 4), local.slot);
            assert_eq!(instructions[a + 1].1, Op::BLOCK);
            assert_eq!(instructions[r].1, Op::RESUME);
            assert_eq!(instructions[r - 2].1, Op::LOCAL_GET);
            assert_eq!(
                read_u16_at(&chunk.code, instructions[r - 2].0 + 4),
                local.slot
            );
            assert_eq!(instructions[r - 1].1, Op::NULL);
            assert_eq!(chunk.code[instructions[r - 1].0 + 4], HT_EXTERN);
            assert_eq!(instructions[h - 1].1, Op::BR);
            assert_eq!(chunk.code[instructions[h - 1].0 + 4], 0);
            assert_eq!(instructions[h].1, Op::I32_CONST);
            assert_eq!(chunk.code[handler + 4], 1);
            assert_eq!(instructions[h + 1].1, Op::END);
            assert_eq!(instructions[h + 1].0 + 4, end);
            let handlers = chunk
                .stack_switch_handlers
                .get(&resume)
                .expect("compiler next handler");
            assert_eq!(handlers.len(), 1);
            assert_eq!(
                (
                    handlers[0].kind,
                    handlers[0].tag_index,
                    handlers[0].label_index
                ),
                (0, 0, handler as u32)
            );
            let done_get = instructions[a + 2..r - 2]
                .iter()
                .find(|&&(p, op)| {
                    op == Op::LOCAL_GET && read_u16_at(&chunk.code, p + 4) == local.slot
                })
                .expect("next done owner")
                .0;
            let completion_get = instructions[r + 1..h]
                .iter()
                .find(|&&(p, op)| {
                    op == Op::LOCAL_GET && read_u16_at(&chunk.code, p + 4) == local.slot
                })
                .expect("next completion owner")
                .0;
            Some(NativeNext {
                start,
                resume,
                handler,
                end,
                slot: local.slot,
                operands: instructions[r - 2].0,
                done_get,
                completion_get,
                completion_br: instructions[h - 1].0,
            })
        })
        .collect()
}
fn emit_cont_carrier_cast(body: &mut Vec<u8>, ctx: &WasmTypeContext) {
    emit_internalize(body);
    emit_ref_cast(body, ctx.continuation_carrier_type_idx);
}
fn emit_cont_carrier_field(body: &mut Vec<u8>, ctx: &WasmTypeContext, field: u32) {
    emit_cont_carrier_cast(body, ctx);
    body.push(0xfb);
    write_leb128_u32(body, 2);
    write_leb128_u32(body, ctx.continuation_carrier_type_idx);
    write_leb128_u32(body, field);
}

#[derive(Default)]
struct NativeDirectArgs {
    producers: std::collections::BTreeMap<usize, u8>,
    calls: std::collections::BTreeSet<usize>,
}
fn native_direct_args(chunk: &Chunk) -> NativeDirectArgs {
    let mut facts = NativeDirectArgs::default();
    let Some((params, _)) = &chunk.func_sig else { return facts; };
    if chunk.direct_defined_calls.is_empty() { return facts; }
    let params: Vec<_> = params.split(',').map(str::trim).filter(|s| !s.is_empty()).collect();
    let mut producers: Vec<(usize, &str)> = Vec::new();
    let mut ip = 0;
    while let Some(op) = next_op(chunk, ip) {
        if let Some(edge) = chunk.direct_defined_calls.get(&ip) {
            let args: Vec<_> = edge.params.split(',').map(str::trim).filter(|s| !s.is_empty()).collect();
            if !edge.source && producers.len() >= args.len() {
                let window = &producers[producers.len()-args.len()..];
                if window.iter().map(|(_, ty)| *ty).eq(args.iter().copied()) {
                    facts.calls.insert(ip);
                    facts.producers.extend(window.iter().map(|(at,ty)| (*at,super::types::val_type_byte(ty))));
                }
            }
            producers.clear();
        } else {
            let operand = ip + 4;
            let ty = match op {
                Op::I32_CONST => Some("i32"),
                Op::I64_CONST => Some("i64"),
                Op::F32_CONST => Some("f32"),
                Op::F64_CONST => Some("f64"),
                Op::LOCAL_GET if operand+2 <= chunk.code.len() => {
                    let slot=u16::from_be_bytes([chunk.code[operand],chunk.code[operand+1]]) as usize;
                    params.get(slot).copied()
                }
                Op::NULL => match chunk.code.get(operand) {
                    Some(0x6f)=>Some("externref"),Some(0x70)=>Some("funcref"),_=>None,
                },
                _ => None,
            };
            if let Some(ty)=ty {producers.push((ip,ty));} else {producers.clear();}
        }
        ip += opcode_size(op,&chunk.code,ip);
    }
    facts
}

// Declared Wasm function calls carry actual funcref, never the Source externref
// carrier. Give their reordered operands native typed locals; Source calls keep
// their existing carrier/argument/lifetime path untouched.
fn declared_ref_calls(chunk: &Chunk) -> Vec<(usize, Vec<u8>)> {
    let mut rows = Vec::new();
    for (&at, signature) in &chunk.call_indirect_sigs {
        if super::callable::source(chunk) { continue; }
        let Some(op) = Op::decode(
            u16::from_be_bytes([chunk.code[at], chunk.code[at + 1]]),
            u16::from_be_bytes([chunk.code[at + 2], chunk.code[at + 3]]),
        ) else {
            continue;
        };
        if !matches!(op, Op::CALL_REF | Op::RETURN_CALL_REF | Op::CALL_INDIRECT | Op::RETURN_CALL_INDIRECT) {
            continue;
        }
        let (params, results) = signature.split_once("->").expect("declared call signature");
        let mut params = params
            .split(',')
            .map(str::trim)
            .filter(|s| !s.is_empty())
            .map(super::types::val_type_byte)
            .collect::<Vec<_>>();
        assert_eq!(
            params.len(),
            usize::from(chunk.code[at + 4]),
            "declared reference call arity"
        );
        if matches!(op, Op::CALL_INDIRECT | Op::RETURN_CALL_INDIRECT) {
            // Result staging is needed only at a mixed-value boundary. Raw
            // verified bodies omit these locals entirely in the prefix.
            params.extend(results.split(',').map(str::trim).filter(|s| !s.is_empty())
                .map(super::types::val_type_byte));
        }
        rows.push((at, params));
    }
    for (&at, edge) in &chunk.direct_defined_calls {
        // Direct typed args/results get actual Wasm locals, not externref slots.
        let params = edge.params.split(',').chain(edge.results.split(','))
            .map(str::trim).filter(|s|!s.is_empty()).map(super::types::val_type_byte).collect();
        rows.push((at,params));
    }
    rows.sort_by_key(|(at, _)| *at);
    rows
}
#[derive(Default)]
struct DeclaredCallPlan {
    profiles: Vec<(u32, Vec<u8>)>,
    sites: std::collections::BTreeMap<usize, (usize, usize)>,
}

impl DeclaredCallPlan {
    fn new(chunk: &Chunk, type_ctx: &WasmTypeContext, raw: bool) -> Self {
        let mut plan = Self::default();
        if raw { return plan; }
    let (temps, i64temps) = count_temp_locals(chunk);
    let mut base = wasm_param_count(chunk).max(u32::from(chunk.local_count))
        + temps
        + i64temps
        + u32::from(has_native_next(chunk));
    if super::callable::source(chunk) {
        // Typed staging follows the real hidden environment and captured-cell
        // locals. Those indices are above hidden_virtual and are already
        // physical indices when callable::remap_locals visits the body.
        base += 1 + super::callable::captured_slots(chunk).len() as u32;
    }
        let mut layouts = std::collections::HashMap::<Vec<u8>, usize>::new();
        for (offset, mut types) in declared_ref_calls(chunk) {
            let values = types.len();
            // Scratch values never survive a call instruction. Reentrant
            // invocations have their own Wasm frame, so equal layouts can
            // share locals without aliasing callers, captures or user slots.
            if !chunk.direct_defined_calls.contains_key(&offset) {
                let selector = if matches!(next_op(chunk, offset), Some(Op::CALL_INDIRECT | Op::RETURN_CALL_INDIRECT)) {
                    type_ctx.table_operands.get(chunk.code[offset + 5] as usize)
                        .map_or(TYPE_I32, |table| table.0)
                } else { 0x70 };
                types.push(selector);
            }
            let profile = if let Some(&profile) = layouts.get(&types) {
                profile
            } else {
                let profile = plan.profiles.len();
                layouts.insert(types.clone(), profile);
                let next = base.checked_add(types.len() as u32).expect("typed local index overflow");
                plan.profiles.push((base, types));
                base = next;
                profile
            };
            plan.sites.insert(offset, (profile, values));
        }
        plan
    }

    fn slots(&self, at: usize) -> Option<(u32, &[u8])> {
        let &(profile, values) = self.sites.get(&at)?;
        let (base, types) = &self.profiles[profile];
        Some((*base, &types[..values]))
    }
}

fn declared_scratch_type(chunk: &Chunk, slot: u32) -> Option<u8> {
    let mut found = None;
    for n in &chunk.local_names {
        if u32::from(n.slot) == slot && n.origin == vybe_runtime::chunk::LocalOrigin::Compiler {
            if let Some(t) = n
                .name
                .strip_prefix("<native-wasm-scratch:")
                .and_then(|s| s.strip_suffix(">"))
            {
                let ty = super::types::val_type_byte(t);
                // A compiler-owned vector stays in a standard v128 local,
                // never in the Source externref ABI or a captured value cell.
                if super::callable::source(chunk)
                    && (ty != 0x7B || super::callable::captured_slots(chunk).contains(&(slot as u16)))
                {
                    return None;
                }
                assert!(found.is_none_or(|old| old == ty));
                found = Some(ty);
            }
        }
    }
    found
}

fn locals_groups(chunk: &Chunk, type_ctx: &WasmTypeContext) -> Vec<(u32, u8)> {
    let wasm_params = wasm_param_count(chunk);
    let raw_numeric = raw_numeric_body(chunk, type_ctx);
    let forwarded_types = raw_numeric
        .as_ref()
        .map(|body| body.local_types.clone())
        .unwrap_or_else(|| forwarded_numeric_local_types(chunk));
    let extra_locals = (chunk.local_count as u32).saturating_sub(wasm_params);
    let (temp_count, i64_temp_count) = if raw_numeric.is_some() {
        (0, 0)
    } else {
        count_temp_locals(chunk)
    };
    let mut groups: Vec<(u32, u8)> = Vec::new();
    let push = |ty: u8, groups: &mut Vec<(u32, u8)>| match groups.last_mut() {
        Some((n, t)) if *t == ty => *n += 1,
        _ => groups.push((1, ty)),
    };
    // Declared locals are indices `wasm_params .. wasm_params+extra_locals`,
    // and slot k IS local index k (slots below the declared count are params).
    for i in 0..extra_locals {
        let slot = wasm_params + i;
        let ty = if let Some(ty) = declared_scratch_type(chunk, slot) {
            ty
        } else if let Some(ty) = forwarded_types.get(slot as usize).copied().flatten() {
            ty
        } else if u16::try_from(slot).is_ok_and(|s| chunk.i64_locals.contains(&s)) {
            TYPE_I64
        } else {
            TYPE_EXTERNREF
        };
        push(ty, &mut groups);
    }
    for _ in 0..temp_count {
        push(TYPE_EXTERNREF, &mut groups);
    }
    for _ in 0..i64_temp_count {
        push(TYPE_I64, &mut groups);
    }
    groups
}

/// Write the locals prefix from [`locals_groups`].
fn write_locals_prefix(body: &mut Vec<u8>, chunk: &Chunk, type_ctx: &WasmTypeContext) {
    let plan = DeclaredCallPlan::new(chunk, type_ctx, raw_numeric_body(chunk, type_ctx).is_some());
    write_locals_prefix_with_plan(body, chunk, type_ctx, &plan);
}

fn write_locals_prefix_with_plan(body: &mut Vec<u8>, chunk: &Chunk, type_ctx: &WasmTypeContext, plan: &DeclaredCallPlan) {
    let groups = locals_groups(chunk, type_ctx);
    let typed_count: usize = plan.profiles.iter().map(|(_, types)| types.len()).sum();
    write_leb128_u32(
        body,
        groups.len() as u32
            + typed_count as u32
            + u32::from(has_native_next(chunk))
            + u32::from(
                super::callable::source(chunk)
                    && !super::callable::captured_slots(chunk).is_empty(),
            ),
    );
    for (count, ty) in groups {
        write_leb128_u32(body, count);
        body.push(ty);
    }
    if has_native_next(chunk) {
        write_leb128_u32(body, 1);
        body.push(0x63);
        write_leb128_i64(body, type_ctx.continuation_type_idx as i64);
    }
    if super::callable::source(chunk) {
        let n = super::callable::captured_slots(chunk).len();
        if n > 0 {
            write_leb128_u32(body, n as u32);
            body.push(TYPE_EXTERNREF);
        }
    }
    for (_, types) in &plan.profiles {
        for &ty in types {
            write_leb128_u32(body, 1);
            body.push(ty);
        }
    }
}

/// Size in bytes of the locals declaration prefix written at the start of a
/// function body. The spec says branch hint offsets are relative to this
/// position, so the hint scanner adds this value to every `chunk.code` offset.
pub(crate) fn locals_prefix_size(chunk: &Chunk, type_ctx: &WasmTypeContext) -> u32 {
    let mut prefix = Vec::new();
    write_locals_prefix(&mut prefix, chunk, type_ctx);
    prefix.len() as u32
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

fn bulk_memory_operand_widths(chunk: &Chunk, op: Op, operands_at: usize, type_ctx: Option<&WasmTypeContext>) -> (bool, bool, bool) {
    let mut cursor = operands_at;
    let first = read_leb_u32(&chunk.code, &mut cursor);
    let (dst, src) = match op {
        Op::MEMORY_COPY | Op::TABLE_COPY => (first, Some(read_leb_u32(&chunk.code, &mut cursor))),
        Op::MEMORY_INIT | Op::TABLE_INIT => (read_leb_u32(&chunk.code, &mut cursor), None),
        _ => (first, None),
    };
    let is_64 = |idx: u32| {
        if matches!(op, Op::TABLE_COPY | Op::TABLE_INIT) {
            if let Some(ctx) = type_ctx {
                return ctx.table_operands.get(idx as usize).is_some_and(|table| table.0 == TYPE_I64);
            }
        }
        let widths = if matches!(op, Op::TABLE_COPY | Op::TABLE_INIT) {
            &chunk.table_is_64
        } else { &chunk.memory_is_64 };
        widths
            .get(idx as usize)
            .copied()
            .unwrap_or(false)
    };
    let dst_64 = is_64(dst);
    match op {
        Op::MEMORY_COPY | Op::TABLE_COPY => {
            let src_64 = is_64(src.expect("memory.copy source index"));
            (dst_64, src_64, dst_64 && src_64)
        }
        Op::MEMORY_FILL => (dst_64, false, dst_64),
        Op::MEMORY_INIT | Op::TABLE_INIT => (dst_64, false, false),
        _ => unreachable!("not a bulk-memory operation"),
    }
}

pub(crate) fn count_temp_locals(chunk: &Chunk) -> (u32, u32) {
    let mut need = if !has_native_next(chunk) { 0 } else { 2 };
    need = need.max(u32::from(super::callable::source(chunk)));
    let mut i64_need = 0u32;
    let mut ip = 0;
    while ip < chunk.code.len() {
        if ip + 3 >= chunk.code.len() {
            break;
        }
        if let Some(op) = Op::decode(
            ((chunk.code[ip] as u16) << 8) | chunk.code[ip + 1] as u16,
            ((chunk.code[ip + 2] as u16) << 8) | chunk.code[ip + 3] as u16,
        ) {
            if op == Op::ARRAY_NEW_FIXED && read_u16_at(&chunk.code, ip + 4) != 0 {
                // Declared packed arrays normalize all fixed initializer words
                // in their original order; operand zero retains its old layout.
                need = need.max(read_u16_at(&chunk.code, ip + 6) as u32);
            } else if matches!(op, Op::ARRAY_NEW | Op::ARRAY_NEW_DATA | Op::ARRAY_GET_S | Op::ARRAY_GET_U)
                && read_u16_at(&chunk.code, ip + 4) != 0
            {
                need = need.max(1);
            } else if !chunk.direct_defined_calls.contains_key(&ip) && (op == Op::CALL_REF || op == Op::RETURN_CALL || op == Op::RETURN_CALL_REF) {
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
            } else if let Some((_, value_type)) = memory_store_shape(op) {
                if value_type == TYPE_I64 {
                    i64_need = i64_need.max(1);
                } else {
                    need = need.max(1);
                }
            } else if atomic_is_i32_two_operand(op)
                || op == Op::MEMORY_ATOMIC_NOTIFY
            {
                need = need.max(1);
            } else if matches!(op, Op::MEMORY_COPY | Op::MEMORY_INIT | Op::MEMORY_FILL | Op::TABLE_COPY | Op::TABLE_INIT) {
                need = need.max(2);
                let (_, middle_64, length_64) = bulk_memory_operand_widths(chunk, op, ip + 4, None);
                i64_need = i64_need.max(middle_64 as u32 + length_64 as u32);
                // The module owns table limits/types. This pre-context scan
                // must reserve enough staging even when a function carries
                // no redundant copy of the table declarations.
                if matches!(op, Op::TABLE_COPY | Op::TABLE_INIT) { i64_need = i64_need.max(2); }
            } else if op == Op::TABLE_GROW {
                need = need.max(2);
            } else if op == Op::MEMORY_ATOMIC_WAIT32 {
                need = need.max(1);
                i64_need = i64_need.max(1);
            } else if atomic_is_i32_cmpxchg(op) {
                // expected + replacement sit above the address.
                need = need.max(2);
            } else if op == Op::CALL {
                // Spill temps for `emit_unbox_raw_params` — an import with
                // a raw param under another argument. Reserved for every
                // static call because the import tables are not built yet
                // here; two externref locals are cheaper than threading them in.
                need = need.max(2);
                let import_idx = read_u16_at(&chunk.code, ip + 4) as usize;
                if chunk
                    .imports
                    .get(import_idx)
                    .is_some_and(|import| import.module == "ecma:array" && import.name == "splice")
                {
                    let argc = chunk.code.get(ip + 6).copied().unwrap_or(0) as u32;
                    need = need.max(argc.saturating_sub(2));
                }
            } else if op == Op::ARRAY_COPY {
                need = need.max(5); // boxed destination, offsets, source, and count
            } else if op == Op::ARRAY_SET {
                need = need.max(3); // value, index, and receiver for GC-vs-dynamic branch
            } else if op == Op::ARRAY_GET {
                need = need.max(2); // index and receiver for GC-vs-dynamic branch
            } else if op == Op::STRUCT_SET {
                need = need.max(2); // need 2 temps for 3-operand reorder
            } else if op == Op::DESC_SET_PROTO {
                // Expands to local.set / global.get / local.get / struct.set —
                // one temp to get the descriptor ref UNDER the prototype value.
                need = need.max(1);
            } else if is_binary_typed_op(op) || op == Op::ARRAY_LENGTH || op == Op::REF_IS_NULL {
                need = need.max(1);
            } else if op == Op::REF_EQ {
                need = need.max(2);
            }
            // Phase E: the `0xFF` ARRAY_* (PUSH/POP/SLICE/JOIN/REVERSE/
            // CONTAINS/INDEX_OF/CONTAINS/CONCAT/SHIFT) no longer exist.
            // Their call-site temps are no longer needed here either.
            ip += opcode_size(op, &chunk.code, ip);
        } else {
            ip += 4;
        }
    }
    (need, i64_need)
}

/// Is this a core WASM op that operates on typed (non-externref) values
/// and takes two operands from the stack?
fn is_binary_typed_op(op: Op) -> bool {
    // f32 binary arithmetic and comparisons
    op == Op::F32_ADD || op == Op::F32_SUB || op == Op::F32_MUL || op == Op::F32_DIV
    || op == Op::F32_MIN || op == Op::F32_MAX || op == Op::F32_COPYSIGN
    || op == Op::F32_LT || op == Op::F32_GT || op == Op::F32_LE || op == Op::F32_GE
    || op == Op::F32_EQ || op == Op::F32_NE
    // f64 binary arithmetic
    || op == Op::F64_ADD || op == Op::F64_SUB || op == Op::F64_MUL || op == Op::F64_DIV
    || op == Op::F64_MIN || op == Op::F64_MAX || op == Op::F64_COPYSIGN
    // f64 comparisons
    || op == Op::F64_LT || op == Op::F64_GT || op == Op::F64_LE || op == Op::F64_GE
    || op == Op::F64_EQ || op == Op::F64_NE
    // i32 binary arithmetic
    || op == Op::I32_ADD || op == Op::I32_SUB || op == Op::I32_MUL
    || op == Op::I32_DIV_S || op == Op::I32_DIV_U
    || op == Op::I32_REM_S || op == Op::I32_REM_U
    || op == Op::I32_AND || op == Op::I32_OR || op == Op::I32_XOR
    || op == Op::I32_SHL || op == Op::I32_SHR_S || op == Op::I32_SHR_U
    || op == Op::I32_ROTL || op == Op::I32_ROTR
    // i32 comparisons
    || op == Op::EQ || op == Op::NE
    || op == Op::I32_LT_S || op == Op::I32_LT_U
    || op == Op::I32_GT_S || op == Op::I32_GT_U
    || op == Op::I32_LE_S || op == Op::I32_LE_U
    || op == Op::I32_GE_S || op == Op::I32_GE_U
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
    field_name_idx: u32,
) -> Option<u32> {
    let value = chunk.constants.get(field_name_idx as usize)?;
    let text = format!("{value}");
    type_ctx.string_const_global.get(&text).copied()
}

fn wasm_struct_field_by_typeidx(
    chunk: &Chunk,
    type_ctx: &WasmTypeContext,
    typeidx: u32,
    field_name_idx: u32,
) -> Option<(u32, u32)> {
    if typeidx == 0 {
        return None;
    }
    // ⛔⛔ TWO FAULTS HERE, AND EACH ONE ALONE MADE THIS RETURN `None`.
    //
    // 1. THE OPERAND *IS* THE FIELD INDEX when the typeidx is real. It was read
    //    as a CONSTANT-POOL index and looked up as a field NAME. Every other
    //    side disagrees: the compiler emits `ResolvedSlot::Indexed`'s `field`
    //    (a `position()` result) at `class_slots.rs:463/507`, the wasm reader
    //    emits `field_idx` at `reader/mod.rs:2883/2893`, and the VM reads it as
    //    an index — `dispatch.rs:4237`, `if typeidx != 0 { /* struct.set $t i */ }`.
    //
    // 2. `chunk.types` IS EMPTY IN EVERY CHUNK BUT 0, and constructors are
    //    exactly where field writes live. Same trap `wasm_struct_type_for_chunk_type`
    //    and `descriptor_global_for_chunk_type` above already document; this
    //    function still had the by-name, by-chunk form. `struct_type_by_index`
    //    answers without a name and without a chunk.
    //
    // Either fault alone yields `None`, and the caller then falls through to
    // `wasm_struct_field_for_name`'s `(0, 0)` sentinel — where type 0 is a REAL
    // type, the class's own. Measured on `class Point { a; b; e; }`: all three
    // initialisers emitted `struct.set $Point $a`, writing `b` and `e` into
    // `a`'s slot. It reads as plausible in the disassembly, and the program
    // still printed `1 2 3` because reads go through the name-keyed property
    // map — so the indexed slots were wrong and nothing could observe it.
    //
    // ⛔ No bounds check is possible here (the field count lives on chunk 0),
    // and none is needed: the compiler derived the index from `position()` over
    // that very field list, so it is in range by construction.
    let _ = chunk;
    Some((
        type_ctx.struct_type_by_index(typeidx as u32)?,
        field_name_idx as u32,
    ))
}

fn wasm_struct_field_for_name(
    chunk: &Chunk,
    type_ctx: &WasmTypeContext,
    field_name_idx: u32,
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
        None => body.push(if results == 0 {
            TYPE_VOID
        } else {
            TYPE_EXTERNREF
        }),
    }
}

pub fn encode_code_section(
    chunks: &[Chunk],
    host_imports: &[HostImport],
    rt_imports: &[(&str, &str)],
    type_ctx: &WasmTypeContext,
    tag_plan: &crate::writer::proposals::exception_handling::ModuleTagPlan,
) -> Vec<u8> {
    let host_import_count = host_imports.len();
    let host_import_map = crate::writer::sections::host_import_index_map(host_imports);

    // Build import name → function index map (module:name for uniqueness)
    let mut rt_idx: std::collections::HashMap<(&str, &str), usize> =
        std::collections::HashMap::new();
    for (i, &(module, name)) in rt_imports.iter().enumerate() {
        rt_idx.insert((module, name), host_import_count + i);
    }

    let mut out = Vec::new();
    write_leb128_u32(
        &mut out,
        chunks.len() as u32
            + 2 * u32::from(type_ctx.native_resume_function_idx.is_some())
            + type_ctx.generators.len() as u32
            + u32::from(type_ctx.source_entry.is_some()),
    );

    for (ci, chunk) in chunks.iter().enumerate() {
        let mut body = Vec::new();
        let raw_numeric = raw_numeric_body(chunk, type_ctx);
        let declared_calls = DeclaredCallPlan::new(chunk, type_ctx, raw_numeric.is_some());
        let i32_regions = SourceI32Regions::new(chunk, &chunks[0], type_ctx, &host_import_map);
        let native_args = native_direct_args(chunk);
        let raw_i32_branch_offsets = i32_branch_consumers(chunk);
        let raw_i32_branches = &raw_i32_branch_offsets;
        let forwarded_types = raw_numeric
            .as_ref()
            .map(|body| body.local_types.clone())
            .unwrap_or_else(|| forwarded_numeric_local_types(chunk));

        // Check how many temp locals we need for stack manipulation
        let (temp_count, _) = if raw_numeric.is_some() {
            (0, 0)
        } else {
            count_temp_locals(chunk)
        };
        let has_temp = temp_count > 0;
        // WAST chunks can have a declared signature different from internal arity.
        let wasm_params = wasm_param_count(chunk);
        // Extra locals beyond params
        let extra_locals = if chunk.local_count as u32 > wasm_params {
            chunk.local_count as u32 - wasm_params
        } else {
            0
        };
        let temp_local_idx = wasm_params + extra_locals;

        // Locals declaration (only declare locals beyond params) — grouped by
        // type, because `i64_locals` makes them heterogeneous.
        write_locals_prefix_with_plan(&mut body, chunk, type_ctx, &declared_calls);
        let locals_prefix_len = body.len();
        if ci == 0 {
            super::callable::initialize_aliases(
                &mut body,
                chunks,
                &type_ctx.callables,
                temp_local_idx,
            );
        }
        if let Some(e) = type_ctx.generators.iter().find(|e| e.chunk_index == ci) {
            // Actual entry control is not necessarily arity-1 after the compiler
            // reserves its environment. Copy incoming control before captures.
            super::callable::local(&mut body, 0x20, u32::from(e.user_arity));
            super::callable::local(&mut body, 0x21, u32::from(e.control_slot));
        }
        super::callable::initialize_cells(
            &mut body,
            chunk,
            &type_ctx.callables,
            total_temp_locals(chunk),
        );

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
        let mut block_result_stack: Vec<u8> = Vec::new();
        let mut else_branch_stack: Vec<bool> = Vec::new();
        let mut pending_i32_box_after_end = false;
        let mut pending_externref_spilled_after_end = false;
        let mut pending_externref_on_stack = false;
        let mut drop_carried_after_bytecode_drop: Option<usize> = None;
        let carried_struct_sets = carried_local_tee_struct_sets(chunk);
        let native_next = native_next_regions(chunk);
        let native_direct = native_direct_resumes(chunk);
        let native_growable = native_growable_arrays(chunk);
        let native_cont_local = temp_local_idx + temp_count + count_temp_locals(chunk).1;

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
            if native_growable.contains(&ip) {
                assert_eq!(op,Op::ARRAY_NEW_FIXED,"growable producer opcode");
                assert_eq!(read_u16_at(&chunk.code,ip+4),0,"private generic array");
                assert_eq!(read_u16_at(&chunk.code,ip+6),0,"private empty growable buffer");
                body.push(0x41); write_leb128_i32(&mut body,0);
                emit_import_call(&mut body,&rt_idx,"ecma:array","newWithLength");
                ip+=8;continue;
            }
            // The shared compiler appends `ref.null extern; return` as a
            // fallthrough safety net. A core function with a numeric result
            // cannot return that ref, even when an earlier return is reached.
            if op == Op::NULL
                && chunk.func_sig.as_ref().is_some_and(|(_, result)| {
                    let mut results = result.split(',').filter(|ty| !ty.trim().is_empty());
                    results.clone().next().is_some()
                        && results.all(|ty| matches!(ty.trim(), "i32" | "i64" | "f32" | "f64"))
                })
                && chunk.code.get(ip + 4).copied() == Some(HT_EXTERN)
                && next_op(chunk, ip + 5) == Some(Op::RETURN)
                && ip + 9 == chunk.code.len()
            {
                body.push(0x00); // unreachable
                break;
            }
            // Close the native-only yield landing; retain the transferred
            // continuation in the same carrier before executing the VM yield arm.
            if let Some(next) = native_next.iter().find(|n| n.handler == ip) {
                body.push(0x0b);
                body.push(0x21);
                write_leb128_u32(&mut body, native_cont_local);
                body.push(0x20);
                write_leb128_u32(&mut body, next.slot as u32);
                emit_cont_carrier_cast(&mut body, type_ctx);
                body.push(0x20);
                write_leb128_u32(&mut body, native_cont_local);
                body.push(0xfb);
                write_leb128_u32(&mut body, 5);
                write_leb128_u32(&mut body, type_ctx.continuation_carrier_type_idx);
                write_leb128_u32(&mut body, 0);
            }
            if native_next.iter().any(|n| n.operands == ip) {
                body.push(0x02);
                write_leb128_i64(&mut body, type_ctx.continuation_landing_type_idx as i64);
            }
            if let Some(next) = native_next
                .iter()
                .find(|n| n.done_get == ip || n.completion_get == ip)
            {
                body.push(0x20);
                write_leb128_u32(&mut body, next.slot as u32);
                emit_cont_carrier_field(&mut body, type_ctx, 1);
                ip += 6;
                continue;
            }
            if native_next.iter().any(|n| n.completion_br == ip) {
                body.push(0x0c);
                write_leb128_u32(&mut body, 1);
                ip += opcode_size(op, &chunk.code, ip);
                continue;
            }
            if let Some(_next) = native_next.iter().find(|n| n.resume == ip) {
                // VM [carrier, value] -> native [value, typed continuation].
                body.push(0x21);
                write_leb128_u32(&mut body, temp_local_idx + 1);
                body.push(0x21);
                write_leb128_u32(&mut body, temp_local_idx);
                body.push(0x20);write_leb128_u32(&mut body,temp_local_idx);
                body.push(0x20);write_leb128_u32(&mut body,temp_local_idx+1);
                body.push(0x10);write_leb128_u32(&mut body,type_ctx.native_carrier_resume_function_idx.expect("carrier service"));
                // Yield branches with [payload,new cont]. Completion drops
                // null cont and executes the original completion side effects.
                body.push(0x22);
                write_leb128_u32(&mut body, native_cont_local);
                body.push(0x20);
                write_leb128_u32(&mut body, native_cont_local);
                body.push(0xd1);
                body.push(0x45);
                body.push(0x0d);
                write_leb128_u32(&mut body, 0);
                body.push(0x1a);
                ip += 6;
                continue;
            }
            if native_direct.contains(&ip) {
                // The raw VM resume producer has no branch handlers. Preserve
                // its one-result contract while replacing the carrier's own
                // continuation after either suspension or completion.
                body.push(0x21);
                write_leb128_u32(&mut body, temp_local_idx + 1);
                body.push(0x21);
                write_leb128_u32(&mut body, temp_local_idx);
                body.push(0x20);write_leb128_u32(&mut body,temp_local_idx);
                body.push(0x20);write_leb128_u32(&mut body,temp_local_idx+1);
                body.push(0x10);write_leb128_u32(&mut body,type_ctx.native_carrier_resume_function_idx.expect("carrier service"));
                body.push(0x21);
                write_leb128_u32(&mut body, native_cont_local);
                body.push(0x20);
                write_leb128_u32(&mut body, temp_local_idx);
                emit_cont_carrier_cast(&mut body, type_ctx);
                body.push(0x20);
                write_leb128_u32(&mut body, native_cont_local);
                body.push(0xfb);
                write_leb128_u32(&mut body, 5);
                write_leb128_u32(&mut body, type_ctx.continuation_carrier_type_idx);
                write_leb128_u32(&mut body, 0);
                ip += 6;
                continue;
            }
            if pending_externref_spilled_after_end
                && (op != Op::END || block_result_stack.last().copied().unwrap_or(0) != 0)
            {
                body.push(0x20); // local.get $temp
                write_leb128_u32(&mut body, temp_local_idx);
                pending_externref_spilled_after_end = false;
                pending_externref_on_stack = true;
            }
            if pending_externref_on_stack
                && op == Op::END
                && block_result_stack.last().copied().unwrap_or(0) == 0
            {
                body.push(0x21); // local.set $temp
                write_leb128_u32(&mut body, temp_local_idx);
                pending_externref_on_stack = false;
                pending_externref_spilled_after_end = true;
            }
            if pending_i32_box_after_end
                && op == Op::END
                && !i32_block_stack.last().copied().unwrap_or(false)
            {
                emit_box_i32(&mut body, &rt_idx);
                if block_result_stack.last().copied().unwrap_or(0) == 0 {
                    body.push(0x21); // local.set $temp
                    write_leb128_u32(&mut body, temp_local_idx);
                    pending_externref_spilled_after_end = true;
                    pending_externref_on_stack = false;
                } else {
                    pending_externref_on_stack = true;
                }
                pending_i32_box_after_end = false;
            } else if pending_i32_box_after_end && op != Op::END {
                box_i32_unless_condition(
                    &mut body,
                    &rt_idx,
                    chunk,
                    ip,
                    i32_block_stack.last().copied().unwrap_or(false), raw_i32_branches);
                pending_i32_box_after_end = false;
            }

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
                i32_block_stack.push(false);
                block_result_stack.push(results);
                else_branch_stack.push(false);
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
            if op == Op::LOCAL_TEE {
                let slot = read_u16_at(&chunk.code, ip);
                if block_result_stack.last().copied().unwrap_or(0) == 0 {
                    drop_carried_after_bytecode_drop =
                        key_registration_drop_before_void_boundary(chunk, ip + 2, slot)
                            .or(drop_carried_after_bytecode_drop);
                }
            }
            if let Some((trace_start, trace_end)) =
                std::env::var("VYBE_WASM_TRACE_BODY").ok().and_then(|raw| {
                    let (start, end) = raw.split_once(':')?;
                    Some((start.parse::<usize>().ok()?, end.parse::<usize>().ok()?))
                })
            {
                let body_at = body.len();
                let chunk_matches = std::env::var("VYBE_WASM_TRACE_CHUNK")
                    .map(|name| name == chunk.name)
                    .unwrap_or(true);
                if chunk_matches && (trace_start..=trace_end).contains(&body_at) {
                    if op == Op::CALL {
                        let funcidx = read_u16_at(&chunk.code, ip);
                        let argc = chunk.code.get(ip + 2).copied().unwrap_or(0);
                        let import = chunks
                            .first()
                            .and_then(|root| root.imports.get(funcidx as usize))
                            .map(|imp| format!("{}:{}", imp.module, imp.name))
                            .unwrap_or_else(|| "<unknown>".to_string());
                        eprintln!(
                            "[wasm-body] body_index={} chunk={} ip={} body={} op={:?} func={} argc={} import={} raw_result={:?} params={:?}",
                            ci,
                            chunk.name,
                            op_start,
                            body_at,
                            op,
                            funcidx,
                            argc,
                            import,
                            type_ctx.raw_result_funcs.get(&(funcidx as u32)),
                            type_ctx.import_param_funcs.get(&(funcidx as u32))
                        );
                    } else {
                        eprintln!(
                            "[wasm-body] body_index={} chunk={} ip={} body={} op={:?}",
                            ci, chunk.name, op_start, body_at, op
                        );
                    }
                }
            }

            // ⛔ RAW-REGION TRACKING, NOT A TYPE CHECKER. `block_i32_results`
            // is the compiler naming the blocks whose result is a bare i32 —
            // the truthiness ladder and nothing else. All this stack answers is
            // "is the value I am about to emit an arm result of one of those",
            // so the i32 producers inside a ladder are left unboxed.
            let in_i32_block = i32_block_stack.last().copied().unwrap_or(false);
            let current_block_results = block_result_stack.last().copied().unwrap_or(0);
            let mut closed_i32_region = false;
            if op == Op::BLOCK || op == Op::IF || op == Op::LOOP {
                i32_block_stack.push(chunk.block_i32_results.contains(&op_start));
                else_branch_stack.push(false);
                let (_, effective_results) = effective_structured_block_counts(chunk, op, op_start);
                block_result_stack.push(effective_results);
            } else if op == Op::END {
                // ⛔ THE RAW REGION HAS TO REJOIN THE VALUE ABI. A truthiness
                // ladder is a bare i32 throughout — that is what
                // `block_i32_results` declares — but at its OUTERMOST `end` the
                // answer stops being an arm result and becomes an ordinary
                // value: stored to a local, passed to a call, returned. All of
                // those are externref, and the i32 was reaching a `local.set`
                // on an externref local.
                //
                // Every closed i32-result block must rejoin the value ABI
                // unless its result is immediately consumed by another raw
                // i32 op or carried as an arm result of an enclosing i32 block.
                let closed = i32_block_stack.pop().unwrap_or(false);
                block_result_stack.pop();
                else_branch_stack.pop();
                closed_i32_region = closed;
            } else if op == Op::ELSE {
                if let Some(in_else) = else_branch_stack.last_mut() {
                    *in_else = true;
                }
            }
            let current_else_branch = else_branch_stack.last().copied().unwrap_or(false);

            if op.group() == 0x00 && !op.is_vm_internal() {
                emit_core_op(
                    &mut body,
                    op,
                    chunk,
                    ci,
                    &mut ip,
                    op_start,
                    chunks,
                    &rt_idx,
                    temp_local_idx,
                    temp_local_idx + temp_count,
                    has_temp,
                    type_ctx,
                    &host_import_map,
                    tag_plan,
                    in_i32_block,
                    current_block_results,
                    &forwarded_types,
                    raw_numeric.as_ref(),
                    &native_args, raw_i32_branches, &declared_calls, &i32_regions);
                if op == Op::DROP && drop_carried_after_bytecode_drop == Some(op_start) {
                    body.push(0x1A);
                    drop_carried_after_bytecode_drop = None;
                }
            } else if op.group() == 0xFB {
                emit_gc_op(
                    &mut body,
                    op,
                    chunk,
                    &mut ip,
                    op_start,
                    &carried_struct_sets,
                    &rt_idx,
                    type_ctx,
                    temp_local_idx,
                    in_i32_block, raw_i32_branches, i32_regions.value_strings.contains(&op_start));
            } else if raw_numeric.is_some() && raw_scalar_shape(op).is_some() {
                body.push(op.group() as u8);
                write_leb128_u32(&mut body, op.sub() as u32);
            } else if op == Op::I32_TRUNC_SAT_F64_S || op == Op::I32_TRUNC_SAT_F64_U {
                // Saturating conversion has the same boxed value boundary as
                // ordinary f64-to-i32 truncation. Serializing only the 0xFC
                // opcode would feed an externref to a raw f64 instruction.
                emit_unbox_f64(&mut body, &rt_idx);
                body.push(op.group() as u8);
                write_leb128_u32(&mut body, op.sub() as u32);
                box_i32_unless_condition(&mut body, &rt_idx, chunk, ip, in_i32_block, raw_i32_branches);
            } else if op.group() == 0xFC {
                // 0xFC-prefix ops per the bulk-memory / reference-types spec
                // need specific trailing immediates that aren't captured by
                // our `operand_format` in bytecode. Translate each case.
                let boxed_table_growth = op == Op::TABLE_GROW && raw_numeric.is_none()
                    && super::callable::source(chunk);
                if boxed_table_growth {
                    // Source callables are externref carriers; table 0 stores funcrefs.
                    super::callable::local(&mut body, 0x21, temp_local_idx);
                    super::callable::local(&mut body, 0x22, temp_local_idx + 1);
                    body.push(0xd1); // ref.is_null
                    body.extend_from_slice(&[0x04, 0x70]); // if (result funcref)
                    super::callable::null(&mut body, -16);
                    body.push(0x05);
                    super::callable::local(&mut body, 0x20, temp_local_idx + 1);
                    super::callable::internal(&mut body, type_ctx.callables.carrier);
                    super::callable::field(&mut body, 2, type_ctx.callables.carrier, 0);
                    body.push(0x0b);
                    super::callable::local(&mut body, 0x20, temp_local_idx);
                    emit_unbox_i32(&mut body, &rt_idx);
                }
                if raw_numeric.is_none()
                    && matches!(op, Op::MEMORY_COPY | Op::MEMORY_INIT | Op::MEMORY_FILL | Op::TABLE_COPY | Op::TABLE_INIT) {
                    let (dst_64, middle_64, length_64) = bulk_memory_operand_widths(chunk, op, ip, Some(type_ctx));
                    let i64_temp = temp_local_idx + temp_count;
                    let middle_temp = if middle_64 { i64_temp } else { temp_local_idx };
                    let length_temp = if length_64 {
                        i64_temp + middle_64 as u32
                    } else {
                        temp_local_idx + 1
                    };
                    body.push(0x21); // local.set length
                    write_leb128_u32(&mut body, length_temp);
                    body.push(0x21); // local.set source or fill value
                    write_leb128_u32(&mut body, middle_temp);
                    if !dst_64 {
                        emit_unbox_i32(&mut body, &rt_idx);
                    }
                    body.push(0x20); // local.get source or fill value
                    write_leb128_u32(&mut body, middle_temp);
                    if !middle_64 {
                        emit_unbox_i32(&mut body, &rt_idx);
                    }
                    body.push(0x20); // local.get length
                    write_leb128_u32(&mut body, length_temp);
                    if !length_64 {
                        emit_unbox_i32(&mut body, &rt_idx);
                    }
                }
                body.push(op.group() as u8);
                write_leb128_u32(&mut body, op.sub() as u32);
                match op {
                    Op::MEMORY_INIT => {
                        // spec: data_idx, memory_idx — LEB u32 on both sides now.
                        let data_idx = read_leb_u32(&chunk.code, &mut ip);
                        let memidx = read_leb_u32(&chunk.code, &mut ip);
                        write_leb128_u32(&mut body, data_idx);
                        write_leb128_u32(&mut body, memidx);
                    }
                    Op::DATA_DROP => {
                        let data_idx = read_leb_u32(&chunk.code, &mut ip);
                        write_leb128_u32(&mut body, data_idx);
                    }
                    Op::MEMORY_COPY => {
                        // spec: dst_mem, src_mem
                        let dst_mem = read_leb_u32(&chunk.code, &mut ip);
                        let src_mem = read_leb_u32(&chunk.code, &mut ip);
                        write_leb128_u32(&mut body, dst_mem);
                        write_leb128_u32(&mut body, src_mem);
                    }
                    Op::MEMORY_FILL => {
                        let memidx = read_leb_u32(&chunk.code, &mut ip);
                        write_leb128_u32(&mut body, memidx);
                    }
                    Op::TABLE_INIT => {
                        let elem_idx = read_leb_u32(&chunk.code, &mut ip);
                        let table_idx = read_leb_u32(&chunk.code, &mut ip);
                        write_leb128_u32(&mut body, elem_idx);
                        write_leb128_u32(&mut body, table_idx);
                    }
                    Op::ELEM_DROP => {
                        let elem_idx = read_leb_u32(&chunk.code, &mut ip);
                        write_leb128_u32(&mut body, elem_idx);
                    }
                    Op::TABLE_COPY => {
                        let dst_table = read_leb_u32(&chunk.code, &mut ip);
                        let src_table = read_leb_u32(&chunk.code, &mut ip);
                        write_leb128_u32(&mut body, dst_table);
                        write_leb128_u32(&mut body, src_table);
                    }
                    Op::TABLE_GROW | Op::TABLE_SIZE | Op::TABLE_FILL => {
                        let table_idx = read_leb_u32(&chunk.code, &mut ip);
                        write_leb128_u32(&mut body, table_idx);
                    }
                    _ if (0x00..=0x07).contains(&(op.sub() as u32)) => {
                        // nontrapping float-to-int conversions have no trailing
                        // immediates in the standard 0xFC space.
                    }
                    _ => {
                        panic!(
                            "unsupported 0xFC opcode in WASM writer: {:#x} {:#x}",
                            op.group(),
                            op.sub()
                        );
                    }
                }
                if boxed_table_growth {
                    box_i32_unless_condition(&mut body, &rt_idx, chunk, ip, in_i32_block, raw_i32_branches);
                }
            } else if op.group() == 0xFD {
                if raw_numeric.is_none() && matches!(op, Op::I8X16_SPLAT | Op::I8X16_SHR_U) {
                    emit_unbox_i32(&mut body, &rt_idx);
                }
                // Mixed Value bodies box scalar operands; this standard SIMD
                // island keeps v128 operands raw. Restrict the address bridge
                // to the exact proven memory32 memarg.
                if raw_numeric.is_none() && matches!(op, Op::V128_LOAD | Op::V128_LOAD32_ZERO)
                    && chunk.code.get(ip..ip + 3) == Some(&[0x80, 0x01, 0][..])
                    && !chunk.memory_is_64.first().copied().unwrap_or(false)
                {
                    emit_unbox_i32(&mut body, &rt_idx);
                }
                emit_simd_prefixed_op(&mut body, op, chunk, &mut ip);
                if raw_numeric.is_none() && op == Op::I8X16_BITMASK {
                    box_i32_unless_condition(&mut body, &rt_idx, chunk, ip, in_i32_block, raw_i32_branches);
                }
            } else if op.group() == 0xFE {
                emit_thread_prefixed_op(
                    &mut body,
                    op,
                    chunk,
                    &mut ip,
                    &rt_idx,
                    temp_local_idx,
                    temp_local_idx + temp_count,
                    in_i32_block,
                    raw_numeric.is_some(), raw_i32_branches);
            } else {
                emit_vm_internal_op(&mut body, op, chunk, &mut ip);
            }

            let parent_keeps_raw_arm = i32_block_stack.last().copied().unwrap_or(false)
                || block_result_stack.last().copied().unwrap_or(0) == 0;
            if closed_i32_region
                && !(parent_keeps_raw_arm
                    && matches!(next_op(chunk, ip), Some(next) if next == Op::ELSE || (next == Op::END && current_else_branch)))
            {
                if matches!(next_op(chunk, ip), Some(next) if next == Op::END)
                    && i32_block_stack.last().copied().unwrap_or(false)
                {
                    pending_i32_box_after_end = true;
                } else {
                    pending_i32_box_after_end = false;
                    box_i32_unless_condition(
                        &mut body,
                        &rt_idx,
                        chunk,
                        ip,
                        i32_block_stack.last().copied().unwrap_or(false), raw_i32_branches);
                    if matches!(next_op(chunk, ip), Some(next) if next == Op::END) {
                        if block_result_stack.last().copied().unwrap_or(0) == 0 {
                            body.push(0x21); // local.set $temp
                            write_leb128_u32(&mut body, temp_local_idx);
                            pending_externref_spilled_after_end = true;
                            pending_externref_on_stack = false;
                        } else {
                            pending_externref_on_stack = true;
                        }
                    }
                }
            }
            let at_void_branch_boundary = matches!(next_op(chunk, ip), Some(next) if next == Op::ELSE || (next == Op::END && current_else_branch));
            let direct_void_branch_value =
                current_block_results == 0 && void_boundary_value_producer(op);
            let closed_nested_value_in_void_branch = op == Op::END
                && current_block_results > 0
                && block_result_stack.last().copied().unwrap_or(0) == 0;
            if at_void_branch_boundary
                && (direct_void_branch_value || closed_nested_value_in_void_branch)
            {
                body.push(0x1A);
            }
            if op != Op::END {
                pending_externref_on_stack = false;
            }
        }

        if pending_externref_spilled_after_end {
            body.push(0x20); // local.get $temp
            write_leb128_u32(&mut body, temp_local_idx);
        }
        if pending_i32_box_after_end {
            emit_box_i32(&mut body, &rt_idx);
        }

        body.push(0x0B); // end function
        let mut native = body[..locals_prefix_len].to_vec();
        native.extend(super::callable::remap_locals(
            &body[locals_prefix_len..],
            chunk,
            total_temp_locals(chunk),
            &type_ctx.callables,
        ));
        let body = native;
        write_leb128_u32(&mut out, body.len() as u32);
        out.extend_from_slice(&body);
    }
    if type_ctx.native_resume_function_idx.is_some() {
        // One ordinary module-local function, no binding/table slot/host hook.
        // Raw resume is kept separate from GC carrier mutation.
        let mut body = vec![0, 0x02];
        write_leb128_i64(&mut body, type_ctx.continuation_landing_type_idx as i64);
        body.extend_from_slice(&[0x20, 0, 0x20, 1, 0xe3]);
        write_leb128_u32(&mut body, type_ctx.continuation_type_idx);
        write_leb128_u32(&mut body, 1);
        body.push(0);
        write_leb128_u32(&mut body, 1);
        write_leb128_u32(&mut body, 0);
        body.push(0xd0);
        write_leb128_i64(&mut body, type_ctx.continuation_type_idx as i64);
        body.extend_from_slice(&[0x0b, 0x0b]);
        write_leb128_u32(&mut out, body.len() as u32);
        out.extend_from_slice(&body);
    }
    if type_ctx.native_carrier_resume_function_idx.is_some() {
        let mut body=vec![1,1,0x69]; // one exnref local2
        body.extend([0x02,0x69]); // exceptional landing has one exnref result
        body.push(0x1f);write_leb128_i64(&mut body,type_ctx.continuation_landing_type_idx as i64);
        body.extend([1,3,0]); // catch_all_ref to surrounding exception landing
        body.extend([0x20,1,0x20,0]);emit_cont_carrier_field(&mut body,type_ctx,0);
        body.extend([0x20,0]);emit_cont_carrier_field(&mut body,type_ctx,2);
        body.push(0xfb);write_leb128_u32(&mut body,0x16);write_leb128_i64(&mut body,type_ctx.native_resume_type_idx as i64);
        body.push(0x14);write_leb128_u32(&mut body,type_ctx.native_resume_type_idx);
        body.extend([0x0b,0x0f,0x0b,0x21,2]); // normal returns; catch saves exact exnref
        body.extend([0x20,0]);emit_cont_carrier_cast(&mut body,type_ctx);
        body.push(0xd0);write_leb128_i64(&mut body,type_ctx.continuation_type_idx as i64);
        body.push(0xfb);write_leb128_u32(&mut body,5);write_leb128_u32(&mut body,type_ctx.continuation_carrier_type_idx);write_leb128_u32(&mut body,0);
        body.extend([0x20,2,0x0a,0x0b]); // throw_ref, preserving JS/native exception identity
        write_leb128_u32(&mut out,body.len() as u32);out.extend(body);
    }
    for e in &type_ctx.generators {
        let body = super::generator_construct::body(e, chunks, type_ctx, &rt_idx);
        write_leb128_u32(&mut out, body.len() as u32);
        out.extend(body);
    }
    if let Some((target, _, _, results)) = type_ctx.source_entry {
        // Source functions have a native-only capture parameter. A top-level
        // entry has no enclosing frame, so it receives null captures.
        let mut body = vec![0];
        super::callable::null(&mut body, -17);
        body.push(0x10);
        write_leb128_u32(&mut body, target);
        body.extend(std::iter::repeat_n(0x1a, usize::from(results)));
        body.push(0x0b);
        write_leb128_u32(&mut out, body.len() as u32);
        out.extend(body);
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
    all_chunks: &[Chunk],
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    temp_idx: u32,
    i64_temp_idx: u32,
    _has_temp: bool,
    type_ctx: &WasmTypeContext,
    host_import_map: &std::collections::HashMap<(String, String, Option<u8>), usize>,
    tag_plan: &crate::writer::proposals::exception_handling::ModuleTagPlan,
    in_i32_block: bool,
    current_block_results: u8,
    forwarded_types: &[Option<u8>],
    raw_numeric: Option<&RawNumericBody>,
    native_args: &NativeDirectArgs,
    raw_i32_branches: &std::collections::HashSet<usize>,
    declared_calls: &DeclaredCallPlan,
    i32_regions: &SourceI32Regions,
) {
    if let Some(edge) = chunk.direct_defined_calls.get(&op_start) {
        assert!(matches!(op,Op::CALL|Op::RETURN_CALL));
        let callee = all_chunks.get(edge.target as usize).expect("direct defined target");
        assert!(callee.direct_entry_has_captures!=Some(true) && (callee.capture_count==0 || callee.direct_entry_has_captures==Some(false)),"direct capture environment requires explicit adapter");
        assert!(!callee.is_async && !callee.is_generator);
        let signature = callee.func_sig.clone().unwrap_or_else(||(
            vec!["externref";callee.arity as usize].join(","),
            vec!["externref";callee.result_arity as usize].join(",")));
        assert_eq!(signature,(edge.params.clone(),edge.results.clone()));
        assert_eq!(edge.source,callee.func_sig.is_none());
        let (base,types) = declared_calls.slots(op_start).expect("direct typed locals");
        let argc=edge.params.split(',').filter(|s|!s.is_empty()).count();
        *ip=op_start+if op==Op::CALL {7}else{6};
        for i in (0..argc).rev() {
            match types[i] {
                _ if native_args.calls.contains(&op_start) => {},
                TYPE_I32 if !in_i32_block => emit_unbox_i32(body,rt_idx),
                TYPE_F32 => emit_unbox_f32(body,rt_idx),
                TYPE_F64 => emit_unbox_f64(body,rt_idx),
                _=>{},
            }
            body.push(0x21);write_leb128_u32(body,base+i as u32);
        }
        if edge.source {
            // Source native entries take one hidden captures array before raw
            // formals. Capture-free direct entry supplies standard typed null.
            // The actual Source entry formal is externref. The callee
            // converts a non-null environment to its captures GC array.
            super::callable::null(body,-17);
        }
        for i in 0..argc {body.push(0x20);write_leb128_u32(body,base+i as u32);}
        body.push(if op==Op::RETURN_CALL {0x12}else{0x10});
        write_leb128_u32(body,type_ctx.callables.targets[edge.target as usize]);
        if op==Op::CALL {
            let output=&types[argc..];
            // Typed standard callers returning directly retain raw results.
            // Other value-ABI consumers get the normal number boxes.
            let retain_raw= !super::callable::source(chunk) && next_op(chunk,*ip)==Some(Op::RETURN);
            if !retain_raw {
                for (i,_) in output.iter().enumerate().rev() {
                    body.push(0x21);write_leb128_u32(body,base+argc as u32+i as u32);
                }
                for (i,ty) in output.iter().enumerate() {
                    body.push(0x20);write_leb128_u32(body,base+argc as u32+i as u32);
                    match *ty {
                        TYPE_I32 => box_i32_unless_condition(body,rt_idx,chunk,*ip,in_i32_block, raw_i32_branches),
                        TYPE_F32 => emit_box_f32(body,rt_idx),
                        TYPE_F64 => emit_box_f64(body,rt_idx),
                        _=>{},
                    }
                }
            }
        }
        return;
    }
    if raw_numeric.is_some() && raw_scalar_shape(op).is_some() {
        body.push(op.sub() as u8);
        return;
    }
    if let Some(inputs) = i32_regions.inputs(&op_start) {
        match inputs {
            [raw] => {
                if !raw { emit_unbox_i32(body, rt_idx); }
                body.push(op.sub() as u8);
            }
            [true, raw] => {
                if !raw { emit_unbox_i32(body, rt_idx); }
                body.push(op.sub() as u8);
            }
            [false, raw] => {
                // A computed RHS must finish before the left conversion can
                // trap. Rejoin the ordinary boxed binary bridge in this case.
                if *raw { emit_box_i32(body, rt_idx); }
                emit_binary_i32_op(false, body, op.sub() as u8, rt_idx, temp_idx);
            }
            _ => unreachable!("i32 operator arity"),
        }
        if !i32_regions.is_raw(&op_start) {
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
        }
        return;
    }
    match op {
        _ if op == Op::LOCAL_GET => {
            let slot = read_u16(&chunk.code, ip);
            if super::callable::source(chunk)
                && let Some(cell) = super::callable::cell_slot(chunk, &type_ctx.callables, slot)
            {
                super::callable::local(body, 0x20, cell);
                super::callable::internal(body, type_ctx.callables.cell);
                super::callable::field(body, 2, type_ctx.callables.cell, 0);
                if i32_condition_follows(chunk, *ip) || raw_i32_branches.contains(&*ip) {
                    emit_unbox_i32(body, rt_idx);
                }
                return;
            }
            body.push(op.sub() as u8);
            write_leb128_u32(body, slot as u32);
            let numeric_param = boxed_numeric_param_type(chunk, slot, forwarded_types);
            let raw_i32_consumer = (i32_condition_follows(chunk, *ip) || raw_i32_branches.contains(&*ip)) || native_args.producers.get(&op_start) == Some(&TYPE_I32);
            let raw_f32_consumer = native_f32_return_follows(chunk, *ip) || native_args.producers.get(&op_start) == Some(&TYPE_F32);
            let raw_f64_consumer = native_f64_return_follows(chunk, *ip) || native_args.producers.get(&op_start) == Some(&TYPE_F64);
            match numeric_param {
                Some(TYPE_I32) if !raw_i32_consumer => emit_box_i32(body, rt_idx),
                Some(TYPE_F32) if !raw_f32_consumer => emit_box_f32(body, rt_idx),
                Some(TYPE_F64) if !raw_f64_consumer => emit_box_f64(body, rt_idx),
                _ => {}
            }
            if numeric_param.is_none()
                && forwarded_types.get(slot as usize).copied().flatten() != Some(TYPE_I32)
                && raw_i32_consumer
            {
                emit_unbox_i32(body, rt_idx);
            }
            if numeric_param.is_none()
                && forwarded_types.get(slot as usize).copied().flatten() != Some(TYPE_F32)
                && raw_f32_consumer
            {
                emit_unbox_f32(body, rt_idx);
            }
            if numeric_param.is_none()
                && forwarded_types.get(slot as usize).copied().flatten() != Some(TYPE_F64)
                && raw_f64_consumer
            {
                emit_unbox_f64(body, rt_idx);
            }
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
            let slot = read_u16(&chunk.code, ip);
            if super::callable::source(chunk)
                && let Some(cell) = super::callable::cell_slot(chunk, &type_ctx.callables, slot)
            {
                super::callable::local(body, 0x21, slot as u32);
                super::callable::local(body, 0x20, cell);
                super::callable::internal(body, type_ctx.callables.cell);
                super::callable::local(body, 0x20, slot as u32);
                super::callable::field(body, 5, type_ctx.callables.cell, 0);
                return;
            }
            match boxed_numeric_param_type(chunk, slot, forwarded_types) {
                Some(TYPE_I32) => emit_unbox_i32(body, rt_idx),
                Some(TYPE_F32) => emit_unbox_f32(body, rt_idx),
                Some(TYPE_F64) => emit_unbox_f64(body, rt_idx),
                _ => {}
            }
            body.push(0x21);
            write_leb128_u32(body, slot as u32);
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
            let slot = read_u16(&chunk.code, ip) as u32;
            if super::callable::source(chunk)
                && let Some(cell) =
                    super::callable::cell_slot(chunk, &type_ctx.callables, slot as u16)
            {
                super::callable::local(body, 0x22, slot);
                super::callable::local(body, 0x20, cell);
                super::callable::internal(body, type_ctx.callables.cell);
                super::callable::local(body, 0x20, slot);
                super::callable::field(body, 5, type_ctx.callables.cell, 0);
                if i32_condition_follows(chunk, *ip) || raw_i32_branches.contains(&*ip) {
                    emit_unbox_i32(body, rt_idx);
                }
                return;
            }
            if current_block_results == 0
                && let Some(consumed) = dup_drop_before_void_boundary(chunk, *ip, slot as u16)
            {
                // Bytecode `dup; drop` leaves the original value on the VM
                // stack. If the very next structured boundary is void, that
                // carried value is dead and invalid in spec Wasm. Drop the
                // original and consume the duplicate/drop pair.
                body.push(0x1A);
                *ip += consumed;
                return;
            }
            let numeric_param = boxed_numeric_param_type(chunk, slot as u16, forwarded_types);
            match numeric_param {
                Some(TYPE_I32) => emit_unbox_i32(body, rt_idx),
                Some(TYPE_F32) => emit_unbox_f32(body, rt_idx),
                Some(TYPE_F64) => emit_unbox_f64(body, rt_idx),
                _ => {}
            }
            body.push(0x22);
            write_leb128_u32(body, slot);
            match numeric_param {
                Some(TYPE_I32) if !(i32_condition_follows(chunk, *ip) || raw_i32_branches.contains(&*ip)) => emit_box_i32(body, rt_idx),
                Some(TYPE_F32) if !native_f32_return_follows(chunk, *ip) => {
                    emit_box_f32(body, rt_idx)
                }
                Some(TYPE_F64) if !native_f64_return_follows(chunk, *ip) => {
                    emit_box_f64(body, rt_idx)
                }
                _ => {}
            }
            if numeric_param.is_none() && (i32_condition_follows(chunk, *ip) || raw_i32_branches.contains(&*ip)) {
                emit_unbox_i32(body, rt_idx);
            }
            if numeric_param.is_none() && native_f32_return_follows(chunk, *ip) {
                emit_unbox_f32(body, rt_idx);
            }
            if numeric_param.is_none() && native_f64_return_follows(chunk, *ip) {
                emit_unbox_f64(body, rt_idx);
            }
        }
        // Spec `call`: u16 funcidx + VM-internal u8 argc. The argc byte is
        // dropped — the .wasm binary carries only LEB(funcidx). Imports
        // occupy the front of the module's function index space. Chunks carry
        // their own local import tables, so remap by declaration identity.
        _ if op == Op::CALL => {
            let funcidx = read_u16(&chunk.code, ip);
            let argc = chunk.code[*ip];
            *ip += 1;
            let module_funcidx = chunk
                .imports
                .get(funcidx as usize)
                .or_else(|| all_chunks.first()?.imports.get(funcidx as usize))
                .and_then(|import| {
                    host_import_map
                        .get(&(import.module.clone(), import.name.clone(), Some(argc)))
                        .or_else(|| {
                            host_import_map.get(&(import.module.clone(), import.name.clone(), None))
                        })
                        .copied()
                })
                .unwrap_or(funcidx as usize) as u32;
            let argc = emit_missing_import_args(body, rt_idx, type_ctx, module_funcidx, argc);
            emit_unbox_raw_params(body, rt_idx, type_ctx, module_funcidx, argc, temp_idx);
            body.push(0x10);
            write_leb128_u32(body, module_funcidx);
            // ⛔ A TRUTHFUL SIGNATURE HAS TWO HALVES. Declaring
            // `wasm:js-string.test` as `-> i32` is only half of it: its result
            // then sits on a stack whose every other value is externref, and
            // the first consumer that is not an `if` rejects it. Box here, and
            // let the peephole drop the box again when the consumer really
            // does want the raw i32.
            match type_ctx.raw_result_funcs.get(&module_funcidx) {
                Some(&crate::encoding::TYPE_I32) => {
                    box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches)
                }
                // ⛔ f64 RESULTS NEED IT TOO. `wasm:js-number.toF64` really
                // returns f64, and the truthiness ladder stores that straight
                // into a local — a local this writer declares `externref`,
                // like every other. Boxing here is what keeps the ladder's
                // `f64.ne` reachable through `emit_binary_f64_cmp`, which
                // unboxes its operands anyway.
                Some(&crate::encoding::TYPE_F64) if !native_f64_return_follows(chunk, *ip) => {
                    emit_box_f64(body, rt_idx)
                }
                Some(&crate::encoding::TYPE_F32) if !native_f32_return_follows(chunk, *ip) => {
                    body.push(0xBB); // f64.promote_f32
                    emit_box_f64(body, rt_idx);
                }
                _ => {}
            }
            if matches!(type_ctx.func_result_arity.get(&module_funcidx), Some(0)) {
                consume_next_drop(chunk, ip);
            }
        }
        // `call_ref`: one `u8` argc, callee on the stack, lowered to
        // call_indirect through the function table. (The old `Op::CALL`
        // alias for this shape is retired; spec `call` is a static import
        // call — see callimportretirement.md.)
        //
        // `return_call` / `return_call_ref` share the internal shape
        // (u8 argc, callee below the args) and the exact same staging;
        // function-reference forms lower to spec `call_ref` / `return_call_ref`.
        // Only RETURN_CALL keeps the older table-index indirect spelling.
        _ if op == Op::CALL_REF || op == Op::RETURN_CALL || op == Op::RETURN_CALL_REF => {
            let spec_byte: u8 = match op {
                Op::CALL_REF => 0x14,
                Op::RETURN_CALL_REF => 0x15,
                _ => 0x13,
            };
            let is_ref_call = op == Op::CALL_REF || op == Op::RETURN_CALL_REF;
            let argc = chunk.code[*ip];
            *ip += 1;
            let results = chunk.code[*ip];
            *ip += 1;
            if is_ref_call && let Some((base, params)) = declared_calls.slots(op_start) {
                // Numeric bytecode producers use the established boxed-number
                // ABI (except i64 and explicitly raw i32 result regions).
                for (i, ty) in params.iter().enumerate().rev() {
                    match *ty {
                        TYPE_I32 if !in_i32_block => emit_unbox_i32(body, rt_idx),
                        TYPE_F32 => emit_unbox_f32(body, rt_idx),
                        TYPE_F64 => emit_unbox_f64(body, rt_idx),
                        _ => {}
                    }
                    body.push(0x21);
                    write_leb128_u32(body, base + i as u32);
                }
                body.push(0x21);
                write_leb128_u32(body, base + u32::from(argc));
                for i in 0..u32::from(argc) {
                    body.push(0x20);
                    write_leb128_u32(body, base + i);
                }
                body.push(0x20);
                write_leb128_u32(body, base + u32::from(argc));
                let signature = &chunk.call_indirect_sigs[&op_start];
                let type_idx = type_ctx.func_type_by_signature[signature];
                body.push(0xfb);
                write_leb128_u32(body, 0x16);
                write_leb128_i32(body, type_idx as i32);
                body.push(spec_byte);
                write_leb128_u32(body, type_idx);
                if op == Op::CALL_REF {
                    let (_, result) = signature.split_once("->").unwrap();
                    let returns = result
                        .split(',')
                        .map(str::trim)
                        .filter(|s| !s.is_empty())
                        .collect::<Vec<_>>();
                    if returns.len() == 1 && next_op(chunk, *ip) != Some(Op::RETURN) {
                        match returns[0] {
                            "i32" => {
                                box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches)
                            }
                            "f32" => emit_box_f32(body, rt_idx),
                            "f64" => emit_box_f64(body, rt_idx),
                            _ => {}
                        }
                    }
                }
                return;
            }
            // Stack: [externref_funcref, arg1, ..., argN] — funcref is below args
            // call_ref needs: [arg1, ..., argN, funcref].
            // call_indirect needs: [arg1, ..., argN, i32_table_idx].
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
            // A typed core caller can still invoke a boxed, generic helper.
            // Callee signature metadata, not the caller's signature alone,
            // determines whether the reference uses the callable carrier.
            if is_ref_call && super::callable::uses_carrier_call(chunk, op_start) {
                let callee = temp_idx + argc as u32;
                // Caller-owned validation precedes all dispatch and callback invocation.
                super::callable::local(body, 0x20, callee);
                super::callable::internal(body, type_ctx.callables.carrier);
                super::callable::field(body, 2, type_ctx.callables.carrier, 2);
                super::callable::i32c(body, 0);
                body.push(0x47);
                body.push(0x04);
                write_blocktype(body, 0, results, type_ctx);
                super::callable::local(body, 0x20, callee);
                super::callable::internal(body, type_ctx.callables.carrier);
                super::callable::field(body, 2, type_ctx.callables.carrier, 3);
                super::callable::null(body, -17);
                for i in 0..argc {
                    super::callable::local(body, 0x20, temp_idx + i as u32);
                }
                super::callable::typed(body, 8, type_ctx.array_type_idx);
                write_leb128_u32(body, argc as u32);
                super::callable::external(body);
                let apply = rt_idx[&("ecma:function", "apply")];
                if op == Op::RETURN_CALL_REF {
                    body.push(0x12);
                    write_leb128_u32(body, apply as u32);
                } else {
                    emit_import_call(body, rt_idx, "ecma:function", "apply");
                    if results == 0 {
                        body.push(0x1a);
                    }
                }
                body.push(0x05);
                if let Some(&tidx) = chunk
                    .call_indirect_sigs
                    .get(&op_start)
                    .and_then(|sig| type_ctx.func_type_by_signature.get(sig))
                {
                    for i in 0..argc {
                        super::callable::local(body, 0x20, temp_idx + i as u32);
                    }
                    super::callable::local(body, 0x20, callee);
                    super::callable::internal(body, type_ctx.callables.carrier);
                    super::callable::field(body, 2, type_ctx.callables.carrier, 0);
                    super::callable::typed(body, 0x16, tidx);
                    body.push(spec_byte);
                    write_leb128_u32(body, tidx);
                } else {
                    let mut alternatives: Vec<u8> = type_ctx
                        .callables
                        .calls
                        .keys()
                        .filter_map(|&(n, r)| (r == results && n != argc).then_some(n))
                        .collect();
                    alternatives.sort();
                    alternatives.dedup();
                    let mut arities = vec![argc];
                    arities.extend(alternatives);
                    for &n in &arities {
                        super::callable::local(body, 0x20, callee);
                        super::callable::internal(body, type_ctx.callables.carrier);
                        super::callable::field(body, 2, type_ctx.callables.carrier, 4);
                        super::callable::i32c(body, n as i32);
                        body.push(0x46);
                        body.push(0x04);
                        write_blocktype(body, 0, results, type_ctx);
                        super::callable::local(body, 0x20, callee);
                        super::callable::internal(body, type_ctx.callables.carrier);
                        super::callable::field(body, 2, type_ctx.callables.carrier, 1);
                        super::callable::external(body);
                        for i in 0..n {
                            if i < argc {
                                super::callable::local(body, 0x20, temp_idx + i as u32);
                            } else {
                                body.push(0x23);
                                write_leb128_u32(body, JS_GLOBAL_UNDEFINED);
                            }
                        }

                        super::callable::local(body, 0x20, callee);
                        super::callable::internal(body, type_ctx.callables.carrier);
                        super::callable::field(body, 2, type_ctx.callables.carrier, 0);
                        let tidx = type_ctx.callables.calls[&(n, results)];
                        super::callable::typed(body, 0x16, tidx);
                        body.push(spec_byte);
                        write_leb128_u32(body, tidx);
                        body.push(0x05);
                    }
                    // A foreign invalid carrier or undeclared arity cannot masquerade as a guest.
                    body.push(0x00);
                    for _ in arities {
                        body.push(0x0b);
                    }
                }
                body.push(0x0b);
            } else {
                // 3. Restore user args
                for i in 0..argc {
                    body.push(0x20);
                    write_leb128_u32(body, temp_idx + i as u32);
                }
                // 4. Push callee in the spec operand position.
                body.push(0x20);
                write_leb128_u32(body, temp_idx + argc as u32);
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
                    if !is_ref_call {
                        emit_unbox_i32(body, rt_idx);
                    }
                    if is_ref_call {
                        // The value ABI stores function references as boxed table
                        // indices. Recover a typed funcref only at the call boundary;
                        // raw funcrefs cannot be stored in externref spill locals.
                        body.push(0xFB);
                        write_leb128_u32(body, 0x16); // ref.cast (ref $signature)
                        write_leb128_i32(body, type_idx as i32);
                    }
                    body.push(spec_byte);
                    write_leb128_u32(body, type_idx); // type index
                    if !is_ref_call {
                        write_leb128_u32(body, type_ctx.internal_callable_table);
                    }
                } else {
                    // No matching type — drop everything, push null.
                    body.push(0x1A); // drop funcref/table_idx source
                    for _ in 0..argc {
                        body.push(0x1A);
                    }
                    body.push(0xD0);
                    body.push(0x6F);
                }
            }
            if op == Op::CALL_REF
                && results > 0
                && current_block_results == 0
                && matches!(next_op(chunk, *ip), Some(next) if next == Op::ELSE || next == Op::END)
            {
                for _ in 0..results {
                    body.push(0x1A);
                }
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
            if raw_numeric.is_some() {
                let signature = &chunk.call_indirect_sigs[&op_start];
                body.push(spec_byte);
                write_leb128_u32(body, type_ctx.func_type_by_signature[signature]);
                write_leb128_u32(body, table_idx as u32);
                return;
            }
            if let Some((base, types)) = declared_calls.slots(op_start) {
                let argc = usize::from(argc);
                let selector = base + types.len() as u32;
                let address = type_ctx.table_operands.get(table_idx as usize)
                    .map_or(TYPE_I32, |table| table.0);
                if address == TYPE_I32 && !in_i32_block { emit_unbox_i32(body, rt_idx); }
                body.push(0x21);
                write_leb128_u32(body, selector);
                for i in (0..argc).rev() {
                    match types[i] {
                        TYPE_I32 if !in_i32_block => emit_unbox_i32(body, rt_idx),
                        TYPE_F32 => emit_unbox_f32(body, rt_idx),
                        TYPE_F64 => emit_unbox_f64(body, rt_idx),
                        _ => {},
                    }
                    body.push(0x21);
                    write_leb128_u32(body, base + i as u32);
                }
                for i in 0..argc {
                    body.push(0x20);
                    write_leb128_u32(body, base + i as u32);
                }
                body.push(0x20);
                write_leb128_u32(body, selector);
                let signature = &chunk.call_indirect_sigs[&op_start];
                let (_, returns) = signature.split_once("->").expect("declared indirect signature");
                body.push(spec_byte);
                write_leb128_u32(body, type_ctx.func_type_by_signature[signature]);
                write_leb128_u32(body, table_idx as u32);
                let direct_return = next_op(chunk, *ip) == Some(Op::RETURN)
                    && chunk.func_sig.as_ref().is_some_and(|(_, result)| result == returns);
                if op == Op::CALL_INDIRECT && !direct_return {
                    // Stage all results before converting any of them. The
                    // relative order of heterogeneous multi-values is exact.
                    for i in (0..usize::from(results)).rev() {
                        body.push(0x21);
                        write_leb128_u32(body, base + (argc + i) as u32);
                    }
                    for i in 0..usize::from(results) {
                        body.push(0x20);
                        write_leb128_u32(body, base + (argc + i) as u32);
                        match types[argc + i] {
                            TYPE_I32 if results == 1 => box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches),
                            TYPE_I32 => emit_box_i32(body, rt_idx),
                            TYPE_F32 => emit_box_f32(body, rt_idx),
                            TYPE_F64 => emit_box_f64(body, rt_idx),
                            _ => {},
                        }
                    }
                }
                return;
            }
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
            if op == Op::CALL_INDIRECT
                && results > 0
                && current_block_results == 0
                && matches!(next_op(chunk, *ip), Some(next) if next == Op::ELSE || next == Op::END)
            {
                for _ in 0..results {
                    body.push(0x1A);
                }
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
        // `return` is a core control opcode with no immediate. The bytecode
        // stream uses fixed-width opcode headers, but the wasm binary writes
        // only the single spec opcode byte here.
        _ if op == Op::RETURN => {
            body.push(0x0F);
        }
        // END pops a label from the structured CF stack
        _ if op == Op::END => {
            body.push(0x0B); // end
        }
        // BLOCK/LOOP/IF carry (param_count, result_count) bytes. Translate
        // to the spec blocktype: 0x40 void / one valtype / positive s33
        // typeidx into the pre-registered externref^M -> externref^N types.
        _ if op == Op::BLOCK || op == Op::LOOP || op == Op::IF => {
            let opcode_offset = ip.saturating_sub(4);
            let (param_count, result_count) =
                effective_structured_block_counts(chunk, op, opcode_offset);
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
            match (param_count, result_count) {
                (0, 0) => body.push(TYPE_VOID),
                (0, 1) if chunk.block_i32_results.contains(&opcode_offset) => body.push(TYPE_I32),
                (0, 1) => body.push(TYPE_EXTERNREF),
                (params, 1) if chunk.block_i32_results.contains(&opcode_offset) => {
                    let tidx = *type_ctx
                        .block_i32_type_by_params
                        .get(&params)
                        .expect("i32 block functype was not pre-registered");
                    write_leb128_i32(body, tidx as i32);
                }
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
        // Native integer results rejoin the value ABI unless consumed as a
        // control-flow predicate.
        _ if op == Op::I32_WRAP_I64 => {
            body.push(op.sub() as u8);
            if raw_numeric.is_none() {
                box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
            }
        }
        _ if op == Op::MEMORY_SIZE || op == Op::MEMORY_GROW => {
            let memidx = read_leb_u32(&chunk.code, ip);
            let memory64 = chunk
                .memory_is_64
                .get(memidx as usize)
                .copied()
                .unwrap_or(false);
            // Memory64 uses i64 for the delta and result; memory32 crosses
            // this writer's boxed-value boundary through i32 conversions.
            if op == Op::MEMORY_GROW && !memory64 && raw_numeric.is_none() {
                emit_unbox_i32(body, rt_idx);
            }
            body.push(op.sub() as u8);
            write_leb128_u32(body, memidx);
            if !memory64 && raw_numeric.is_none() {
                box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
            }
        }
        // Memory load: the address is the ONLY operand, so it is on top and
        // unboxes in place; the loaded value rejoins the value ABI.
        _ if memory_load_shape(op).is_some() => {
            let (default_align, result) = memory_load_shape(op).unwrap_or((0, 0));
            let (align, offset, memidx) = read_optional_memarg(chunk, ip, default_align);
            let memory64 = chunk
                .memory_is_64
                .get(memidx as usize)
                .copied()
                .unwrap_or(false);
            if !memory64 && raw_numeric.is_none() {
                emit_unbox_i32(body, rt_idx); // externref addr → i32
            }
            body.push(op.sub() as u8);
            encode_memarg_with_memidx(body, align, offset, memidx);
            if result == TYPE_F32
                && raw_numeric.is_none()
                && next_op(chunk, *ip) == Some(Op::I32_REINTERPRET_F32)
            {
                // Reinterpreting loaded bits is not floating arithmetic. Do
                // not promote to Number and quiet a signaling NaN in between.
                body.push(Op::I32_REINTERPRET_F32.sub() as u8);
                *ip += 4;
                box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
                return;
            }
            match result {
                crate::encoding::TYPE_I32 if raw_numeric.is_none() => {
                    box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches)
                }
                crate::encoding::TYPE_F64 if raw_numeric.is_none() => emit_box_f64(body, rt_idx),
                crate::encoding::TYPE_F32 if raw_numeric.is_none() => {
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
            let (align, offset, memidx) = read_optional_memarg(chunk, ip, default_align);
            let memory64 = chunk
                .memory_is_64
                .get(memidx as usize)
                .copied()
                .unwrap_or(false);
            if raw_numeric.is_some() {
                body.push(op.sub() as u8);
                encode_memarg_with_memidx(body, align, offset, memidx);
                return;
            }
            let value_temp = if value == crate::encoding::TYPE_I64 {
                i64_temp_idx
            } else {
                temp_idx
            };
            body.push(0x21); // local.set saved value
            write_leb128_u32(body, value_temp);
            if !memory64 {
                emit_unbox_i32(body, rt_idx); // externref addr → i32
            }
            body.push(0x20); // local.get saved value
            write_leb128_u32(body, value_temp);
            match value {
                crate::encoding::TYPE_I32 => emit_unbox_i32(body, rt_idx),
                crate::encoding::TYPE_F64 => emit_unbox_f64(body, rt_idx),
                crate::encoding::TYPE_F32 => {
                    emit_unbox_f64(body, rt_idx);
                    body.push(0xB6); // f32.demote_f64
                }
                _ => {}
            }
            body.push(op.sub() as u8);
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
            let gidx = read_u32_be(&chunk.code, ip);
            // Preserve actual writable binding identity. Generated RefFunc
            // initializers install unambiguous carriers through real alias cells;
            // names never manufacture numeric table-index callable values.
            body.push(0x23);
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
            let wasm_gidx = read_u32_be(&chunk.code, ip);
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
            let tbl = read_leb_u32(&chunk.code, ip);
            body.push(0x25);
            write_leb128_u32(body, tbl);
        }
        _ if op == Op::TABLE_SET => {
            let tbl = read_leb_u32(&chunk.code, ip);
            body.push(0x26);
            write_leb128_u32(body, tbl);
        }
        // The source value ABI selects externrefs and boxes its i32 condition.
        // Untyped Wasm select only accepts numeric/vector operands. Keep the
        // whole-body proven integer lane raw; bridge the boxed ABI explicitly.
        _ if op == Op::SELECT || op == Op::SELECT_T => {
            if raw_numeric.is_some() {
                body.push(0x1B);
            } else {
                emit_unbox_i32(body, rt_idx);
                body.push(0x1C);
                write_leb128_u32(body, 1);
                body.push(TYPE_EXTERNREF);
            }
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
            if raw_numeric.is_none() {
                box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
            }
        }
        // GC proposal (core prefix): `ref.eq` compares two references in the
        // EQ hierarchy and produces i32.
        //
        // It still cannot blindly cast arbitrary host externrefs to `eqref`.
        // Imported host values are legal values in the VM ABI but they may not
        // live in the GC eq hierarchy after `any.convert_extern`, so the cast
        // would trap before answering false. Take the native `ref.eq` fast
        // path only when both internalised operands test as eqref; otherwise
        // use the existing non-coercing host SameValue operation. Abstract
        // equality is not reference identity and can invoke guest coercions.
        _ if op == Op::REF_EQ => {
            let lhs_temp = temp_idx + 1;
            let rhs_temp = temp_idx;

            body.push(0x21); // local.set $rhs_temp (save b)
            write_leb128_u32(body, rhs_temp);
            body.push(0x21); // local.set $lhs_temp (save a)
            write_leb128_u32(body, lhs_temp);

            body.push(0x20); // local.get $lhs_temp
            write_leb128_u32(body, lhs_temp);
            emit_internalize(body);
            emit_ref_test_eq(body);
            body.push(0x20); // local.get $rhs_temp
            write_leb128_u32(body, rhs_temp);
            emit_internalize(body);
            emit_ref_test_eq(body);
            body.push(0x71); // i32.and

            body.push(0x04); // if
            body.push(TYPE_I32); // result i32
            body.push(0x20); // local.get $lhs_temp
            write_leb128_u32(body, lhs_temp);
            emit_internalize(body);
            emit_ref_cast_eq(body);
            body.push(0x20); // local.get $rhs_temp
            write_leb128_u32(body, rhs_temp);
            emit_internalize(body);
            emit_ref_cast_eq(body);
            body.push(0xD3); // ref.eq
            body.push(0x05); // else
            body.push(0x20); // local.get $lhs_temp
            write_leb128_u32(body, lhs_temp);
            body.push(0x20); // local.get $rhs_temp
            write_leb128_u32(body, rhs_temp);
            emit_import_call(body, rt_idx, "ecma:object", "is");
            body.push(0x0B); // end

            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
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

        // ── floating binary arithmetic: unbox both → native op → rebox ──
        // Stack: [externref_a, externref_b]
        // Pattern: local.set $temp (save b) → toF64 (unbox a) → local.get $temp → toF64 (unbox b)
        //          → f64.op → fromF64 (rebox result)
        _ if op == Op::F32_ADD
            || op == Op::F32_SUB
            || op == Op::F32_MUL
            || op == Op::F32_DIV
            || op == Op::F32_MIN
            || op == Op::F32_MAX
            || op == Op::F32_COPYSIGN
            || op == Op::F64_ADD
            || op == Op::F64_SUB
            || op == Op::F64_MUL
            || op == Op::F64_DIV
            || op == Op::F64_MIN
            || op == Op::F64_MAX
            || op == Op::F64_COPYSIGN =>
        {
            let single = matches!(
                op,
                Op::F32_ADD
                    | Op::F32_SUB
                    | Op::F32_MUL
                    | Op::F32_DIV
                    | Op::F32_MIN
                    | Op::F32_MAX
                    | Op::F32_COPYSIGN
            );
            emit_binary_float_op(
                body,
                op.sub() as u8,
                rt_idx,
                temp_idx,
                single,
                if single {
                    !native_f32_return_follows(chunk, *ip)
                } else {
                    !native_f64_return_follows(chunk, *ip)
                },
            );
        }

        // ── f64 comparisons: unbox both → compare → rebox i32 result ──
        // ⛔ `F64_EQ`/`F64_NE` WERE ABSENT and fell to the catch-all, which
        // emits the bare opcode onto two BOXED operands: "expected f64, found
        // externref". Every comparison that unboxes has to be listed here —
        // an omission is not a missing optimisation, it is invalid wasm.
        _ if op == Op::F32_LT
            || op == Op::F32_GT
            || op == Op::F32_LE
            || op == Op::F32_GE
            || op == Op::F32_EQ
            || op == Op::F32_NE
            || op == Op::F64_LT
            || op == Op::F64_GT
            || op == Op::F64_LE
            || op == Op::F64_GE
            || op == Op::F64_EQ
            || op == Op::F64_NE =>
        {
            emit_binary_float_cmp(
                !(is_i32_block_result(chunk, *ip, in_i32_block)
                    || (i32_condition_follows(chunk, *ip) || raw_i32_branches.contains(&*ip))),
                body,
                op.sub() as u8,
                rt_idx,
                temp_idx,
                matches!(
                    op,
                    Op::F32_LT | Op::F32_GT | Op::F32_LE | Op::F32_GE | Op::F32_EQ | Op::F32_NE
                ),
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
            if raw_numeric.is_some() {
                body.push(op.sub() as u8);
                return;
            }
            emit_binary_i32_op(
                !(is_i32_block_result(chunk, *ip, in_i32_block)
                    || (i32_condition_follows(chunk, *ip) || raw_i32_branches.contains(&*ip))),
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
            if raw_numeric.is_some() {
                body.push(op.sub() as u8);
                return;
            }
            emit_binary_i32_cmp(
                !(is_i32_block_result(chunk, *ip, in_i32_block)
                    || (i32_condition_follows(chunk, *ip) || raw_i32_branches.contains(&*ip))),
                body,
                op.sub() as u8,
                rt_idx,
                temp_idx,
            );
        }

        // ── floating unary ops: unbox → native op → rebox ──
        _ if op == Op::F32_NEG
            || op == Op::F32_ABS
            || op == Op::F32_CEIL
            || op == Op::F32_FLOOR
            || op == Op::F32_TRUNC
            || op == Op::F32_NEAREST
            || op == Op::F32_SQRT
            || op == Op::F64_NEG
            || op == Op::F64_ABS
            || op == Op::F64_CEIL
            || op == Op::F64_FLOOR
            || op == Op::F64_TRUNC
            || op == Op::F64_NEAREST
            || op == Op::F64_SQRT =>
        {
            let single = matches!(
                op,
                Op::F32_NEG
                    | Op::F32_ABS
                    | Op::F32_CEIL
                    | Op::F32_FLOOR
                    | Op::F32_TRUNC
                    | Op::F32_NEAREST
                    | Op::F32_SQRT
            );
            if single {
                emit_unbox_f32(body, rt_idx);
            } else {
                emit_unbox_f64(body, rt_idx);
            }
            body.push(op.sub() as u8);
            if single && !native_f32_return_follows(chunk, *ip) {
                emit_box_f32(body, rt_idx);
            } else if !single && !native_f64_return_follows(chunk, *ip) {
                emit_box_f64(body, rt_idx);
            }
        }

        // ── i32 unary ops ──
        _ if op == Op::I32_EQZ => {
            if raw_numeric.is_some() {
                body.push(op.sub() as u8);
                return;
            }
            emit_unbox_i32(body, rt_idx);
            body.push(op.sub() as u8);
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
        }
        _ if op == Op::I32_CLZ || op == Op::I32_CTZ || op == Op::I32_POPCNT => {
            if raw_numeric.is_some() {
                body.push(op.sub() as u8);
                return;
            }
            emit_unbox_i32(body, rt_idx);
            body.push(op.sub() as u8);
            emit_box_i32(body, rt_idx);
        }

        // Exact adjacent compiler bit-cast pairs retain raw f32 between the
        // standard operations. Promoting/re-demoting in between would alter
        // signaling NaN bit patterns and add arithmetic absent from bytecode.
        _ if op == Op::F32_REINTERPRET_I32
            && !in_i32_block
            && next_op(chunk, *ip) == Some(Op::F32_STORE) =>
        {
            // The ordinary value ABI has [address, boxed i32]. Convert the
            // address before restoring the bits, then perform the two Wasm
            // operations without any floating conversion between them.
            let mut after_store = *ip + 4;
            let (align, offset, memidx) = read_optional_memarg(chunk, &mut after_store, 2);
            body.push(0x21);
            write_leb128_u32(body, temp_idx);
            if !chunk.memory_is_64.get(memidx as usize).copied().unwrap_or(false) {
                emit_unbox_i32(body, rt_idx);
            }
            body.push(0x20);
            write_leb128_u32(body, temp_idx);
            emit_unbox_i32(body, rt_idx);
            body.push(op.sub() as u8);
            body.push(Op::F32_STORE.sub() as u8);
            encode_memarg_with_memidx(body, align, offset, memidx);
            *ip = after_store;
        }
        _ if op == Op::F32_DEMOTE_F64 && next_op(chunk, *ip) == Some(Op::I32_REINTERPRET_F32) => {
            emit_unbox_f64(body, rt_idx);
            body.push(op.sub() as u8);
            body.push(Op::I32_REINTERPRET_F32.sub() as u8);
            *ip += 4; // paired no-immediate opcode header
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
        }
        _ if op == Op::F32_REINTERPRET_I32 && next_op(chunk, *ip) == Some(Op::F64_PROMOTE_F32) => {
            emit_unbox_i32(body, rt_idx);
            body.push(op.sub() as u8);
            body.push(Op::F64_PROMOTE_F32.sub() as u8);
            *ip += 4;
            emit_box_f64(body, rt_idx);
        }
        // Bit transport stays raw across adjacent, side-effect-free casts.
        // JS Number boxing can quiet signaling NaNs before reinterpretation.
        _ if op == Op::F64_CONST && next_op(chunk, *ip + 8) == Some(Op::I64_REINTERPRET_F64) => {
            body.push(op.sub() as u8);
            body.extend_from_slice(&chunk.code[*ip..*ip + 8]);
            body.push(Op::I64_REINTERPRET_F64.sub() as u8);
            *ip += 8 + 4; // f64 bits + paired no-immediate opcode header
        }
        _ if op == Op::F64_REINTERPRET_I64
            && next_op(chunk, *ip) == Some(Op::I64_REINTERPRET_F64) =>
        {
            body.push(op.sub() as u8);
            body.push(Op::I64_REINTERPRET_F64.sub() as u8);
            *ip += 4; // retain raw i64; no JS Number transport between casts
        }
        _ if op == Op::I64_REINTERPRET_F64 => {
            emit_unbox_f64(body, rt_idx);
            body.push(op.sub() as u8); // i64 remains raw
        }
        _ if op == Op::F64_REINTERPRET_I64 => {
            body.push(op.sub() as u8);
            emit_box_f64(body, rt_idx);
        }
        _ if op == Op::I64_EXTEND_I32_S || op == Op::I64_EXTEND_I32_U => {
            if raw_numeric.is_none() {
                emit_unbox_i32(body, rt_idx);
            }
            body.push(op.sub() as u8); // i64 remains raw
        }
        _ if op == Op::I64_TRUNC_F64_S || op == Op::I64_TRUNC_F64_U => {
            emit_unbox_f64(body, rt_idx);
            body.push(op.sub() as u8);
        }
        _ if op == Op::F64_CONVERT_I64_S || op == Op::F64_CONVERT_I64_U => {
            body.push(op.sub() as u8);
            emit_box_f64(body, rt_idx);
        }
        _ if matches!(
            op,
            Op::I64_EQZ
                | Op::I64_EQ
                | Op::I64_NE
                | Op::I64_LT_S
                | Op::I64_LT_U
                | Op::I64_LE_S
                | Op::I64_LE_U
                | Op::I64_GT_S
                | Op::I64_GT_U
                | Op::I64_GE_S
                | Op::I64_GE_U
        ) =>
        {
            body.push(op.sub() as u8);
            if raw_numeric.is_none() {
                box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
            }
        }

        // ── Conversions: unbox source type → convert → rebox target type ──
        _ if op == Op::I32_FROM_F64 || op == Op::I32_TRUNC_F64_U => {
            // Both signed and unsigned truncation consume a raw f64. The
            // unsigned form must not fall through onto the boxed value stack.
            emit_unbox_f64(body, rt_idx);
            body.push(op.sub() as u8);
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
        }
        _ if op == Op::F64_FROM_I32 || op == Op::F64_CONVERT_I32_U => {
            // Signedness changes interpretation of the bits, not the ABI.
            emit_unbox_i32(body, rt_idx);
            body.push(op.sub() as u8);
            emit_box_f64(body, rt_idx);
        }

        // Function values use boxed table indices in the externref value ABI.
        // The element section contains real ref.func values; CALL_REF recovers
        // and casts the table entry to its declared function type.
        _ if op == Op::REF_FUNC => {
            let chunk_idx = read_u16(&chunk.code, ip);
            let uv_raw = chunk.code[*ip];
            let uv_count = (uv_raw & 0x7f) as usize; // mask 0x80 no-intern flag
            *ip += 1;
            let descriptors = *ip;
            *ip += uv_count * 3;
            if !super::callable::source(chunk) {
                body.push(0xd2);
                write_leb128_u32(body, type_ctx.callables.import_count + chunk_idx as u32);
            } else if uv_count == 0 && uv_raw & 0x80 == 0 {
                body.push(0x23);
                write_leb128_u32(body, type_ctx.callables.cache_base + chunk_idx as u32);
                super::callable::external(body);
            } else {
                let meta = *type_ctx
                    .callables
                    .chunks
                    .get(chunk_idx as usize)
                    .expect("REF_FUNC target in module");
                super::callable::carrier(
                    body,
                    &type_ctx.callables,
                    chunk_idx as u32,
                    |b| {
                        if uv_count == 0 {
                            super::callable::null(b, type_ctx.callables.captures as i32);
                        } else {
                            for i in 0..uv_count {
                                let at = descriptors + i * 3;
                                let slot = read_u16_at(&chunk.code, at + 1);
                                if chunk.code[at] != 0 {
                                    let cell = super::callable::cell_slot(
                                        chunk,
                                        &type_ctx.callables,
                                        slot,
                                    )
                                    .expect("captured local cell planned");
                                    super::callable::local(b, 0x20, cell);
                                    super::callable::internal(b, type_ctx.callables.cell);
                                } else {
                                    super::callable::local(
                                        b,
                                        0x20,
                                        super::callable::hidden_virtual(
                                            chunk,
                                            super::callable::temps(chunk, &type_ctx.callables),
                                        ),
                                    );
                                    super::callable::internal(b, type_ctx.callables.captures);
                                    super::callable::i32c(b, slot as i32);
                                    super::callable::typed(b, 11, type_ctx.callables.captures);
                                }
                            }
                            super::callable::typed(b, 8, type_ctx.callables.captures);
                            write_leb128_u32(b, uv_count as u32);
                        }
                    },
                    meta,
                );
                super::callable::external(body);
            }
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
            // Synchronous Source generators use logical yield tag zero. The
            // native module reserves tag zero for exceptions and tag one for
            // its (externref)->externref suspension protocol. Reader-authored
            // functions retain their declared tag index unchanged.
            let native_tag = if tag_idx == 0
                && !super::generator_construct::entries(std::slice::from_ref(chunk)).is_empty()
            {
                1
            } else {
                u32::from(tag_idx)
            };
            write_leb128_u32(body, native_tag);
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
            // A literal RHS has no evaluation effects and is already an i32.
            // Unbox the stacked value first, then emit the literal directly:
            // avoid boxing the constant, spilling it, and unboxing it again.
            // Source entries use the boxed-value stack; declared raw Wasm
            // bodies and direct-call argument producers retain their own ABI.
            let after_literal = *ip + sz;
            if op == Op::I32_CONST
                && raw_numeric.is_none()
                && super::callable::source(chunk)
                && !native_args.producers.contains_key(&op_start)
                && let Some(binary) = next_op(chunk, after_literal)
                && binary.group() == 0
                && matches!(binary.sub(), 0x46..=0x4f | 0x6a..=0x78)
            {
                if !i32_regions.inputs(&after_literal).is_some_and(|inputs| inputs.first() == Some(&true)) {
                    emit_unbox_i32(body, rt_idx);
                }
                body.push(0x41);
                body.extend_from_slice(&chunk.code[*ip..after_literal]);
                body.push(binary.sub() as u8);
                *ip = after_literal + 4;
                if !i32_regions.is_raw(&after_literal) {
                    box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
                }
                return;
            }
            body.push(op.sub() as u8);
            body.extend_from_slice(&chunk.code[*ip..*ip + sz]);
            *ip += sz;
            if op == Op::F64_CONST && raw_numeric.is_none() && !native_f64_return_follows(chunk, *ip) && !native_args.producers.contains_key(&op_start) {
                emit_box_f64(body, rt_idx);
            } else if op == Op::I32_CONST && raw_numeric.is_none() && !native_args.producers.contains_key(&op_start) && !i32_regions.is_raw(&op_start) {
                box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
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
            // The value ABI boxes f32 through f64, except when a typed f32
            // return consumes the raw constant directly.
            if raw_numeric.is_none() && !native_f32_return_follows(chunk, *ip) && !native_args.producers.contains_key(&op_start) {
                emit_box_f32(body, rt_idx);
            }
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

fn atomic_is_i32_store(op: Op) -> bool {
    op == Op::I32_ATOMIC_STORE || op == Op::I32_ATOMIC_STORE8 || op == Op::I32_ATOMIC_STORE16
}

fn atomic_is_i32_rmw(op: Op) -> bool {
    op == Op::I32_ATOMIC_RMW_ADD
        || op == Op::I32_ATOMIC_RMW8_ADD_U
        || op == Op::I32_ATOMIC_RMW16_ADD_U
        || op == Op::I32_ATOMIC_RMW_SUB
        || op == Op::I32_ATOMIC_RMW8_SUB_U
        || op == Op::I32_ATOMIC_RMW16_SUB_U
        || op == Op::I32_ATOMIC_RMW_AND
        || op == Op::I32_ATOMIC_RMW8_AND_U
        || op == Op::I32_ATOMIC_RMW16_AND_U
        || op == Op::I32_ATOMIC_RMW_OR
        || op == Op::I32_ATOMIC_RMW8_OR_U
        || op == Op::I32_ATOMIC_RMW16_OR_U
        || op == Op::I32_ATOMIC_RMW_XOR
        || op == Op::I32_ATOMIC_RMW8_XOR_U
        || op == Op::I32_ATOMIC_RMW16_XOR_U
        || op == Op::I32_ATOMIC_RMW_XCHG
        || op == Op::I32_ATOMIC_RMW8_XCHG_U
        || op == Op::I32_ATOMIC_RMW16_XCHG_U
}

fn atomic_is_i32_two_operand(op: Op) -> bool {
    atomic_is_i32_store(op) || atomic_is_i32_rmw(op)
}

fn atomic_is_i32_cmpxchg(op: Op) -> bool {
    op == Op::I32_ATOMIC_RMW_CMPXCHG
        || op == Op::I32_ATOMIC_RMW8_CMPXCHG_U
        || op == Op::I32_ATOMIC_RMW16_CMPXCHG_U
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
/// The wait32 timeout stays raw i64; only its address and expected value use
/// the source value ABI. No BigInt or host conversion is needed for that lane.
fn emit_thread_prefixed_op(
    body: &mut Vec<u8>,
    op: Op,
    chunk: &Chunk,
    ip: &mut usize,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    temp_idx: u32,
    i64_temp_idx: u32,
    in_i32_block: bool,
    raw_numeric: bool,
    raw_i32_branches: &std::collections::HashSet<usize>,
) {
    if op == Op::ATOMIC_FENCE {
        body.push(0xFE);
        write_leb128_u32(body, op.sub() as u32);
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
    let write_op = |body: &mut Vec<u8>| {
        body.push(0xFE);
        write_leb128_u32(body, op.sub() as u32);
        write_leb128_u32(body, spec_align);
        if is_memory64 {
            write_leb128_u64(body, offset);
        } else {
            write_leb128_u32(body, offset as u32);
        }
        if let Some(memidx) = memidx {
            write_leb128_u32(body, memidx);
        }
    };

    if op == Op::MEMORY_ATOMIC_WAIT32 && super::callable::source(chunk) {
        super::callable::local(body, 0x21, i64_temp_idx);
        super::callable::local(body, 0x21, temp_idx);
        if !is_memory64 {
            emit_unbox_i32(body, rt_idx);
        }
        super::callable::local(body, 0x20, temp_idx);
        emit_unbox_i32(body, rt_idx);
        super::callable::local(body, 0x20, i64_temp_idx);
        write_op(body);
        box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
        return;
    }

    if atomic_is_i32_address_only_load(op) {
        emit_unbox_i32(body, rt_idx); // externref addr → i32
        write_op(body);
        // i32 result → back onto the externref value ABI, unless the consumer
        // is a raw-i32 region. `box_i32_unless_condition`, not a bare box: an
        // atomic feeding an `if` directly must stay raw, exactly as for every
        // other i32 producer in this writer.
        box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
        return;
    }

    // General bodies carry notify operands through the value ABI, as for
    // atomic stores/RMWs. Proven typed integer bodies retain raw operands.
    let boxed_notify = op == Op::MEMORY_ATOMIC_NOTIFY && !raw_numeric;
    if atomic_is_i32_two_operand(op) || boxed_notify {
        body.push(0x21); // local.set $temp (save value)
        write_leb128_u32(body, temp_idx);
        if !boxed_notify || !is_memory64 {
            emit_unbox_i32(body, rt_idx); // externref addr → i32
        } // memory64 notify addresses already use the raw i64 lane.
        body.push(0x20); // local.get $temp (restore value)
        write_leb128_u32(body, temp_idx);
        emit_unbox_i32(body, rt_idx); // externref value → i32
        write_op(body);
        if atomic_is_i32_rmw(op) || boxed_notify {
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
        }
        return;
    }

    if atomic_is_i32_cmpxchg(op) {
        body.push(0x21); // local.set $temp+1 (replacement)
        write_leb128_u32(body, temp_idx + 1);
        body.push(0x21); // local.set $temp (expected)
        write_leb128_u32(body, temp_idx);
        emit_unbox_i32(body, rt_idx); // externref addr → i32
        body.push(0x20); // local.get $temp (expected)
        write_leb128_u32(body, temp_idx);
        emit_unbox_i32(body, rt_idx);
        body.push(0x20); // local.get $temp+1 (replacement)
        write_leb128_u32(body, temp_idx + 1);
        emit_unbox_i32(body, rt_idx);
        write_op(body);
        box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
        return;
    }

    write_op(body);
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
    let (align, offset, memory) = read_optional_memarg(chunk, ip, default_align);
    encode_memarg_with_memidx(body, align, offset, memory);
}

fn default_simd_align(op: Op) -> u32 {
    op.natural_align_bytes().expect("SIMD memory opcode has a natural alignment").trailing_zeros()
}

// ── Binary operation helpers ─────────────────────────────────────────

/// Emit a floating binary op across the boxed-value boundary.
/// Uses temp local to save b while unboxing a.
fn emit_binary_float_op(
    body: &mut Vec<u8>,
    wasm_opcode: u8,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    temp_idx: u32,
    single: bool,
    box_result: bool,
) {
    // Save b to temp local
    body.push(0x21); // local.set
    write_leb128_u32(body, temp_idx);
    if single {
        emit_unbox_f32(body, rt_idx);
    } else {
        emit_unbox_f64(body, rt_idx);
    }
    // Restore b from temp local
    body.push(0x20); // local.get
    write_leb128_u32(body, temp_idx);
    if single {
        emit_unbox_f32(body, rt_idx);
    } else {
        emit_unbox_f64(body, rt_idx);
    }
    body.push(wasm_opcode);
    if box_result {
        if single {
            emit_box_f32(body, rt_idx);
        } else {
            emit_box_f64(body, rt_idx);
        }
    }
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
    matches!(next_op(chunk, ip), Some(next) if next == Op::IF || next == Op::BR_IF || next == Op::BR_TABLE)
        || native_i32_return_follows(chunk, ip)
}

// Preserve native result lanes through a straight-line suffix of pure
// constants and typed parameter reads. Source bodies and intervening work
// retain the value ABI; no general stack-type inference is claimed here.
fn native_result_lane_follows(chunk: &Chunk, mut ip: usize, current: &str) -> bool {
    let Some((params, results)) = &chunk.func_sig else { return false; };
    let results: Vec<_> = results.split(',').map(str::trim).filter(|s| !s.is_empty()).collect();
    if results.is_empty() { return false; }
    let params: Vec<_> = params.split(',').map(str::trim).filter(|s| !s.is_empty()).collect();
    let mut suffix = Vec::new();
    loop {
        let Some(op) = next_op(chunk, ip) else { return false; };
        if op == Op::RETURN {
            let Some(lane) = results.len().checked_sub(suffix.len() + 1) else { return false; };
            return results[lane] == current && results[lane + 1..] == suffix;
        }
        if suffix.len() + 1 >= results.len() { return false; }
        let operand = ip + 4;
        let ty = match op {
            Op::I32_CONST => "i32",
            Op::I64_CONST => "i64",
            Op::F32_CONST => "f32",
            Op::F64_CONST => "f64",
            Op::LOCAL_GET => {
                if operand + 2 > chunk.code.len() { return false; }
                let slot = u16::from_be_bytes([chunk.code[operand], chunk.code[operand+1]]) as usize;
                let Some(ty) = params.get(slot) else { return false; };
                *ty
            }
            Op::NULL => match chunk.code.get(operand) {
                Some(0x6f) => "externref",
                Some(0x70) => "funcref",
                _ => return false,
            },
            _ => return false,
        };
        suffix.push(ty);
        ip = operand + op.operand_format().size_in(&chunk.code, operand);
    }
}
fn native_i32_return_follows(chunk: &Chunk, ip: usize) -> bool {
    native_result_lane_follows(chunk, ip, "i32")
}
fn native_f32_return_follows(chunk: &Chunk, ip: usize) -> bool {
    native_result_lane_follows(chunk, ip, "f32")
}
fn native_f64_return_follows(chunk: &Chunk, ip: usize) -> bool {
    native_result_lane_follows(chunk, ip, "f64")
}

/// Offsets of unconditional branches whose label consumes one raw i32.
/// Block/if labels use their results; loop labels use parameters, so an i32
/// loop result alone cannot authorize a raw branch payload. Count try_table
/// labels too, since they participate in standard relative branch depths.
/// This is collected once per body, rather than rescanning for each producer.
fn i32_branch_consumers(chunk: &Chunk) -> std::collections::HashSet<usize> {
    let mut branches = std::collections::HashSet::new();
    let function_i32 = chunk.func_sig.as_ref()
        .is_some_and(|(_, results)| results.trim() == "i32");
    if chunk.block_i32_results.is_empty() && !function_i32 {
        return branches;
    }
    // The function body also has an implicit label with its result signature.
    let mut controls = vec![function_i32];
    let mut ip = 0;
    while ip + 3 < chunk.code.len() {
        let Some(op) = Op::decode(
            u16::from_be_bytes([chunk.code[ip], chunk.code[ip + 1]]),
            u16::from_be_bytes([chunk.code[ip + 2], chunk.code[ip + 3]]),
        ) else {
            return std::collections::HashSet::new();
        };
        if matches!(op, Op::BLOCK | Op::IF | Op::LOOP | Op::TRY_TABLE) {
            controls.push(matches!(op, Op::BLOCK | Op::IF)
                && chunk.block_i32_results.contains(&ip));
        } else if op == Op::END {
            controls.pop();
        } else if op == Op::BR {
            let mut operand = ip + 4;
            let depth = read_leb_u32(&chunk.code, &mut operand) as usize;
            if controls.len().checked_sub(depth + 1)
                .and_then(|target| controls.get(target)).copied().unwrap_or(false)
            {
                branches.insert(ip);
            }
        }
        ip += opcode_size(op, &chunk.code, ip);
    }
    branches
}

/// True when the NEXT bytecode instruction can consume a raw i32 operand at
/// the bytecode boundary.
///
/// Keep this deliberately narrow. Internal `I32_*` arithmetic/comparison ops
/// are emitted by helpers that first unbox their operands from the externref
/// value ABI; treating them as raw consumers leaves plain `i32.const` values on
/// the stack and the helper immediately spills them into externref temps. Spec
/// branches are consumers that really do take the raw i32 directly. `drop` is
/// the other legal boundary consumer: Wasm `drop` is type-polymorphic, and
/// boxing a raw host result only so the next instruction can discard it is both
/// wasteful and, for call sites that already emitted an explicit drop, can
/// leave the real drop with nothing to consume.
/// A return declared as native i32 is also a raw consumer.
fn i32_raw_consumer_follows(chunk: &Chunk, ip: usize) -> bool {
    matches!(
        next_op(chunk, ip),
        Some(next) if next == Op::IF || next == Op::BR_IF || next == Op::BR_TABLE || next == Op::DROP
    ) || native_i32_return_follows(chunk, ip)
}

fn void_boundary_value_producer(op: Op) -> bool {
    matches!(
        op,
        Op::LOCAL_GET
            | Op::GLOBAL_GET
            | Op::I32_CONST
            | Op::I64_CONST
            | Op::F32_CONST
            | Op::F64_CONST
            | Op::NULL
            | Op::REF_IS_NULL
            | Op::REF_EQ
            | Op::SELECT
            | Op::SELECT_T
            | Op::STRUCT_GET
            | Op::ARRAY_GET
            | Op::ARRAY_LENGTH
            | Op::I32_EQZ
            | Op::I64_EQZ
            | Op::I32_ADD
            | Op::I32_SUB
            | Op::I32_MUL
            | Op::I32_DIV_S
            | Op::I32_DIV_U
            | Op::I32_REM_S
            | Op::I32_REM_U
            | Op::I32_AND
            | Op::I32_OR
            | Op::I32_XOR
            | Op::I32_SHL
            | Op::I32_SHR_S
            | Op::I32_SHR_U
            | Op::I32_ROTL
            | Op::I32_ROTR
            | Op::I32_EQ
            | Op::I32_NE
            | Op::I32_LT_S
            | Op::I32_LT_U
            | Op::I32_GT_S
            | Op::I32_GT_U
            | Op::I32_LE_S
            | Op::I32_LE_U
            | Op::I32_GE_S
            | Op::I32_GE_U
            | Op::I64_ADD
            | Op::I64_SUB
            | Op::I64_MUL
            | Op::I64_AND
            | Op::I64_OR
            | Op::I64_XOR
            | Op::I64_SHL
            | Op::I64_SHR_S
            | Op::I64_SHR_U
            | Op::I64_ROTL
            | Op::I64_ROTR
            | Op::I64_EQ
            | Op::I64_NE
            | Op::I64_LT_S
            | Op::I64_LT_U
            | Op::I64_GT_S
            | Op::I64_GT_U
            | Op::I64_LE_S
            | Op::I64_LE_U
            | Op::I64_GE_S
            | Op::I64_GE_U
            | Op::F32_ADD
            | Op::F32_SUB
            | Op::F32_MUL
            | Op::F32_DIV
            | Op::F32_EQ
            | Op::F32_NE
            | Op::F32_LT
            | Op::F32_GT
            | Op::F32_LE
            | Op::F32_GE
            | Op::F64_ADD
            | Op::F64_SUB
            | Op::F64_MUL
            | Op::F64_DIV
            | Op::F64_EQ
            | Op::F64_NE
            | Op::F64_LT
            | Op::F64_GT
            | Op::F64_LE
            | Op::F64_GE
    )
}

fn bundled_helper_chunk_for_global(name: &str) -> Option<&'static str> {
    match name {
        "__vybe_encoding_base64_decode" => Some("__stdlib_encoding_base64_decode"),
        "__vybe_encoding_bin2hex" => Some("__stdlib_encoding_bin2hex"),
        "__vybe_csv_parse_line" => Some("__stdlib_csv_parse_line"),
        "__vybe_csv_parse_document" => Some("__stdlib_csv_parse_document"),
        "__vybe_csv_format_minimal" => Some("__stdlib_csv_format_minimal"),
        "__vybe_csv_format_whitespace" => Some("__stdlib_csv_format_whitespace"),
        "__vybe_bytes_to_binary_string" => Some("__stdlib_bytes_to_binary_string"),
        "__vybe_bytes_to_binary_string_all" => Some("__stdlib_bytes_to_binary_string_all"),
        "__vybe_binary_string_to_bytes" => Some("__stdlib_binary_string_to_bytes"),
        "__vybe_encoding_hex2bin" => Some("__stdlib_encoding_hex2bin"),
        "__vybe_encoding_str_rot13" => Some("__stdlib_encoding_str_rot13"),
        "__vybe_encoding_quoted_printable_encode" => {
            Some("__stdlib_encoding_quoted_printable_encode")
        }
        "__vybe_encoding_quoted_printable_decode" => {
            Some("__stdlib_encoding_quoted_printable_decode")
        }
        "__vybe_encoding_convert_uuencode" => Some("__stdlib_encoding_convert_uuencode"),
        "__vybe_encoding_convert_uudecode" => Some("__stdlib_encoding_convert_uudecode"),
        "__vybe_encoding_base64_decode_lenient" => Some("__stdlib_encoding_base64_decode_lenient"),
        "__vybe_sort_by_key" => Some("__stdlib_sort_by_key"),
        "__vybe_sort_by_key_receiver" => Some("__stdlib_sort_by_key_receiver"),
        "__vybe_heapify" => Some("__stdlib_heapify"),
        "__vybe_nsmallest" => Some("__stdlib_nsmallest"),
        "__vybe_nlargest" => Some("__stdlib_nlargest"),
        "__vybe_nsmallest_key" => Some("__stdlib_nsmallest_key"),
        "__vybe_nlargest_key" => Some("__stdlib_nlargest_key"),
        "__vybe_nsmallest_key_receiver" => Some("__stdlib_nsmallest_key_receiver"),
        "__vybe_nlargest_key_receiver" => Some("__stdlib_nlargest_key_receiver"),
        "__vybe_autoderef" => Some("__stdlib_autoderef"),
        "__vybe_dyneq" => Some("__stdlib_dyneq"),
        "__vybe_dynlt" => Some("__stdlib_dynlt"),
        "__vybe_dyngt" => Some("__stdlib_dyngt"),
        "__vybe_dynle" => Some("__stdlib_dynle"),
        "__vybe_dynge" => Some("__stdlib_dynge"),
        _ => None,
    }
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

pub(crate) fn effective_structured_block_counts(
    chunk: &Chunk,
    op: Op,
    op_start: usize,
) -> (u8, u8) {
    let params = chunk.code.get(op_start + 4).copied().unwrap_or(0);
    let results = chunk.code.get(op_start + 5).copied().unwrap_or(0);
    if op == Op::IF && params == 0 && results == 0 {
        if matches!(next_op(chunk, op_start + 6), Some(next) if next == Op::DROP) {
            return (1, 1);
        }
        if if_has_immediate_value_arms(chunk, op_start) {
            return (0, 1);
        }
    }
    (params, results)
}

fn if_has_immediate_value_arms(chunk: &Chunk, op_start: usize) -> bool {
    let then_ip = op_start + 6;
    let Some(then_op) = next_op(chunk, then_ip) else {
        return false;
    };
    if !void_boundary_value_producer(then_op) {
        return false;
    }
    let else_ip = then_ip + opcode_size(then_op, &chunk.code, then_ip);
    if !matches!(next_op(chunk, else_ip), Some(next) if next == Op::ELSE) {
        return false;
    }
    let else_body_ip = else_ip + opcode_size(Op::ELSE, &chunk.code, else_ip);
    let Some(else_op) = next_op(chunk, else_body_ip) else {
        return false;
    };
    if !void_boundary_value_producer(else_op) {
        return false;
    }
    let end_ip = else_body_ip + opcode_size(else_op, &chunk.code, else_body_ip);
    matches!(next_op(chunk, end_ip), Some(next) if next == Op::END)
}

fn consume_next_drop(chunk: &Chunk, ip: &mut usize) -> bool {
    if matches!(next_op(chunk, *ip), Some(next) if next == Op::DROP) {
        *ip += opcode_size(Op::DROP, &chunk.code, *ip);
        true
    } else {
        false
    }
}

fn consume_next_drop_before_void_boundary(chunk: &Chunk, ip: &mut usize) -> bool {
    if !matches!(next_op(chunk, *ip), Some(next) if next == Op::DROP) {
        return false;
    }
    let after_drop = *ip + opcode_size(Op::DROP, &chunk.code, *ip);
    if matches!(next_op(chunk, after_drop), Some(next) if next == Op::END || next == Op::ELSE) {
        *ip = after_drop;
        true
    } else {
        false
    }
}

fn dup_drop_before_void_boundary(chunk: &Chunk, ip: usize, slot: u16) -> Option<usize> {
    if !matches!(next_op(chunk, ip), Some(next) if next == Op::LOCAL_GET) {
        return None;
    }
    if read_u16_at(&chunk.code, ip + 4) != slot {
        return None;
    }
    let after_get = ip + opcode_size(Op::LOCAL_GET, &chunk.code, ip);
    if !matches!(next_op(chunk, after_get), Some(next) if next == Op::DROP) {
        return None;
    }
    let after_drop = after_get + opcode_size(Op::DROP, &chunk.code, after_get);
    if matches!(next_op(chunk, after_drop), Some(next) if next == Op::END || next == Op::ELSE) {
        Some(after_drop - ip)
    } else {
        None
    }
}

fn key_registration_drop_before_void_boundary(
    chunk: &Chunk,
    ip: usize,
    slot: u16,
) -> Option<usize> {
    if !matches!(next_op(chunk, ip), Some(next) if next == Op::LOCAL_GET) {
        return None;
    }
    if read_u16_at(&chunk.code, ip + 4) != slot {
        return None;
    }
    let after_get = ip + opcode_size(Op::LOCAL_GET, &chunk.code, ip);
    if !matches!(next_op(chunk, after_get), Some(next) if next == Op::STRUCT_GET) {
        return None;
    }
    let after_struct_get = after_get + opcode_size(Op::STRUCT_GET, &chunk.code, after_get);
    if !matches!(next_op(chunk, after_struct_get), Some(next) if next == Op::GLOBAL_GET) {
        return None;
    }
    let after_global_get =
        after_struct_get + opcode_size(Op::GLOBAL_GET, &chunk.code, after_struct_get);
    if !matches!(next_op(chunk, after_global_get), Some(next) if next == Op::CALL) {
        return None;
    }
    let after_call = after_global_get + opcode_size(Op::CALL, &chunk.code, after_global_get);
    if !matches!(next_op(chunk, after_call), Some(next) if next == Op::DROP) {
        return None;
    }
    let after_drop = after_call + opcode_size(Op::DROP, &chunk.code, after_call);
    if matches!(next_op(chunk, after_drop), Some(next) if next == Op::END || next == Op::ELSE) {
        Some(after_call)
    } else {
        None
    }
}

fn carried_local_tee_struct_sets(chunk: &Chunk) -> std::collections::HashSet<usize> {
    let mut carried = std::collections::HashSet::new();
    let mut ip = 0;
    let mut candidate: Option<(u16, bool)> = None;
    while ip + 3 < chunk.code.len() {
        let Some(op) = Op::decode(
            ((chunk.code[ip] as u16) << 8) | chunk.code[ip + 1] as u16,
            ((chunk.code[ip + 2] as u16) << 8) | chunk.code[ip + 3] as u16,
        ) else {
            ip += 4;
            continue;
        };
        if op == Op::LOCAL_TEE {
            candidate = Some((read_u16_at(&chunk.code, ip + 4), false));
        } else if op == Op::LOCAL_GET {
            let slot = read_u16_at(&chunk.code, ip + 4);
            if let Some((candidate_slot, seen_get)) = &mut candidate
                && *candidate_slot == slot
            {
                *seen_get = true;
            }
        } else if op == Op::STRUCT_SET {
            if matches!(candidate, Some((_, true))) {
                carried.insert(ip);
            }
        } else if matches!(
            op,
            Op::DROP | Op::LOCAL_SET | Op::END | Op::ELSE | Op::BR | Op::BR_IF | Op::RETURN
        ) {
            candidate = None;
        }
        ip += opcode_size(op, &chunk.code, ip);
    }
    carried
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
    in_i32_block && matches!(next_op(chunk, ip), Some(next) if next == Op::ELSE || next == Op::END)
}

/// Box an i32 to the externref value ABI unless the very next instruction is
/// going to want it raw. See `i32_condition_follows`.
fn box_i32_unless_condition(
    body: &mut Vec<u8>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    chunk: &Chunk,
    ip: usize,
    in_i32_block: bool,
    raw_i32_branches: &std::collections::HashSet<usize>,
) {
    // ⛔ AN i32-RESULT BLOCK IS A RAW REGION. `emit_if_i32`/`emit_block_i32`
    // are the compiler saying "this block's result is a bare i32, not a
    // value" — that is what the whole truthiness ladder is built from, and
    // every arm of it ends in an i32 producer. Boxing those made each arm
    // yield an externref out of a block declared `(result i32)`.
    if is_i32_block_result(chunk, ip, in_i32_block)
        || i32_raw_consumer_follows(chunk, ip)
        || raw_i32_branches.contains(&ip)
    {
        return;
    }
    emit_box_i32(body, rt_idx);
}

fn emit_missing_import_args(
    body: &mut Vec<u8>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    type_ctx: &WasmTypeContext,
    funcidx: u32,
    argc: u8,
) -> u8 {
    let Some(params) = type_ctx.import_param_funcs.get(&funcidx) else {
        return argc;
    };
    let declared = params.len();
    let supplied = argc as usize;
    if supplied >= declared {
        return argc;
    }
    for &param in &params[supplied..] {
        match param {
            crate::encoding::TYPE_I32 => {
                body.push(0x41);
                write_leb128_i32(body, 0);
                emit_box_i32(body, rt_idx);
            }
            crate::encoding::TYPE_I64 => {
                body.push(0x42);
                write_leb128_i64(body, 0);
            }
            crate::encoding::TYPE_F32 => {
                body.push(0x43);
                body.extend_from_slice(&0f32.to_le_bytes());
                body.push(0xBB); // f64.promote_f32
                emit_box_f64(body, rt_idx);
            }
            crate::encoding::TYPE_F64 => {
                body.push(0x44);
                body.extend_from_slice(&0f64.to_le_bytes());
                emit_box_f64(body, rt_idx);
            }
            crate::encoding::TYPE_EXTERNREF => {
                body.push(0x23);
                write_leb128_u32(body, JS_GLOBAL_UNDEFINED);
            }
            _ => {
                body.push(0xD0);
                body.push(crate::encoding::HT_EXTERN);
            }
        }
    }
    declared as u8
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
/// Static calls reserve spill locals for every parameter above the deepest
/// raw one. Variadic imports such as `splice` may require more than two.
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
    let Some(deepest_raw) = params
        .iter()
        .position(|&t| t != crate::encoding::TYPE_EXTERNREF)
    else {
        return;
    };
    let spill = params.len() - 1 - deepest_raw;
    let unbox = |body: &mut Vec<u8>, vt: u8, index: usize| {
        if let Some(&(nullable, heap)) = type_ctx.import_ref_params.get(&(funcidx, index)) {
            assert_eq!(
                heap, type_ctx.string_array_type_idx as i64,
                "only declared packed-i16 import refs are supported"
            );
            emit_internalize(body);
            body.push(0xFB);
            write_leb128_u32(body, if nullable { 0x17 } else { 0x16 });
            write_leb128_i64(body, heap);
            return;
        }
        match vt {
            crate::encoding::TYPE_I32 => emit_unbox_i32(body, rt_idx),
            crate::encoding::TYPE_F32 => {
                emit_unbox_f64(body, rt_idx);
                body.push(0xB6); // f32.demote_f64
            }
            crate::encoding::TYPE_F64 => emit_unbox_f64(body, rt_idx),
            _ => {}
        }
    };
    // Spill the arguments sitting above the deepest raw one; temp_idx + i
    // holds param n-1-i.
    for i in 0..spill {
        body.push(0x21); // local.set
        write_leb128_u32(body, temp_idx + i as u32);
    }
    unbox(body, params[deepest_raw], deepest_raw);
    for i in (0..spill).rev() {
        body.push(0x20); // local.get
        write_leb128_u32(body, temp_idx + i as u32);
        unbox(body, params[params.len() - 1 - i], params.len() - 1 - i);
    }
}

/// Emit a floating comparison with either a raw i32 or boxed result.
fn emit_binary_float_cmp(
    box_result: bool,
    body: &mut Vec<u8>,
    wasm_opcode: u8,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    temp_idx: u32,
    single: bool,
) {
    body.push(0x21);
    write_leb128_u32(body, temp_idx); // local.set $temp (save b)
    if single {
        emit_unbox_f32(body, rt_idx);
    } else {
        emit_unbox_f64(body, rt_idx);
    }
    body.push(0x20);
    write_leb128_u32(body, temp_idx); // local.get $temp (restore b)
    if single {
        emit_unbox_f32(body, rt_idx);
    } else {
        emit_unbox_f64(body, rt_idx);
    }
    body.push(wasm_opcode);
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
    op_start: usize,
    carried_struct_sets: &std::collections::HashSet<usize>,
    rt_idx: &std::collections::HashMap<(&str, &str), usize>,
    type_ctx: &WasmTypeContext,
    temp_idx: u32,
    in_i32_block: bool,
    raw_i32_branches: &std::collections::HashSet<usize>,
    value_string_input: bool,
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
            let named_type = read_leb_u32(&chunk.code, ip);
            let field_name_idx = read_leb_u32(&chunk.code, ip);
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
                    .unwrap_or_else(|| wasm_struct_field_for_name(chunk, type_ctx, field_name_idx));
            emit_internalize(body); // externref → anyref
            emit_ref_cast(body, typeidx); // anyref → (ref $struct)
            body.push(0xFB);
            write_leb128_u32(body, 0x02); // struct.get
            write_leb128_u32(body, typeidx);
            write_leb128_u32(body, fieldidx);
            // Result is externref (field type) — no conversion needed
        }
        _ if op == Op::STRUCT_SET => {
            let named_type = read_leb_u32(&chunk.code, ip);
            let field_name_idx = read_leb_u32(&chunk.code, ip);
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
                // `[[Set]]` answers a Boolean (§10.1.9), but Vybe's
                // `STRUCT_SET` bytecode has spec `struct.set` shape: it pushes
                // nothing. Discard only the host flag introduced by this
                // lowering. A following bytecode `DROP` is a real instruction
                // for some value already on the VM stack (for example a
                // carried `local.tee` receiver) and must be preserved.
                body.push(0x1A); // drop
                if !carried_struct_sets.contains(&op_start) {
                    consume_next_drop_before_void_boundary(chunk, ip);
                }
                return;
            }
            let (typeidx, fieldidx) =
                wasm_struct_field_by_typeidx(chunk, type_ctx, named_type, field_name_idx)
                    .unwrap_or_else(|| wasm_struct_field_for_name(chunk, type_ctx, field_name_idx));
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
            if !carried_struct_sets.contains(&op_start) {
                consume_next_drop_before_void_boundary(chunk, ip);
            }
        }
        _ if op == Op::ARRAY_NEW_FIXED => {
            // Spec: `array.new_fixed $t N` (0xFB 0x08), pops N values.
            // Our bytecode now carries both immediates. A 0 type index means a
            // dynamic-language array literal, which maps to the module's
            // generic `(array (mut externref))`; a stamped index names its own
            // array type.
            let typeidx = read_u16(&chunk.code, ip);
            let elem_count = read_u16(&chunk.code, ip);
            if let Some(&packed) = type_ctx.packed_i16_types.get(&typeidx) {
                for i in 0..elem_count {
                    body.push(0x21);
                    write_leb128_u32(body, temp_idx + i as u32);
                }
                for i in (0..elem_count).rev() {
                    body.push(0x20);
                    write_leb128_u32(body, temp_idx + i as u32);
                    emit_unbox_i32(body, rt_idx);
                }
                body.push(0xFB);
                write_leb128_u32(body, 0x08);
                write_leb128_u32(body, packed);
                write_leb128_u32(body, elem_count as u32);
                emit_externalize(body);
            } else {
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
        }
        _ if op == Op::ARRAY_NEW => {
            // Spec: `array.new $t` (0xFB 0x06), pops [value, length i32].
            // Our bytecode emits the 2-byte type-index immediate like the
            // fixed variant; consume it, drop the type index, and pass
            // through to the engine (which will fill len copies of value).
            let typeidx = read_u16(&chunk.code, ip);
            if let Some(&packed) = type_ctx.packed_i16_types.get(&typeidx) {
                body.push(0x21);
                write_leb128_u32(body, temp_idx);
                emit_unbox_i32(body, rt_idx);
                body.push(0x20);
                write_leb128_u32(body, temp_idx);
                emit_unbox_i32(body, rt_idx);
                body.push(0xFB);
                write_leb128_u32(body, 0x06);
                write_leb128_u32(body, packed);
                emit_externalize(body);
            } else {
                // Value arrays keep the initializer boxed, but the length is
                // still a raw Wasm i32, just as for array.new_default.
                emit_unbox_i32(body, rt_idx);
                body.push(0xFB);
                write_leb128_u32(body, 0x06);
                write_leb128_u32(body, type_ctx.array_type_idx);
                emit_externalize(body);
            }
        }
        _ if op == Op::ARRAY_NEW_DEFAULT => {
            // Spec: `array.new_default $t` (0xFB 0x07), pops [length].
            let typeidx = read_u16(&chunk.code, ip);
            emit_unbox_i32(body, rt_idx);
            body.push(0xFB);
            write_leb128_u32(body, 0x07);
            write_leb128_u32(
                body,
                type_ctx
                    .packed_i16_types
                    .get(&typeidx)
                    .copied()
                    .unwrap_or(type_ctx.array_type_idx),
            );
            emit_externalize(body);
        }
        _ if op == Op::ARRAY_NEW_DATA => {
            // Spec: `array.new_data $t $d`, pops [offset, size].
            let typeidx = read_u16(&chunk.code, ip);
            let data_idx = read_u16(&chunk.code, ip);
            if let Some(&packed) = type_ctx.packed_i16_types.get(&typeidx) {
                body.push(0x21);
                write_leb128_u32(body, temp_idx);
                emit_unbox_i32(body, rt_idx);
                body.push(0x20);
                write_leb128_u32(body, temp_idx);
                emit_unbox_i32(body, rt_idx);
                body.push(0xFB);
                write_leb128_u32(body, 0x09);
                write_leb128_u32(body, packed);
                write_leb128_u32(body, data_idx as u32);
                emit_externalize(body);
            } else {
                body.push(0xFB);
                write_leb128_u32(body, 0x09);
                write_leb128_u32(body, type_ctx.array_type_idx);
                write_leb128_u32(body, data_idx as u32);
                emit_externalize(body);
            }
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
            let typeidx = read_u16(&chunk.code, ip);
            if let Some(&packed) = type_ctx.packed_i16_types.get(&typeidx) {
                body.push(0x21);
                write_leb128_u32(body, temp_idx);
                emit_internalize(body);
                emit_ref_cast_array(body, packed);
                body.push(0x20);
                write_leb128_u32(body, temp_idx);
                emit_unbox_i32(body, rt_idx);
                body.push(0xFB);
                write_leb128_u32(body, op.sub() as u32);
                write_leb128_u32(body, packed);
                box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
            } else {
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
            let _typeidx = read_leb_u32(&chunk.code, ip);
            let field_name_idx = read_leb_u32(&chunk.code, ip);
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
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
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
            let &array_get_idx = rt_idx
                .get(&("ecma:array", "getValue"))
                .expect("ecma:array.getValue runtime import is always collected");
            // Stack: [externref_arr, externref_idx]. `ARRAY_GET` is the VM's
            // shared collection read: GC arrays must lower to spec `array.get`,
            // while dynamic/ECMA arrays stay on the adapter surface.
            body.push(0x21);
            write_leb128_u32(body, temp_idx); // local.set $temp (save idx)
            body.push(0x22);
            write_leb128_u32(body, temp_idx + 1); // local.tee $temp2 (save arr, leave arr)
            emit_internalize(body);
            emit_ref_test_array(body, type_ctx.array_type_idx);
            body.push(0x04); // if
            body.push(TYPE_EXTERNREF); // result externref
            {
                body.push(0x20);
                write_leb128_u32(body, temp_idx + 1); // local.get arr
                emit_internalize(body);
                emit_ref_cast_array(body, type_ctx.array_type_idx);
                body.push(0x20);
                write_leb128_u32(body, temp_idx); // local.get idx
                emit_unbox_i32(body, rt_idx);
                body.push(0xFB);
                write_leb128_u32(body, 0x0B); // array.get
                write_leb128_u32(body, type_ctx.array_type_idx);
            }
            body.push(0x05); // else
            {
                if !type_ctx.packed_i16_types.is_empty() {
                body.push(0x20);
                write_leb128_u32(body, temp_idx + 1);
                emit_internalize(body);
                emit_ref_test_array(body, type_ctx.string_array_type_idx);
                body.push(0x04);
                body.push(TYPE_EXTERNREF);
                body.push(0x20);
                write_leb128_u32(body, temp_idx + 1);
                emit_internalize(body);
                emit_ref_cast_array(body, type_ctx.string_array_type_idx);
                body.push(0x20);
                write_leb128_u32(body, temp_idx + 0);
                emit_unbox_i32(body, rt_idx);
                body.push(0xFB);
                write_leb128_u32(body, 13);
                write_leb128_u32(body, type_ctx.string_array_type_idx);
                emit_box_i32(body, rt_idx);
                body.push(0x05);
                }
                body.push(0x20);
                write_leb128_u32(body, temp_idx + 1); // local.get arr
                body.push(0x20);
                write_leb128_u32(body, temp_idx); // local.get idx
                body.push(0x10); // call ecma:array.getValue(arr, key)
                write_leb128_u32(body, array_get_idx as u32);
                if !type_ctx.packed_i16_types.is_empty() { body.push(0x0B); }
            }
            body.push(0x0B); // end
        }
        _ if op == Op::ARRAY_SET => {
            let &array_set_idx = rt_idx
                .get(&("ecma:array", "setValue"))
                .expect("ecma:array.setValue runtime import is always collected");
            // Stack: [externref_arr, externref_idx, externref_val].
            // Branch exactly as ARRAY_GET: typed GC arrays use `array.set`,
            // dynamic arrays use the ECMA array adapter.
            body.push(0x21);
            write_leb128_u32(body, temp_idx); // local.set $temp (save val)
            body.push(0x21);
            write_leb128_u32(body, temp_idx + 1); // local.set $temp2 (save idx)
            body.push(0x22);
            write_leb128_u32(body, temp_idx + 2); // local.tee $temp3 (save arr, leave arr)
            emit_internalize(body);
            emit_ref_test_array(body, type_ctx.array_type_idx);
            body.push(0x04); // if
            body.push(0x40); // empty block type
            {
                body.push(0x20);
                write_leb128_u32(body, temp_idx + 2); // local.get arr
                emit_internalize(body);
                emit_ref_cast_array(body, type_ctx.array_type_idx);
                body.push(0x20);
                write_leb128_u32(body, temp_idx + 1); // local.get idx
                emit_unbox_i32(body, rt_idx);
                body.push(0x20);
                write_leb128_u32(body, temp_idx); // local.get val
                body.push(0xFB);
                write_leb128_u32(body, 0x0E); // array.set
                write_leb128_u32(body, type_ctx.array_type_idx);
            }
            body.push(0x05); // else
            {
                if !type_ctx.packed_i16_types.is_empty() {
                body.push(0x20);
                write_leb128_u32(body, temp_idx + 2);
                emit_internalize(body);
                emit_ref_test_array(body, type_ctx.string_array_type_idx);
                body.push(0x04);
                body.push(0x40);
                body.push(0x20);
                write_leb128_u32(body, temp_idx + 2);
                emit_internalize(body);
                emit_ref_cast_array(body, type_ctx.string_array_type_idx);
                body.push(0x20);
                write_leb128_u32(body, temp_idx + 1);
                emit_unbox_i32(body, rt_idx);
                body.push(0x20);
                write_leb128_u32(body, temp_idx);
                emit_unbox_i32(body, rt_idx);
                body.push(0xFB);
                write_leb128_u32(body, 14);
                write_leb128_u32(body, type_ctx.string_array_type_idx);
                body.push(0x05);
                }
                body.push(0x20);
                write_leb128_u32(body, temp_idx + 2); // local.get arr
                body.push(0x20);
                write_leb128_u32(body, temp_idx + 1); // local.get key
                body.push(0x20);
                write_leb128_u32(body, temp_idx); // local.get val
                body.push(0x10); // call ecma:array.setValue(arr, key, val)
                write_leb128_u32(body, array_set_idx as u32);
                if !type_ctx.packed_i16_types.is_empty() { body.push(0x0B); }
            }
            body.push(0x0B); // end
        }
        _ if op == Op::ARRAY_LENGTH => {
            let &array_len_idx = rt_idx
                .get(&("ecma:array", "length"))
                .expect("ecma:array.length runtime import is always collected");
            body.push(0x22);
            write_leb128_u32(body, temp_idx); // local.tee arr
            emit_internalize(body);
            emit_ref_test_array(body, type_ctx.array_type_idx);
            body.push(0x04); // if
            body.push(TYPE_I32); // result i32
            {
                body.push(0x20);
                write_leb128_u32(body, temp_idx); // local.get arr
                emit_internalize(body);
                emit_ref_cast_array(body, type_ctx.array_type_idx);
                body.push(0xFB);
                write_leb128_u32(body, 0x0F); // array.len
            }
            body.push(0x05); // else
            {
                if !type_ctx.packed_i16_types.is_empty() {
                body.push(0x20);
                write_leb128_u32(body, temp_idx + 0);
                emit_internalize(body);
                emit_ref_test_array(body, type_ctx.string_array_type_idx);
                body.push(0x04);
                body.push(TYPE_I32);
                body.push(0x20);
                write_leb128_u32(body, temp_idx + 0);
                emit_internalize(body);
                emit_ref_cast_array(body, type_ctx.string_array_type_idx);
                body.push(0xFB);
                write_leb128_u32(body, 15);
                body.push(0x05);
                }
                body.push(0x20);
                write_leb128_u32(body, temp_idx); // local.get arr
                body.push(0x10); // call ecma:array.length(arr)
                write_leb128_u32(body, array_len_idx as u32);
                if !type_ctx.packed_i16_types.is_empty() { body.push(0x0B); }
            }
            body.push(0x0B); // end
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
        }
        _ if op == Op::ARRAY_FILL => {
            emit_internalize(body);
            body.push(0xFB);
            write_leb128_u32(body, 0x10);
            write_leb128_u32(body, type_ctx.array_type_idx);
        }
        _ if op == Op::ARRAY_COPY => {
            // The value ABI boxes all five operands. Restore the spec stack
            // [destination reference, i32 offset, source reference, i32 offset,
            // i32 count] after spilling, including casts to the array type.
            for offset in (0..5).rev() {
                body.push(0x21);
                write_leb128_u32(body, temp_idx + offset);
            }
            for offset in 0..5 {
                body.push(0x20);
                write_leb128_u32(body, temp_idx + offset);
                if offset == 0 || offset == 2 {
                    emit_internalize(body);
                    emit_ref_cast_array(body, type_ctx.array_type_idx);
                } else {
                    emit_unbox_i32(body, rt_idx);
                }
            }
            body.push(0xFB);
            write_leb128_u32(body, 0x11);
            write_leb128_u32(body, type_ctx.array_type_idx);
            write_leb128_u32(body, type_ctx.array_type_idx);
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
                box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
            } else if internal {
                emit_externalize(body); // the cast result rejoins the value ABI
            }
        }
        _ if op == Op::REF_TEST || op == Op::REF_CAST => {
            // `ref.test` / `ref.cast` (non-null): `0xFB {0x14|0x16} <heaptype>`.
            let ht_bytes = read_heaptype_operand(chunk, ip, type_ctx);
            if op == Op::REF_TEST && super::callable::source(chunk) && ht_bytes == [0x73] {
                body.push(0x1a);
                super::callable::i32c(body, 0);
                box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
                return;
            }
            if op == Op::REF_TEST
                && super::callable::source(chunk)
                && matches!(ht_bytes.as_slice(), [0x70] | [0x6b])
            {
                let func = ht_bytes == [0x70];
                super::callable::local(body, 0x21, temp_idx);
                super::callable::local(body, 0x20, temp_idx);
                super::callable::gc(body, 0x1a);
                super::callable::typed(body, 0x14, type_ctx.callables.carrier);
                body.extend([0x04, TYPE_I32]);
                super::callable::local(body, 0x20, temp_idx);
                super::callable::internal(body, type_ctx.callables.carrier);
                super::callable::field(body, 2, type_ctx.callables.carrier, 2);
                super::callable::i32c(body, 2);
                body.push(if func { 0x49 } else { 0x46 });
                body.push(0x05);
                if func {
                    super::callable::i32c(body, 0);
                } else {
                    super::callable::local(body, 0x20, temp_idx);
                    super::callable::gc(body, 0x1a);
                    body.extend([0xfb, 0x14, 0x6b]);
                }
                body.push(0x0b);
                box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
                return;
            }
            let internal = heaptype_is_internal(&ht_bytes);
            if internal {
                emit_internalize(body); // externref → anyref
            }
            body.push(0xFB);
            write_leb128_u32(body, op.sub() as u32);
            body.extend_from_slice(&ht_bytes);
            if op == Op::REF_TEST {
                box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
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
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
        }
        _ if op == Op::STRING_MEASURE_WTF8 && value_string_input => {
            // The Source value ABI carries JS strings as externrefs. The
            // stringref proposal consumes a real string reference and returns
            // a native i32; bridge both boundaries explicitly.
            emit_internalize(body);
            body.extend([0xfb, 0x16, 0x67]); // ref.cast string
            emit_stringref_gc_op(body, op);
            box_i32_unless_condition(body, rt_idx, chunk, *ip, in_i32_block, raw_i32_branches);
        }
        _ if emit_stringref_gc_op(body, op) => {}
        _ => {
            panic!(
                "unsupported 0xFB opcode in WASM writer: {:#x} {:#x}",
                op.group(),
                op.sub()
            );
        }
    }
}

fn emit_stringref_gc_op(body: &mut Vec<u8>, op: Op) -> bool {
    let sub = op.sub() as u32;
    let is_mem_op = matches!(
        sub,
        0x80 | 0x81 | 0x86 | 0x87 | 0x8B | 0x8C | 0x8D | 0x8E | 0x92 | 0x9B
    );
    let is_no_immediate_op = matches!(
        sub,
        0x83..=0x8A | 0x90 | 0x91 | 0x93 | 0x98..=0x9C | 0xA0..=0xA4 | 0xB0..=0xB7
    );

    if is_mem_op {
        body.push(0xFB);
        write_leb128_u32(body, sub);
        write_leb128_u32(body, 0);
        return true;
    }

    if is_no_immediate_op {
        body.push(0xFB);
        write_leb128_u32(body, sub);
        return true;
    }

    if sub == 0x82 {
        panic!("string.const requires a stringref literal section; writer has no internal operand");
    }

    false
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
            return encode_concrete_heaptype(idx);
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
                Some(wasm_idx) => write_exact_heaptype(&mut buf, wasm_idx),
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
        None => (spelled, HeapType::from_spec_reftype_name(spelled).is_some()),
    };
    // A user type wins over an abstract spelling: a module may legitimately
    // declare a type called `any`, and the type table is the authority on
    // what this module actually emitted.
    if let Some(idx) = type_ctx.struct_type(bare) {
        return (encode_concrete_heaptype(idx), nullable);
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
                Some(wasm_idx) => write_concrete_heaptype(&mut buf, wasm_idx),
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

/// Emit `ref.test eq` — a non-trapping check before the `ref.eq` fast path.
fn emit_ref_test_eq(body: &mut Vec<u8>) {
    body.push(0xFB);
    write_leb128_u32(body, 0x15);
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
    // Heaptype immediates are signed LEBs. Writing a concrete typeidx as
    // unsigned LEB makes values such as 98 serialize as one byte `0x62`, which
    // is the Custom Descriptors exact-heaptype marker rather than type 98.
    write_concrete_heaptype(body, type_idx);
}

/// Emit `ref.cast (ref null $arr_typeidx)` for array refs
fn emit_ref_cast_array(body: &mut Vec<u8>, arr_type_idx: u32) {
    emit_ref_cast(body, arr_type_idx);
}

/// Emit `ref.test` for the writer's canonical GC array type. Callers pass an
/// internalized reference on the stack and get an i32 condition.
fn emit_ref_test_array(body: &mut Vec<u8>, arr_type_idx: u32) {
    body.push(0xFB);
    write_leb128_u32(body, 0x15);
    write_concrete_heaptype(body, arr_type_idx);
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ref_cast_heaptype_uses_signed_leb_for_concrete_typeidx() {
        let mut body = Vec::new();
        emit_ref_cast(&mut body, 98);

        assert_eq!(
            body,
            vec![0xFB, 0x17, 0xE2, 0x00],
            "concrete heaptype 98 must be signed LEB, not unsigned 0x62"
        );
    }
}

pub(crate) fn total_temp_locals(c: &Chunk) -> u32 {
    let (x, i) = count_temp_locals(c);
    x + i + u32::from(has_native_next(c))
}

#[cfg(test)]
mod marker_scan_tests {
    use super::*;
    use vybe_runtime::chunk::{LocalName, LocalOrigin};
    #[test]
    fn marker_free_and_source_markers_have_no_native_regions() {
        let mut c = Chunk::new("marker-free");
        c.emit_op(Op::NOP, 1);
        assert!(native_direct_resumes(&c).is_empty());
        assert!(native_next_regions(&c).is_empty());
        c.local_names.push(LocalName::new(0, "<native-generator-resume:bad>", LocalOrigin::Source));
        c.local_names.push(LocalName::new(1, "<native-generator-next:bad>", LocalOrigin::Source));
        assert!(native_direct_resumes(&c).is_empty());
        assert!(native_next_regions(&c).is_empty());
    }
    #[test]
    fn compiler_direct_resume_keeps_coordinates() {
        let mut c = Chunk::new("resume");
        c.emit_op_u16(Op::RESUME, 0, 1);
        c.local_names.push(LocalName::new(0, "<native-generator-resume:0:6>", LocalOrigin::Compiler));
        assert_eq!(native_direct_resumes(&c), vec![0]);
    }
    #[test]
    fn malformed_compiler_markers_still_fail_validation() {
        for next in [false, true] {
            let mut c = Chunk::new("invalid");
            c.emit_op(Op::NOP, 1);
            let name = if next { "<native-generator-next:bad>" } else { "<native-generator-resume:bad>" };
            c.local_names.push(LocalName::new(0, name, LocalOrigin::Compiler));
            assert!(std::panic::catch_unwind(|| { if next { let _ = native_next_regions(&c); } else { let _ = native_direct_resumes(&c); } }).is_err());
        }
    }
}
