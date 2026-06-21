//! # exception-handling proposal
//!
//! Spec: `proposals/spec/proposals/exception-handling/`.
//! Adds `try_table`, `throw`, `throw_ref`, tags, and catch handlers.
//! The proposal has been through several revisions — our emitter
//! targets the **"new EH"** (`try_table`) variant accepted in 2023.
//!
//! ## Status in Vybe
//!
//! | Feature                | Status | Notes |
//! |------------------------|--------|-------|
//! | `try_start` / `try_end` (VM-internal variant) | ✅ | Used by compilers emitting structured try/except |
//! | `try_table`            | ✅ | Emitted by wrapping TRY_START/TRY_END regions in structural blocks with `catch_all_ref` |
//! | `throw`                | ✅ | Emitted as `throw $vybe_exception` — VM uses its handler stack |
//! | `throw_ref`            | ✅ | Shares the single `$vybe_exception` tag (externref param) |
//! | Tag section            | ✅ | One tag `$vybe_exception (param externref)` declared |
//!
//! All exceptions carry a single externref payload (the exception
//! object). Type-based catch dispatch happens inside the catch
//! handler's bytecode (matching on `__exception_type` field), not at
//! the WASM tag level — this lets us share one tag across every
//! language without a combinatorial type explosion.

use super::encoding::*;
use crate::Chunk;

pub const IMPORTS: &[(&str, &str)] = &[];
pub const GLOBAL_IMPORTS: &[(&str, &str)] = &[];
pub fn custom_sections(_chunks: &[Chunk]) -> Vec<(&'static str, Vec<u8>)> {
    Vec::new()
}

/// Tag index of the single `$vybe_exception` tag in the emitted module.
/// Every `throw` in the emitted .wasm references this tag.
pub const VYBE_EXCEPTION_TAG: u32 = 0;

/// Encode the tag section (section id 13). Always declares the
/// `$vybe_exception (param externref)` tag for the exception-handling
/// proposal, and — if `suspend_tag_type_idx` is non-zero — an
/// additional `$vybe_suspend (param externref) (result externref)` tag
/// for the stack-switching proposal.
pub fn encode_tag_section(exception_type_idx: u32) -> Vec<u8> {
    encode_tag_section_with(exception_type_idx, None)
}

/// Full-control tag section encoder. Pass `Some(typeidx)` for the
/// suspend/resume tag to declare it alongside the exception tag.
pub fn encode_tag_section_with(
    exception_type_idx: u32,
    suspend_tag_type_idx: Option<u32>,
) -> Vec<u8> {
    encode_tag_section_with_continuation_tags(exception_type_idx, suspend_tag_type_idx, &[])
}

pub fn encode_tag_section_with_continuation_tags(
    exception_type_idx: u32,
    suspend_tag_type_idx: Option<u32>,
    continuation_tag_type_indices: &[u32],
) -> Vec<u8> {
    let mut out = Vec::new();
    let count = 1u32
        + u32::from(suspend_tag_type_idx.is_some())
        + continuation_tag_type_indices.len() as u32;
    write_leb128_u32(&mut out, count);
    out.push(0x00);
    write_leb128_u32(&mut out, exception_type_idx);
    if let Some(idx) = suspend_tag_type_idx {
        out.push(0x00);
        write_leb128_u32(&mut out, idx);
    }
    for idx in continuation_tag_type_indices {
        out.push(0x00);
        write_leb128_u32(&mut out, *idx);
    }
    out
}

/// Does this chunk (or any chunk in the module) actually need the
/// exception-handling infrastructure? If no compiler emitted any
/// `throw` or try region, we can skip the tag section entirely.
pub fn module_uses_exceptions(chunks: &[Chunk]) -> bool {
    for chunk in chunks {
        let mut ip = 0;
        while ip + 1 < chunk.code.len() {
            if let Some(op) = crate::opcode::Op::decode(chunk.code[ip], chunk.code[ip + 1] as u16) {
                if op == crate::opcode::Op::THROW
                    || op == crate::opcode::Op::THROW_REF
                    || op == crate::opcode::Op::TRY_START
                {
                    return true;
                }
                ip += super::code::opcode_size(op, &chunk.code, ip);
            } else {
                ip += 2;
            }
        }
    }
    false
}
