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

use crate::encoding::*;
use vybe_runtime::Chunk;

pub const IMPORTS: &[(&str, &str)] = &[];
pub const GLOBAL_IMPORTS: &[(&str, &str)] = &[];
pub fn custom_sections(_chunks: &[Chunk]) -> Vec<(&'static str, Vec<u8>)> {
    Vec::new()
}

/// Tag index of the shared `$vybe_exception` tag in the emitted module. It is
/// deliberately index 0: the stack-switching tags follow it at 1.., and every
/// language-level `throw` (which imports the chunk tag `vybe:exception`) maps
/// onto it.
pub const VYBE_EXCEPTION_TAG: u32 = 0;

/// The chunk tag name every language front end imports for its own exceptions.
/// Must match `vybe_compiler::primitives::errors::EXCEPTION_TAG_NAME`.
const VYBE_EXCEPTION_TAG_NAME: &str = "vybe:exception";

/// One tag in the emitted module's tag index space.
pub struct ModuleTag {
    pub debug_name: String,
    pub arity: u8,
}

/// How a module's chunk-level tags land in the ONE tag index space a `.wasm`
/// module has.
///
/// A tag's identity is the whole matching rule of `try_table`, and it was being
/// thrown away: every `catch`/`catch_ref` clause and every `throw` was written
/// with tagidx 0, so a module with several distinct tags serialized into one
/// where they were all the same tag — clauses that must not match would match.
///
/// Chunks are the module's functions, so their tag lists are merged: an
/// IMPORTED tag resolves by name to one shared entity (that is what makes a
/// `throw $t` in one function reachable by a `catch $t` in another), while each
/// LOCAL declaration is a fresh entity and gets its own slot.
pub struct ModuleTagPlan {
    /// Tags to declare, in module tag-index order, starting at `reserved`.
    pub extra: Vec<ModuleTag>,
    /// chunk index → chunk tag index → module tag index.
    pub chunk_maps: Vec<Vec<u32>>,
}

impl ModuleTagPlan {
    /// The module tag index for chunk `ci`'s tag `tag_idx`. Falls back to the
    /// shared exception tag for a chunk that declared none — a hand-built chunk
    /// whose `throw` predates the tag section.
    pub fn module_tag(&self, ci: usize, tag_idx: u16) -> u32 {
        self.chunk_maps
            .get(ci)
            .and_then(|m| m.get(tag_idx as usize))
            .copied()
            .unwrap_or(VYBE_EXCEPTION_TAG)
    }
}

/// Build the plan. `reserved` is how many tag indices the fixed prefix already
/// claims (the exception tag, then the stack-switching suspend/continuation
/// tags), so newly declared tags start after them.
pub fn plan_module_tags(chunks: &[Chunk], reserved: u32) -> ModuleTagPlan {
    let mut plan = ModuleTagPlan {
        extra: Vec::new(),
        chunk_maps: Vec::with_capacity(chunks.len()),
    };
    let mut by_import_name: std::collections::HashMap<String, u32> =
        std::collections::HashMap::new();
    for chunk in chunks {
        let mut map = Vec::with_capacity(chunk.tags.len());
        for tag in &chunk.tags {
            // The one tag every front end shares already has a slot.
            if tag.debug_name == VYBE_EXCEPTION_TAG_NAME {
                map.push(VYBE_EXCEPTION_TAG);
                continue;
            }
            let idx = if tag.imported {
                match by_import_name.get(&tag.debug_name) {
                    Some(&i) => i,
                    None => {
                        let i = reserved + plan.extra.len() as u32;
                        by_import_name.insert(tag.debug_name.clone(), i);
                        plan.extra.push(ModuleTag {
                            debug_name: tag.debug_name.clone(),
                            arity: tag.arity,
                        });
                        i
                    }
                }
            } else {
                // A local declaration is a FRESH entity — never shared, not
                // even with a same-named one in another function.
                let i = reserved + plan.extra.len() as u32;
                plan.extra.push(ModuleTag {
                    debug_name: tag.debug_name.clone(),
                    arity: tag.arity,
                });
                i
            };
            map.push(idx);
        }
        plan.chunk_maps.push(map);
    }
    plan
}

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
    encode_tag_section_full(
        exception_type_idx,
        suspend_tag_type_idx,
        continuation_tag_type_indices,
        &[],
    )
}

/// The whole tag section. `extra_type_indices` declares the module's OWN tags
/// — the ones a `(tag $e (param …))` in the source produced — after the fixed
/// prefix, one entry per [`ModuleTagPlan::extra`] entry and in the same order.
pub fn encode_tag_section_full(
    exception_type_idx: u32,
    suspend_tag_type_idx: Option<u32>,
    continuation_tag_type_indices: &[u32],
    extra_type_indices: &[u32],
) -> Vec<u8> {
    let mut out = Vec::new();
    let count = 1u32
        + u32::from(suspend_tag_type_idx.is_some())
        + continuation_tag_type_indices.len() as u32
        + extra_type_indices.len() as u32;
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
    for idx in extra_type_indices {
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
        while ip + 3 < chunk.code.len() {
            let g = ((chunk.code[ip] as u16) << 8) | chunk.code[ip + 1] as u16;
            let s = ((chunk.code[ip + 2] as u16) << 8) | chunk.code[ip + 3] as u16;
            if let Some(op) = vybe_runtime::opcode::Op::decode(g, s) {
                if op == vybe_runtime::opcode::Op::THROW
                    || op == vybe_runtime::opcode::Op::THROW_REF
                    || op == vybe_runtime::opcode::Op::TRY_TABLE
                {
                    return true;
                }
                ip += crate::writer::code::opcode_size(op, &chunk.code, ip);
            } else {
                ip += 4;
            }
        }
    }
    false
}
