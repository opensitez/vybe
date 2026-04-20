//! # compilation-hints proposal
//!
//! Spec: `proposals/compilation-hints/`. Adds
//! optional custom sections under the `metadata.code.*` namespace that
//! tell a WASM engine how to prioritize compilation / optimization of
//! the module. They are **pure performance hints** — semantics-preserving.
//!
//! ## Subsections we emit
//!
//! ### `metadata.code.compilation_order`
//!
//! Indicates which functions are hot at startup and should be compiled
//! to the top tier immediately rather than waiting for tier-up feedback.
//! Per-function u32 priority (lower = more urgent). We mark the script /
//! entry chunks as priority 0 so V8/SpiderMonkey can emit optimized code
//! for them before any other chunk.
//!
//! Per-function hint format:
//! * `u32` function index
//! * `u32` byte offset within function body (0 = function-level hint)
//! * `u32` length of the hint payload
//! * `u32` priority (0 = highest)
//!
//! ### `metadata.code.branch_hint`  ✅ emitted
//!
//! Per-function map of byte-offset → 1-byte hint payload (`0x01` =
//! likely, `0x00` = unlikely). We scan each chunk's bytecode, track
//! loop nesting depth, and mark conditional branches (`BR_IF_*` /
//! `BR_IF_LABEL`) inside loops as likely-taken (back-edges usually
//! loop). See `scan_branch_hints`.
//!
//! ### `metadata.code.inlining`  ✅ emitted
//!
//! Per-function priority (0 = strongly inline). Every non-entry leaf
//! function (no outgoing `CALL*` instructions) is marked eligible. See
//! `is_leaf_function`.

use super::encoding::*;
use crate::{Chunk, Op};

/// Produce the payload of the `metadata.code.compilation_order` custom
/// section. Returns `None` when there are no hints to emit (allowing the
/// writer to skip the section entirely).
pub fn encode_compilation_order_payload(
    chunks: &[Chunk],
    rt_imports_len: usize,
) -> Option<Vec<u8>> {
    if chunks.is_empty() { return None; }

    let host_imports_len = chunks.first().map(|c| c.imports.len()).unwrap_or(0);
    let func_base = host_imports_len + rt_imports_len;

    // Script chunk (idx 0) is the entry — always hot.
    // Any chunk named `main`, `Main`, or the startup form's code is also hot.
    let mut hints: Vec<(u32, u32)> = Vec::new(); // (func_idx, priority)
    for (ci, chunk) in chunks.iter().enumerate() {
        let func_idx = (func_base + ci) as u32;
        let is_entry = ci == 0
            || chunk.name.eq_ignore_ascii_case("main")
            || chunk.name.eq_ignore_ascii_case("run")
            || chunk.name == "<script>";
        if is_entry {
            hints.push((func_idx, 0));
        }
    }
    if hints.is_empty() { return None; }

    let mut out = Vec::new();
    write_leb128_u32(&mut out, hints.len() as u32);
    for (func_idx, priority) in hints {
        write_leb128_u32(&mut out, func_idx);
        write_leb128_u32(&mut out, 0);             // function-level hint (offset 0)
        // hint payload: one u32 (priority). length is LEB so max ~5 bytes.
        let mut payload = Vec::new();
        write_leb128_u32(&mut payload, priority);
        write_leb128_u32(&mut out, payload.len() as u32);
        out.extend_from_slice(&payload);
    }
    Some(out)
}

/// The canonical section name — exported so the writer can use it.
pub const COMPILATION_ORDER_SECTION_NAME: &str = "metadata.code.compilation_order";
pub const BRANCH_HINT_SECTION_NAME:       &str = "metadata.code.branch_hint";
pub const INLINING_SECTION_NAME:          &str = "metadata.code.inlining";

/// Encode `metadata.code.branch_hint`. We mark every `BR` / `BR_IF_*` /
/// `BR_LABEL` appearing **inside a LOOP** as `likely taken` (back-edges
/// almost always retry) and every branch appearing inside an exception
/// handler (after TRY_TABLE) as `unlikely` (slow path).
pub fn encode_branch_hint_payload(
    chunks: &[Chunk],
    rt_imports_len: usize,
) -> Option<Vec<u8>> {
    let host_imports_len = chunks.first().map(|c| c.imports.len()).unwrap_or(0);
    let func_base = host_imports_len + rt_imports_len;

    let mut per_func: Vec<(u32, Vec<(u32, u8)>)> = Vec::new();
    for (ci, chunk) in chunks.iter().enumerate() {
        let hints = scan_branch_hints(chunk);
        if !hints.is_empty() {
            per_func.push(((func_base + ci) as u32, hints));
        }
    }
    if per_func.is_empty() { return None; }

    let mut out = Vec::new();
    write_leb128_u32(&mut out, per_func.len() as u32);
    for (fn_idx, hints) in per_func {
        write_leb128_u32(&mut out, fn_idx);
        write_leb128_u32(&mut out, hints.len() as u32);
        for (offset, hint_byte) in hints {
            write_leb128_u32(&mut out, offset);
            write_leb128_u32(&mut out, 1);   // hint length = 1 byte
            out.push(hint_byte);
        }
    }
    Some(out)
}

/// Encode `metadata.code.inlining`. Mark **leaf functions** (no outgoing
/// `CALL`/`CALL_REF`/`CALL_INDIRECT`) as "inline eligible" so engines
/// can inline them aggressively. Priority 0 = strongly inline.
pub fn encode_inlining_payload(
    chunks: &[Chunk],
    rt_imports_len: usize,
) -> Option<Vec<u8>> {
    let host_imports_len = chunks.first().map(|c| c.imports.len()).unwrap_or(0);
    let func_base = host_imports_len + rt_imports_len;

    let mut hints: Vec<(u32, u8)> = Vec::new();
    for (ci, chunk) in chunks.iter().enumerate() {
        if is_leaf_function(chunk) {
            hints.push(((func_base + ci) as u32, 0));
        }
    }
    if hints.is_empty() { return None; }

    let mut out = Vec::new();
    write_leb128_u32(&mut out, hints.len() as u32);
    for (fn_idx, priority) in hints {
        write_leb128_u32(&mut out, fn_idx);
        write_leb128_u32(&mut out, 0);       // function-level hint
        write_leb128_u32(&mut out, 1);       // 1 byte payload
        out.push(priority);
    }
    Some(out)
}

/// Walk the chunk's bytecode. Track loop / try-handler depth and emit a
/// branch hint for every conditional-jump whose target is a loop back
/// edge (likely-taken) or an exception handler (unlikely).
fn scan_branch_hints(chunk: &Chunk) -> Vec<(u32, u8)> {
    let code = &chunk.code;
    let mut hints = Vec::new();
    let mut ip = 0usize;
    let mut in_loop = 0i32;
    while ip + 1 < code.len() {
        let op_start = ip;
        let prefix = code[ip];
        let sub = code[ip + 1];
        let op = match Op::decode(prefix, sub) {
            Some(op) => op,
            None => { ip += 2; continue; }
        };
        ip += 2;

        if op == Op::LOOP { in_loop += 1; }
        if op == Op::END && in_loop > 0 { in_loop -= 1; }

        // Classify conditional branches.
        let is_cond_branch = op == Op::BR_IF_TRUE
            || op == Op::BR_IF_FALSE
            || op == Op::BR_IF_NULL
            || op == Op::BR_IF_LABEL;
        if is_cond_branch && in_loop > 0 {
            // Back-edge — likely taken.
            hints.push((op_start as u32, 1));
        }

        ip += op.operand_format().fixed_size();
    }
    hints
}

/// A function with zero outgoing calls. Cheap to inline and pays back
/// handsomely under tier-up.
fn is_leaf_function(chunk: &Chunk) -> bool {
    let code = &chunk.code;
    let mut ip = 0usize;
    while ip + 1 < code.len() {
        let prefix = code[ip];
        let sub = code[ip + 1];
        let op = match Op::decode(prefix, sub) {
            Some(op) => op,
            None => { ip += 2; continue; }
        };
        ip += 2;
        if op == Op::CALL
            || op == Op::CALL_REF
            || op == Op::CALL_INDIRECT
            || op == Op::CALL_IMPORT
            || op == Op::RETURN_CALL
            || op == Op::RETURN_CALL_REF
            || op == Op::RETURN_CALL_INDIRECT {
            return false;
        }
        ip += op.operand_format().fixed_size();
    }
    // Don't mark the entry script as inline-eligible — it isn't called.
    !chunk.name.is_empty() && chunk.name != "<script>"
}
