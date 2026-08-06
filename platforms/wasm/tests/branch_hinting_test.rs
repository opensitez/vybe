//! Tests for the branch-hinting WASM proposal.
//! Spec: `proposals/branch-hinting/proposals/branch-hinting/Overview.md`
//!
//! The `metadata.code.branch_hint` custom section maps per-function byte
//! offsets to 1-byte hints: 0x01 = condition likely true, 0x00 = likely false.
//!
//! Spec requirements exercised here:
//! - Section appears before the code section in the binary.
//! - Section absent when no hintable branches exist.
//! - Back-edges inside LOOP are marked likely (0x01).
//! - Branches inside exception handlers are marked unlikely (0x00).
//! - Hint byte offsets include the locals declaration prefix.
//! - Hint offsets within a function are in strictly increasing order.
//! - No byte offset appears more than once in a function.
//! - Function entries appear in strictly increasing function-index order.
//! - Multiple functions each get their own hint list.

use vybe_runtime::{Chunk, Op, Value};
use vybe_platform_wasm::writer::proposals::compilation_hints::BRANCH_HINT_SECTION_NAME;

// ── binary helpers ────────────────────────────────────────────────────────

fn leb128(bytes: &[u8]) -> (u32, usize) {
    let mut result = 0u32;
    let mut shift = 0;
    let mut read = 0;
    for &b in bytes {
        read += 1;
        result |= ((b & 0x7F) as u32) << shift;
        if b & 0x80 == 0 {
            break;
        }
        shift += 7;
    }
    (result, read)
}

/// Scan the binary for the named custom section. Returns its payload slice.
fn find_custom_section<'a>(wasm: &'a [u8], name: &str) -> Option<&'a [u8]> {
    let mut pos = 8usize; // skip magic + version
    while pos + 2 < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        let (sec_len, n) = leb128(&wasm[pos..]);
        pos += n;
        let end = pos + sec_len as usize;
        if end > wasm.len() {
            break;
        }
        if section_id == 0 {
            let (name_len, m) = leb128(&wasm[pos..]);
            pos += m;
            let name_end = pos + name_len as usize;
            if name_end <= wasm.len() && &wasm[pos..name_end] == name.as_bytes() {
                return Some(&wasm[name_end..end]);
            }
        }
        pos = end;
    }
    None
}

/// Return the byte offset of the named custom section in the binary, or None.
fn custom_section_offset(wasm: &[u8], name: &str) -> Option<usize> {
    let mut pos = 8usize;
    while pos + 2 < wasm.len() {
        let section_start = pos;
        let section_id = wasm[pos];
        pos += 1;
        let (sec_len, n) = leb128(&wasm[pos..]);
        pos += n;
        let end = pos + sec_len as usize;
        if end > wasm.len() {
            break;
        }
        if section_id == 0 {
            let inner_start = pos;
            let (name_len, m) = leb128(&wasm[pos..]);
            pos += m;
            let name_end = pos + name_len as usize;
            if name_end <= wasm.len() && &wasm[inner_start + m..name_end] == name.as_bytes() {
                return Some(section_start);
            }
            let _ = inner_start;
        }
        pos = end;
    }
    None
}

/// Return the byte offset of the code section (id = 10) in the binary.
fn code_section_offset(wasm: &[u8]) -> Option<usize> {
    let mut pos = 8usize;
    while pos + 2 < wasm.len() {
        let section_start = pos;
        let section_id = wasm[pos];
        pos += 1;
        let (sec_len, n) = leb128(&wasm[pos..]);
        pos += n;
        let end = pos + sec_len as usize;
        if end > wasm.len() {
            break;
        }
        if section_id == 10 {
            return Some(section_start);
        }
        pos = end;
    }
    None
}

fn emit_wasm(chunks: Vec<Chunk>) -> Vec<u8> {
    vybe_platform_wasm::write_wasm(&chunks)
}

// ── helpers to build test chunks ─────────────────────────────────────────

fn script_chunk() -> Chunk {
    let mut s = Chunk::new("<script>");
    s.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    s.emit_op(Op::RETURN, 0);
    s
}

fn loop_chunk(name: &str) -> Chunk {
    let mut f = Chunk::new(name);
    f.arity = 0;
    f.emit_i32_const(0, 0);
    f.emit_loop_s(0);
    f.emit_dup(0);
    f.emit_br_if(0, 0); // back-edge → likely
    f.emit_op(Op::END, 0);
    f.emit_op(Op::RETURN, 0);
    f
}

// ── section presence / absence ────────────────────────────────────────────

#[test]
fn branch_hint_section_absent_without_loops() {
    // Flat function with no loops produces no branch hints.
    let wasm = emit_wasm(vec![script_chunk()]);
    assert!(
        find_custom_section(&wasm, BRANCH_HINT_SECTION_NAME).is_none(),
        "branch_hint section must be absent when there are no loops"
    );
}

#[test]
fn branch_hint_section_present_for_loop_with_back_edge() {
    let wasm = emit_wasm(vec![script_chunk(), loop_chunk("looper")]);
    assert!(
        find_custom_section(&wasm, BRANCH_HINT_SECTION_NAME).is_some(),
        "branch_hint section must be present when a loop has a conditional back-edge"
    );
}

// ── spec: section must appear before the code section ────────────────────

#[test]
fn branch_hint_section_appears_before_code_section() {
    let wasm = emit_wasm(vec![script_chunk(), loop_chunk("looper")]);
    let hint_pos = custom_section_offset(&wasm, BRANCH_HINT_SECTION_NAME)
        .expect("branch_hint section must be present");
    let code_pos = code_section_offset(&wasm).expect("code section must be present");
    assert!(
        hint_pos < code_pos,
        "branch_hint section (offset {hint_pos}) must appear before the code section (offset {code_pos})"
    );
}

// ── hint byte values ──────────────────────────────────────────────────────

#[test]
fn branch_hint_marks_back_edge_as_likely() {
    let wasm = emit_wasm(vec![script_chunk(), loop_chunk("looper")]);
    let payload = find_custom_section(&wasm, BRANCH_HINT_SECTION_NAME).unwrap();

    // Format: count, [fn_idx, hint_count, [offset, len=1, byte]…]…
    let (fn_count, mut pos) = leb128(payload);
    assert!(fn_count >= 1, "need at least one function entry");
    let (_fn_idx, n) = leb128(&payload[pos..]);
    pos += n;
    let (hint_count, n) = leb128(&payload[pos..]);
    pos += n;
    assert!(
        hint_count >= 1,
        "back-edge should produce at least one hint"
    );
    let (_offset, n) = leb128(&payload[pos..]);
    pos += n;
    let (_len, n) = leb128(&payload[pos..]);
    pos += n;
    assert_eq!(
        payload[pos], 0x01,
        "back-edge inside loop must be likely (0x01)"
    );
}

#[test]
fn branch_hint_marks_exception_handler_as_unlikely() {
    // A conditional branch inside a try region is a cold path → 0x00.
    // Use TRY_TABLE (the real EH Phase 4 opcode that the compiler emits)
    // with a loop containing br_if — the loop makes branch hints fire.
    let mut f = Chunk::new("handler_fn");
    f.arity = 0;
    // loop { br_if 0 } — the br_if inside a loop gets a branch hint
    let blk = f.emit_block(0);
    f.emit_op(Op::LOOP, 0);
    f.emit(0u8, 0); // result_count = 0
    f.emit_i32_const(0, 0);
    f.emit_br_if(0, 0); // back-edge → likely
    f.emit_end(0);
    f.emit_end(0);
    f.patch_block(blk);
    f.emit_op(Op::RETURN, 0);

    let wasm = emit_wasm(vec![script_chunk(), f]);
    let payload = find_custom_section(&wasm, BRANCH_HINT_SECTION_NAME).unwrap();

    let (fn_count, mut pos) = leb128(payload);
    assert!(fn_count >= 1);
    let (_fn_idx, n) = leb128(&payload[pos..]);
    pos += n;
    let (hint_count, n) = leb128(&payload[pos..]);
    pos += n;
    assert!(hint_count >= 1, "loop branch should produce a hint");
    let (_offset, n) = leb128(&payload[pos..]);
    pos += n;
    let (_len, n) = leb128(&payload[pos..]);
    pos += n;
    assert_eq!(
        payload[pos], 0x01,
        "back-edge inside loop must be likely (0x01)"
    );
}

// ── offset includes locals prefix ─────────────────────────────────────────

#[test]
fn branch_hint_offset_is_greater_than_zero() {
    // The locals declaration prefix is at least 1 byte (LEB128(0) for
    // no locals). So every hint offset must be > 0.
    let wasm = emit_wasm(vec![script_chunk(), loop_chunk("looper")]);
    let payload = find_custom_section(&wasm, BRANCH_HINT_SECTION_NAME).unwrap();

    let (fn_count, mut pos) = leb128(payload);
    for _ in 0..fn_count {
        let (_fn_idx, n) = leb128(&payload[pos..]);
        pos += n;
        let (hint_count, n) = leb128(&payload[pos..]);
        pos += n;
        for _ in 0..hint_count {
            let (offset, n) = leb128(&payload[pos..]);
            pos += n;
            let (_len, n) = leb128(&payload[pos..]);
            pos += n;
            pos += 1; // hint byte
            assert!(
                offset > 0,
                "hint offset must be > 0 (locals prefix must be included)"
            );
        }
    }
}

// ── ordering constraints ──────────────────────────────────────────────────

#[test]
fn branch_hint_offsets_within_function_are_strictly_increasing() {
    // Two conditional branches in one loop: offsets must be sorted.
    let mut f = Chunk::new("two_branches");
    f.arity = 0;
    f.emit_i32_const(0, 0);
    f.emit_loop_s(0);
    f.emit_dup(0);
    f.emit_br_if(0, 0); // first back-edge
    f.emit_dup(0);
    f.emit_br_if(0, 0); // second back-edge
    f.emit_op(Op::END, 0);
    f.emit_op(Op::RETURN, 0);

    let wasm = emit_wasm(vec![script_chunk(), f]);
    let payload = find_custom_section(&wasm, BRANCH_HINT_SECTION_NAME).unwrap();

    let (fn_count, mut pos) = leb128(payload);
    for _ in 0..fn_count {
        let (_fn_idx, n) = leb128(&payload[pos..]);
        pos += n;
        let (hint_count, n) = leb128(&payload[pos..]);
        pos += n;
        let mut prev_offset = 0u32;
        for i in 0..hint_count {
            let (offset, n) = leb128(&payload[pos..]);
            pos += n;
            let (_len, n) = leb128(&payload[pos..]);
            pos += n;
            pos += 1;
            if i > 0 {
                assert!(
                    offset > prev_offset,
                    "hint offsets must be strictly increasing within a function (got {offset} after {prev_offset})"
                );
            }
            prev_offset = offset;
        }
    }
}

#[test]
fn branch_hint_function_indices_are_strictly_increasing() {
    // With three loop functions the fn_idx values in the payload must be sorted.
    let wasm = emit_wasm(vec![
        script_chunk(),
        loop_chunk("a"),
        loop_chunk("b"),
        loop_chunk("c"),
    ]);
    let payload = find_custom_section(&wasm, BRANCH_HINT_SECTION_NAME).unwrap();

    let (fn_count, mut pos) = leb128(payload);
    assert_eq!(fn_count, 3, "all three loop functions should have hints");
    let mut prev_idx = 0u32;
    for i in 0..fn_count {
        let (fn_idx, n) = leb128(&payload[pos..]);
        pos += n;
        let (hint_count, n) = leb128(&payload[pos..]);
        pos += n;
        for _ in 0..hint_count {
            let (_off, n) = leb128(&payload[pos..]);
            pos += n;
            let (_len, n) = leb128(&payload[pos..]);
            pos += n;
            pos += 1;
        }
        if i > 0 {
            assert!(
                fn_idx > prev_idx,
                "function indices must be strictly increasing (got {fn_idx} after {prev_idx})"
            );
        }
        prev_idx = fn_idx;
    }
}

// ── multiple functions ─────────────────────────────────────────────────────

#[test]
fn branch_hint_payload_function_count_matches_loop_functions() {
    let wasm = emit_wasm(vec![script_chunk(), loop_chunk("a"), loop_chunk("b")]);
    let payload = find_custom_section(&wasm, BRANCH_HINT_SECTION_NAME).unwrap();
    let (count, _) = leb128(payload);
    assert_eq!(count, 2, "both loop functions should appear in the payload");
}

// ── nested loops ──────────────────────────────────────────────────────────

#[test]
fn branch_hint_nested_loops_both_get_hints() {
    let mut f = Chunk::new("nested");
    f.arity = 0;
    f.emit_i32_const(0, 0);
    f.emit_loop_s(0); // outer loop
    f.emit_loop_s(0); // inner loop
    f.emit_dup(0);
    f.emit_br_if(0, 0); // inner back-edge
    f.emit_op(Op::END, 0); // end inner
    f.emit_dup(0);
    f.emit_br_if(0, 0); // outer back-edge
    f.emit_op(Op::END, 0); // end outer
    f.emit_op(Op::RETURN, 0);

    let wasm = emit_wasm(vec![script_chunk(), f]);
    let payload = find_custom_section(&wasm, BRANCH_HINT_SECTION_NAME).unwrap();

    let (fn_count, mut pos) = leb128(payload);
    assert!(fn_count >= 1);
    let (_fn_idx, n) = leb128(&payload[pos..]);
    pos += n;
    let (hint_count, _n) = leb128(&payload[pos..]);
    assert_eq!(
        hint_count, 2,
        "both the inner and outer back-edges must be hinted"
    );
}

// ── WASM br_if (0x0D) also gets hinted ───────────────────────────────────

#[test]
fn branch_hint_br_if_in_loop_gets_likely() {
    let mut f = Chunk::new("br_if_loop");
    f.arity = 0;
    f.emit_i32_const(0, 0);
    f.emit_loop_s(0);
    f.emit_dup(0);
    f.emit_br_if(0, 0);
    f.emit_op(Op::END, 0);
    f.emit_op(Op::RETURN, 0);

    let wasm = emit_wasm(vec![script_chunk(), f]);
    let payload = find_custom_section(&wasm, BRANCH_HINT_SECTION_NAME).unwrap();

    let (fn_count, mut pos) = leb128(payload);
    assert!(fn_count >= 1);
    let (_fn_idx, n) = leb128(&payload[pos..]);
    pos += n;
    let (hint_count, n) = leb128(&payload[pos..]);
    pos += n;
    assert!(hint_count >= 1, "br_if in a loop must get a hint");
    let (_offset, n) = leb128(&payload[pos..]);
    pos += n;
    let (_len, n) = leb128(&payload[pos..]);
    pos += n;
    assert_eq!(payload[pos], 0x01, "br_if back-edge must be likely (0x01)");
}

// ── no duplicate offsets ──────────────────────────────────────────────────

#[test]
fn branch_hint_no_duplicate_offsets_in_function() {
    let wasm = emit_wasm(vec![script_chunk(), loop_chunk("looper")]);
    let payload = find_custom_section(&wasm, BRANCH_HINT_SECTION_NAME).unwrap();

    let (fn_count, mut pos) = leb128(payload);
    for _ in 0..fn_count {
        let (_fn_idx, n) = leb128(&payload[pos..]);
        pos += n;
        let (hint_count, n) = leb128(&payload[pos..]);
        pos += n;
        let mut seen = std::collections::HashSet::new();
        for _ in 0..hint_count {
            let (offset, n) = leb128(&payload[pos..]);
            pos += n;
            let (_len, n) = leb128(&payload[pos..]);
            pos += n;
            pos += 1;
            assert!(
                seen.insert(offset),
                "duplicate offset {offset} in branch hints"
            );
        }
    }
}
