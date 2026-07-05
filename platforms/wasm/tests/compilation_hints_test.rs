//! Tests for the compilation-hints WASM proposal.
//! Spec: `proposals/compilation-hints/`
//!
//! Covers the `metadata.code.compilation_order` and
//! `metadata.code.inlining` custom sections. These are trailing custom
//! sections (no ordering constraint relative to the code section).

use vybe_platform_wasm::writer::proposals::compilation_hints::{
    COMPILATION_ORDER_SECTION_NAME, INLINING_SECTION_NAME,
};
use vybe_bytecode::{Chunk, Op};

// ── helpers ───────────────────────────────────────────────────────────────

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

fn find_custom_section<'a>(wasm: &'a [u8], name: &str) -> Option<&'a [u8]> {
    let mut pos = 8usize;
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

fn emit_wasm(chunks: Vec<Chunk>) -> Vec<u8> {
    vybe_platform_wasm::write_wasm(&chunks)
}

fn script_chunk() -> Chunk {
    let mut s = Chunk::new("<script>");
    s.emit_op(Op::NULL, 0);
    s.emit_op(Op::RETURN, 0);
    s
}

fn leaf_chunk(name: &str) -> Chunk {
    let mut f = Chunk::new(name);
    f.arity = 0;
    f.emit_op(Op::NULL, 0);
    f.emit_op(Op::RETURN, 0);
    f
}

// ── metadata.code.compilation_order ──────────────────────────────────────

#[test]
fn compilation_order_section_marks_entry_as_priority_zero() {
    let wasm = emit_wasm(vec![script_chunk()]);
    let payload = find_custom_section(&wasm, COMPILATION_ORDER_SECTION_NAME)
        .expect("compilation_order section must be present");
    let (count, mut pos) = leb128(payload);
    assert!(count >= 1, "must have at least one hint");
    let (_fn_idx, n) = leb128(&payload[pos..]);
    pos += n;
    let (_offset, n) = leb128(&payload[pos..]);
    pos += n;
    let (_len, n) = leb128(&payload[pos..]);
    pos += n;
    let (priority, _) = leb128(&payload[pos..]);
    assert_eq!(priority, 0, "entry chunk must have priority 0 (highest)");
}

#[test]
fn compilation_order_section_marks_main_chunk_hot() {
    let mut main_chunk = Chunk::new("main");
    main_chunk.arity = 0;
    main_chunk.emit_op(Op::NULL, 0);
    main_chunk.emit_op(Op::RETURN, 0);

    let wasm = emit_wasm(vec![script_chunk(), main_chunk]);
    let payload = find_custom_section(&wasm, COMPILATION_ORDER_SECTION_NAME)
        .expect("compilation_order section must be present");
    let (count, _) = leb128(payload);
    assert!(count >= 2, "both <script> and main should be marked hot");
}

#[test]
fn compilation_order_absent_for_non_entry_only_module() {
    // A module with only non-hot functions gets no compilation_order hints
    // for those functions. The <script> chunk still appears.
    let wasm = emit_wasm(vec![script_chunk(), leaf_chunk("helper")]);
    let payload = find_custom_section(&wasm, COMPILATION_ORDER_SECTION_NAME)
        .expect("compilation_order must still be present (script is hot)");
    let (count, _) = leb128(payload);
    assert_eq!(
        count, 1,
        "only <script> should be marked hot, not the helper"
    );
}

// ── metadata.code.inlining ────────────────────────────────────────────────

#[test]
fn inlining_section_present_for_leaf_function() {
    let wasm = emit_wasm(vec![script_chunk(), leaf_chunk("leaf")]);
    assert!(
        find_custom_section(&wasm, INLINING_SECTION_NAME).is_some(),
        "inlining section must be present for a leaf function"
    );
}

#[test]
fn inlining_section_leaf_has_priority_zero() {
    let wasm = emit_wasm(vec![script_chunk(), leaf_chunk("leaf")]);
    let payload = find_custom_section(&wasm, INLINING_SECTION_NAME).unwrap();
    let (count, mut pos) = leb128(payload);
    assert!(count >= 1);
    let (_fn_idx, n) = leb128(&payload[pos..]);
    pos += n;
    let (_offset, n) = leb128(&payload[pos..]);
    pos += n;
    let (_len, n) = leb128(&payload[pos..]);
    pos += n;
    let priority = payload[pos];
    assert_eq!(priority, 0, "leaf function must have inlining priority 0");
}

#[test]
fn inlining_section_absent_for_non_leaf() {
    let mut callee = Chunk::new("callee");
    callee.arity = 0;
    callee.emit_op(Op::NULL, 0);
    callee.emit_op(Op::RETURN, 0);

    let mut caller = Chunk::new("caller");
    caller.arity = 0;
    caller.emit_op_u16(Op::CALL, 2, 0); // call chunk index 2
    caller.emit_op(Op::RETURN, 0);

    let wasm = emit_wasm(vec![script_chunk(), callee, caller]);
    if let Some(p) = find_custom_section(&wasm, INLINING_SECTION_NAME) {
        // If present, the non-leaf caller must not be listed.
        let (count, mut pos) = leb128(p);
        let import_base = {
            // script chunk's import count (0 here) + rt_imports (0 in test)
            0u32
        };
        // Collect all fn_idx entries.
        let mut fn_indices = Vec::new();
        for _ in 0..count {
            let (fn_idx, n) = leb128(&p[pos..]);
            pos += n;
            let (_off, n) = leb128(&p[pos..]);
            pos += n;
            let (_len, n) = leb128(&p[pos..]);
            pos += n;
            pos += 1; // priority byte
            fn_indices.push(fn_idx);
            let _ = import_base;
        }
        // caller is chunk index 2 (0-based) → fn_idx = 2 (no imports in test)
        // It must NOT appear in the inlining list.
        assert!(
            !fn_indices.contains(&2),
            "non-leaf caller (fn_idx=2) must not appear in inlining section"
        );
    }
    // If absent entirely that's also fine.
}

#[test]
fn inlining_section_lists_multiple_leaf_functions() {
    let wasm = emit_wasm(vec![
        script_chunk(),
        leaf_chunk("leaf_a"),
        leaf_chunk("leaf_b"),
        leaf_chunk("leaf_c"),
    ]);
    let payload = find_custom_section(&wasm, INLINING_SECTION_NAME)
        .expect("inlining section must be present");
    let (count, _) = leb128(payload);
    assert_eq!(
        count, 3,
        "all three leaf functions must appear in inlining section"
    );
}

#[test]
fn inlining_section_does_not_include_script_entry() {
    // <script> chunk is the entry — it is never called, so it must not
    // be marked inline-eligible.
    let wasm = emit_wasm(vec![script_chunk()]);
    // With only the entry chunk there should be no inlining section.
    assert!(
        find_custom_section(&wasm, INLINING_SECTION_NAME).is_none(),
        "script entry chunk must not appear in inlining section"
    );
}
