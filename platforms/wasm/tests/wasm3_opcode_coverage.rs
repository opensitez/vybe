//! Every WASM 3.0 opcode is EXERCISED by our own suite — derived from the
//! opcode table, not from a list.
//!
//! "Do we support all of WASM 3.0?" cannot be answered by a pass count on
//! someone else's suite: a file that never mentions an instruction says nothing
//! about it, and a file we cannot parse says nothing about anything. It also
//! cannot be answered by a hand-kept checklist, because a checklist cannot
//! report what nobody typed into it — the same failure `surface_from_wit.rs`
//! removed for WASI.
//!
//! So the DECLARED side is `Op`'s own per-category name tables, which
//! `from_wasm_name`/`operand_format` already treat as the single source of
//! truth, and the COVERED side is the mnemonics that actually appear in
//! `tests/wast`. The gap between them is the answer.
//!
//! ## What counts as WASM 3.0
//!
//! The merged proposals, which is opcode groups `0x00` (core, incl.
//! sign-extension and reference types), `0xFB` (GC + function references),
//! `0xFC` (bulk memory, non-trapping float→int, table ops) and `0xFD` (SIMD +
//! relaxed SIMD).
//!
//! Deliberately EXCLUDED, because they are not 3.0: `0xFE` threads, `0xF0`
//! canon (component model), `0xF1` call tags, `0xFF` VM-internal.

use std::collections::BTreeSet;
use std::path::{Path, PathBuf};

/// The opcode groups the merged 3.0 proposals occupy.
const WASM3_GROUPS: &[u16] = &[0x00, 0xFB, 0xFC, 0xFD];

/// Mnemonics with no WAT spelling to look for, each with its reason.
///
/// Kept SHORT and justified one at a time: an entry here is a hole in the gate,
/// so "it was failing" is never a reason to add one. Two kinds qualify, and
/// nothing else does — a genuinely untested 3.0 instruction belongs in
/// `tests/wast`, not here.
fn is_not_written_in_wat(name: &str) -> bool {
    // ── 1. Opcodes that are not WASM 3.0, sharing a table with ones that are.
    //
    // `stringref` was never merged (3.0's string story is the js-string-builtins
    // HOST IMPORTS, not instructions). `*_desc*` and `ref.get_desc` are the
    // custom-descriptors proposal. `delegate` is the LEGACY exception proposal,
    // replaced by `try_table`. `resume_throw*` is stack switching.
    if name.starts_with("stringview_")
        || name.starts_with("string.")
        || name.starts_with("stringref.")
        || name.contains("_desc")
        || name == "ref.get_desc"
        || name == "delegate"
        || name.starts_with("resume_")
    {
        return true;
    }
    // ── 2. Opcodes with no distinct MNEMONIC of their own.
    //
    // The WAT text spells the nullable cast and test as `ref.cast (ref null $t)`
    // / `ref.test (ref null $t)` and the typed select as `select (result t)`.
    // The `_null` / `_t` names are the BINARY encodings' names, so grepping the
    // text for them finds nothing however thoroughly they are exercised — and
    // they are: `ref.cast`/`ref.test`/`select` carry the coverage.
    matches!(name, "ref.cast_null" | "ref.test_null" | "select_t")
}

fn repo_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()
        .expect("the repository root is two levels above platforms/wasm")
}

/// Every mnemonic the opcode table names, in the 3.0 groups.
fn declared_wasm3_mnemonics() -> BTreeSet<String> {
    use vybe_runtime::opcode::Op;
    let mut out = BTreeSet::new();
    for &group in WASM3_GROUPS {
        for sub in 0..=u16::MAX {
            if let Some(name) = Op::new(group, sub).wasm_name_opt() {
                if !is_not_written_in_wat(name) {
                    out.insert(name.to_string());
                }
            }
        }
    }
    out
}

/// Every source byte of our own wast suite, concatenated.
fn corpus_text() -> String {
    fn walk(dir: &Path, out: &mut String) {
        let Ok(entries) = std::fs::read_dir(dir) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                walk(&path, out);
            } else if matches!(
                path.extension().and_then(|e| e.to_str()),
                Some("wat") | Some("wast")
            ) {
                if let Ok(text) = std::fs::read_to_string(&path) {
                    out.push_str(&text);
                    out.push('\n');
                }
            }
        }
    }
    let mut out = String::new();
    walk(&repo_root().join("tests/wast"), &mut out);
    out
}

/// The table must actually yield mnemonics, or every assertion below is vacuous.
#[test]
fn the_opcode_table_names_the_wasm3_surface() {
    let declared = declared_wasm3_mnemonics();
    assert!(
        declared.len() > 400,
        "the opcode table produced only {} WASM 3.0 mnemonics — the enumeration is \
         broken, and a coverage claim derived from it would be meaningless",
        declared.len()
    );
    for expected in [
        "i32.add",
        "memory.copy",
        "table.grow",
        "ref.cast",
        "struct.new",
        "array.len",
        "i8x16.shuffle",
        "f32x4.relaxed_madd",
        "i32.trunc_sat_f32_s",
        "return_call",
    ] {
        assert!(
            declared.contains(expected),
            "{expected} is a WASM 3.0 instruction and the opcode table does not name it"
        );
    }
}

/// The corpus must be found, or "everything is covered" would be vacuous too.
#[test]
fn the_wast_corpus_is_readable() {
    let text = corpus_text();
    assert!(
        text.len() > 100_000,
        "read only {} bytes out of tests/wast — the corpus walk is broken",
        text.len()
    );
}

/// Does `mnemonic` occur in `text` as a WHOLE token?
///
/// A plain substring test is too lenient in exactly the direction that matters:
/// `i32.load` occurs inside `i32.load8_u`, so a corpus that exercises only the
/// narrow load would report the wide one covered. WAT ends a mnemonic at the
/// first non-idchar, so that is the boundary checked here.
fn mentions_mnemonic(text: &str, mnemonic: &str) -> bool {
    let is_idchar = |c: char| c.is_ascii_alphanumeric() || c == '_' || c == '.';
    text.match_indices(mnemonic).any(|(at, _)| {
        let before_ok = at == 0 || !is_idchar(text[..at].chars().next_back().unwrap_or(' '));
        let after_ok = text[at + mnemonic.len()..]
            .chars()
            .next()
            .is_none_or(|c| !is_idchar(c));
        before_ok && after_ok
    })
}

/// EVERY WASM 3.0 opcode appears in our own suite.
#[test]
fn every_wasm3_opcode_is_exercised_by_our_suite() {
    let text = corpus_text();
    let missing: Vec<String> = declared_wasm3_mnemonics()
        .into_iter()
        .filter(|name| !mentions_mnemonic(&text, name))
        .collect();
    assert!(
        missing.is_empty(),
        "{} WASM 3.0 instructions are named by the opcode table and appear NOWHERE in \
         tests/wast:\n  {}",
        missing.len(),
        missing.join("\n  ")
    );
}
