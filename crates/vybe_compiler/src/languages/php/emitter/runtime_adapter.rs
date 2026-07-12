//! PHP runtime-surface helpers routed via `common:php.*`.

use crate::emitter::collections;
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

pub fn emit_helper(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    // `array_map(null, a, b, …)` (walker-rewritten to `zip`) → array of tuples
    // padded to the LONGEST input (PHP semantics). Shared `vybe_emitter` op.
    if name == "php.zip" {
        collections::emit_zip(chunks, current, argc, collections::ZipLen::Longest, line);
        return true;
    }
    // Regex adapters. `preg_match_all($pat, $s)` / `preg_replace($pat, $repl, $s)`
    // are pattern-first; `ecma:regexp` is subject-first. Reorder args through
    // scratch locals and call the host fn directly instead of routing through
    // the `__ecma_regexp_*_pat_first` bundle chunks (just this reorder + call).
    match name {
        "php.regex_match_all_pat_first" => {
            emit_regexp_pat_first(chunks, current, "matchAll", line);
            return true;
        }
        "php.regex_replace_pat_first" => {
            emit_regexp_replace_pat_first(chunks, current, line);
            return true;
        }
        _ => {}
    }
    let global = match name {
        "php.isnumeric" => "__vybe_isnumeric",
        "php.sort_in_place" => "__vybe_sort_in_place",
        "php.sort_with_comparator" => "__vybe_sort_with_comparator",
        "php.uniq" => "__vybe_uniq",
        _ => return false,
    };
    collections::emit_runtime_helper_call(chunks, current, global, argc, line);
    true
}

/// Pattern-first 2-arg regex adapter (`preg_match_all`→`matchAll`).
/// Stack `[pat, subject]` → `ecma:regexp.<method>(subject, pat)`.
fn emit_regexp_pat_first(chunks: &mut [Chunk], current: usize, method: &str, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    chunks[current].emit_op_u16(Op::LOCAL_SET, base + 1, line); // subject (top)
    chunks[current].emit_op_u16(Op::LOCAL_SET, base, line); // pat
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line); // subject
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line); // pat
    let idx = chunks[current].add_import("ecma:regexp", method);
    chunks[current].emit_call(idx, 2, line);
}

/// `preg_replace($pat, $repl, $subject)` →
/// `ecma:regexp.replaceAll(subject, pat, repl)` (always-global, PHP semantics).
fn emit_regexp_replace_pat_first(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    chunks[current].emit_op_u16(Op::LOCAL_SET, base + 2, line); // subject (top)
    chunks[current].emit_op_u16(Op::LOCAL_SET, base + 1, line); // repl
    chunks[current].emit_op_u16(Op::LOCAL_SET, base, line); // pat
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 2, line); // subject
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line); // pat
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line); // repl
    let idx = chunks[current].add_import("ecma:regexp", "replaceAll");
    chunks[current].emit_call(idx, 3, line);
}
