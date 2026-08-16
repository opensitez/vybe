//! Regex primitives.
//!
//! ECMA-262 is the regex engine: `ecma:regexp.*` host functions do the
//! matching, and everything here is argument-order adaptation on top of
//! them.
//!
//! The adaptation exists because stdlib conventions disagree about where
//! the pattern goes. PHP `preg_*` and Python `re.*` put the pattern
//! FIRST — it reads as the subject of the call. ECMA-262 puts the string
//! first because these are `String.prototype` methods and the string is
//! the receiver. Neither is wrong; they just can't share a signature, so
//! these chunks take arguments in language order and re-push them in
//! ECMA order.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// `preg_replace($pat, $repl, $str)` / `re.sub(pat, repl, str)` →
/// `ecma:regexp.replaceAll(str, pat, repl)`.
///
/// Routes to `replaceAll`, not `replace`: PHP and Python replace EVERY
/// match by default, while JS `str.replace` is single-match unless the
/// regex carries `/g`. Going through `replaceAll` preserves the
/// always-global semantic without having to force a `/g` flag into the
/// pattern string.
pub(crate) fn build_regex_replace_pat_first(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_regex_replace_pat_first");
    let idx = c.add_import("ecma:regexp", "replaceAll");
    c.arity = 3;
    c.local_count = 3; // pat(0), repl(1), str(2)
    // Push (str, pat, repl) — ecma:regexp.replaceAll order.
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_call(idx, 3, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

/// `preg_split($pat, $str)` / `re.split(pat, str)` →
/// `ecma:regexp.split(str, pat)`.
pub(crate) fn build_regex_split_pat_first(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_regex_split_pat_first");
    let idx = c.add_import("ecma:regexp", "split");
    c.arity = 2;
    c.local_count = 2; // pat(0), str(1)
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(idx, 2, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

/// `preg_match_all($pat, $str)` / `re.findall(pat, str)` →
/// `ecma:regexp.matchAll(str, pat)`.
pub(crate) fn build_regex_match_all_pat_first(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_regex_match_all_pat_first");
    let idx = c.add_import("ecma:regexp", "matchAll");
    c.arity = 2;
    c.local_count = 2; // pat(0), str(1)
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(idx, 2, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
