//! Python `re` surface over the ECMA `ecma:regexp` runtime — the SAME host ops
//! JS/Java/Lua drive. Python regex ≈ JS regex, so patterns pass through. No new
//! host fns. Match objects are the JS exec arrays (`m[0]` full match, `m[i]`
//! group i, `m.index` position); the walker rewrites `m.group(i)`→`m[i]` etc.
//!
//! Conventions (from java/string_adapter): `new(pattern[, flags])`→regexp;
//! `exec(regexp, str)`; `match/matchAll/search/split(str, regexp)`;
//! `replace/replaceAll(str, regexp, replacement)`.

use vybe_compiler::primitives::tuples;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn lget(c: &mut Chunk, s: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, s, line);
}
fn lset(c: &mut Chunk, s: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_SET, s, line);
}
fn call(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}
fn stash(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for off in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + off, line);
    }
    base
}

/// Build a regexp from slot `pat` with optional flags string; leaves regexp on stack.
fn build_regexp(chunks: &mut [Chunk], current: usize, pat: u16, flags: Option<&str>, line: u32) {
    lget(&mut chunks[current], pat, line);
    match flags {
        Some(f) => {
            chunks[current].emit_string_const(f, line);
            call(chunks, current, "ecma:regexp", "new", 2, line);
        }
        None => call(chunks, current, "ecma:regexp", "new", 1, line),
    }
}

/// `re.search(pat, s)` → `exec(new(pat), s)` → match array or null.
pub fn emit_search(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line);
    build_regexp(chunks, current, base, None, line);
    let r = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], r, line);
    lget(&mut chunks[current], r, line);
    lget(&mut chunks[current], base + 1, line);
    call(chunks, current, "ecma:regexp", "exec", 2, line);
}

/// `re.match(pat, s)` — anchored at start: `exec(new("^(?:"+pat+")"), s)`.
pub fn emit_match(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line);
    // pat2 = "^(?:" + pat + ")"
    chunks[current].emit_string_const("^(?:", line);
    lget(&mut chunks[current], base, line);
    call(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_string_const(")", line);
    call(chunks, current, "wasm:js-string", "concat", 2, line);
    let pat2 = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], pat2, line);
    build_regexp(chunks, current, pat2, None, line);
    let r = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], r, line);
    lget(&mut chunks[current], r, line);
    lget(&mut chunks[current], base + 1, line);
    call(chunks, current, "ecma:regexp", "exec", 2, line);
}

/// `re.sub(pat, repl, s)` → `replaceAll(s, new(pat, "g"), repl)`.
pub fn emit_sub(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line); // pat, repl, s
    lget(&mut chunks[current], base + 2, line); // s
    build_regexp(chunks, current, base, Some("g"), line); // regexp
    lget(&mut chunks[current], base + 1, line); // repl
    call(chunks, current, "ecma:regexp", "replaceAll", 3, line);
}

/// `re.split(pat, s)` → `split(s, new(pat))`.
pub fn emit_split(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line);
    lget(&mut chunks[current], base + 1, line); // s
    build_regexp(chunks, current, base, None, line);
    call(chunks, current, "ecma:regexp", "split", 2, line);
}

/// `re.escape(s)` → `ecma:regexp.escape(s)`.
pub fn emit_escape(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line);
    lget(&mut chunks[current], base, line);
    call(chunks, current, "ecma:regexp", "escape", 1, line);
}

/// `re.findall(pat, s)` → list of full matches (no group), group[1] (1 group),
/// or tuple(groups) (>1 group), via `matchAll(s, new(pat, "g"))`.
pub fn emit_findall(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(chunks, current, argc, line); // pat, s
    lget(&mut chunks[current], base + 1, line); // s
    build_regexp(chunks, current, base, Some("g"), line);
    call(chunks, current, "ecma:regexp", "matchAll", 2, line);
    let matches = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], matches, line);

    let res = chunks[current].alloc_scratch(1);
    call(chunks, current, "ecma:array", "new", 0, line);
    lset(&mut chunks[current], res, line);

    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    let mlen = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], matches, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    lset(&mut chunks[current], n, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], i, line);

    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    lget(&mut chunks[current], i, line);
    lget(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    // m = matches[i]; mlen = m.length
    lget(&mut chunks[current], matches, line);
    lget(&mut chunks[current], i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(&mut chunks[current], m, line);
    lget(&mut chunks[current], m, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    lset(&mut chunks[current], mlen, line);

    // value = (mlen > 2) ? tuple(m[1:]) : (mlen == 2) ? m[1] : m[0]
    lget(&mut chunks[current], mlen, line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    {
        // tuple(m.slice(1))
        lget(&mut chunks[current], m, line);
        chunks[current].emit_i32_const(1, line);
        call(chunks, current, "ecma:array", "slice", 2, line);
        tuples::emit_tag(chunks, current, line);
    }
    chunks[current].emit_else(line);
    {
        lget(&mut chunks[current], mlen, line);
        chunks[current].emit_i32_const(2, line);
        chunks[current].emit_op(Op::I32_EQ, line);
        chunks[current].emit_if_value(line);
        lget(&mut chunks[current], m, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        chunks[current].emit_else(line);
        lget(&mut chunks[current], m, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);
    let val = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], val, line);

    // res.push(val)
    lget(&mut chunks[current], res, line);
    lget(&mut chunks[current], val, line);
    call(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    lget(&mut chunks[current], res, line);
}
