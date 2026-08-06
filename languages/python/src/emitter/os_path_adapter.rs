//! Python `os.path` pure-string POSIX helpers.
//!
//! These compose CPython `posixpath` semantics on top of `ecma:string` /
//! `ecma:array` primitives, emitted directly at the call site (args already on
//! the stack). Routed via `common:python.ospath_*` from the profile `[builtins]`
//! table. The FS predicates (`exists`/`isfile`/`isdir`/`lexists`) stay
//! host-backed in the profile; only the string math lives here.
//!
//! Arguments arrive pre-pushed on the stack (left to right), matching the
//! `emit_common` free-function calling convention (same as `math_adapter`).

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use vybe_compiler::primitives::{collections, ops, strings, tuples};

/// `substring(start, END)` extends to the end of the string.
const END: i32 = 0x7FFF_FFFF;

fn call_import(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    // Register on the CURRENT chunk (see `string_adapter::call_import`).
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

/// Pop `argc` stack values (deepest first) into freshly allocated scratch slots
/// and return the base slot. `base + i` is the i-th argument.
fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

fn lget(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn lset(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}
fn sconst(chunks: &mut [Chunk], current: usize, s: &str, line: u32) {
    chunks[current].emit_string_const(s, line);
}
fn iconst(chunks: &mut [Chunk], current: usize, v: i32, line: u32) {
    chunks[current].emit_i32_const(v, line);
}
fn op(chunks: &mut [Chunk], current: usize, o: Op, line: u32) {
    chunks[current].emit_op(o, line);
}
/// String equality of the top two stack strings → i32 (0/1).
fn str_eq(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
}
/// Concatenate the top two stack strings. `[a, b] -> [a+b]`.
fn concat2(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
}
/// Push `s[start:end]` where `s` is in `slot`; `start`/`end` are raw i32 pushers.
fn push_substr(
    chunks: &mut [Chunk],
    current: usize,
    slot: u16,
    start: impl FnOnce(&mut [Chunk], usize),
    end: impl FnOnce(&mut [Chunk], usize),
    line: u32,
) {
    lget(chunks, current, slot, line);
    start(chunks, current);
    end(chunks, current);
    strings::emit_substring(&mut chunks[current], line);
}
/// Push `slot.rfind(needle)` as a raw i32 (-1 when absent).
fn push_rfind(chunks: &mut [Chunk], current: usize, slot: u16, needle: &str, line: u32) {
    lget(chunks, current, slot, line);
    sconst(chunks, current, needle, line);
    strings::emit_last_index_of(&mut chunks[current], line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    op(chunks, current, Op::I32_TRUNC_SAT_F64_S, line);
}
/// Push `len(slot)` as a raw i32.
fn push_len(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    lget(chunks, current, slot, line);
    strings::emit_length(&mut chunks[current], line);
}
/// Push `len(arr_slot)` as a raw i32 (array length via host `ecma:array.length`).
fn push_arr_len(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    lget(chunks, current, slot, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
}

// ── CPython posixpath: head-stripping ──────────────────────────────────────

/// Emit the `posixpath` head-trim: drop trailing `'/'` from the string in
/// `head_slot` **unless** it is empty or all slashes, leaving the result on the
/// stack. Equivalent to: `stripped = head.rstrip('/'); (stripped=='' and
/// head!='') ? head : stripped`.
fn emit_strip_head(chunks: &mut [Chunk], current: usize, head_slot: u16, line: u32) {
    let n = chunks[current].alloc_scratch(1);
    push_len(chunks, current, head_slot, line);
    lset(chunks, current, n, line);

    // while n>0 and head[n-1:n]=='/': n -= 1
    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, n, line);
    op(chunks, current, Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line); // n==0 → break
    push_substr(
        chunks,
        current,
        head_slot,
        |c, cur| {
            lget(c, cur, n, line);
            iconst(c, cur, 1, line);
            op(c, cur, Op::I32_SUB, line);
        },
        |c, cur| lget(c, cur, n, line),
        line,
    );
    sconst(chunks, current, "/", line);
    str_eq(chunks, current, line);
    op(chunks, current, Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line); // not '/' → break
    lget(chunks, current, n, line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_SUB, line);
    lset(chunks, current, n, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    // result = (n==0 and len(head)>0) ? head : head[0:n]
    lget(chunks, current, n, line);
    op(chunks, current, Op::I32_EQZ, line);
    push_len(chunks, current, head_slot, line);
    iconst(chunks, current, 0, line);
    op(chunks, current, Op::I32_GT_S, line);
    op(chunks, current, Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, head_slot, line);
    chunks[current].emit_else(line);
    push_substr(
        chunks,
        current,
        head_slot,
        |c, cur| iconst(c, cur, 0, line),
        |c, cur| lget(c, cur, n, line),
        line,
    );
    chunks[current].emit_end(line);
}

// ── CPython posixpath: normpath ────────────────────────────────────────────

/// Emit `normpath(src_slot)` and store the resulting string in `out_slot`.
/// Collapses `.`/empty components and resolves `..`, preserving a leading `/`.
fn emit_normpath_to(chunks: &mut [Chunk], current: usize, src: u16, out: u16, line: u32) {
    let initslash = chunks[current].alloc_scratch(1);
    let parts = chunks[current].alloc_scratch(1);
    let comps = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let comp = chunks[current].alloc_scratch(1);
    let clen = chunks[current].alloc_scratch(1);
    let last = chunks[current].alloc_scratch(1);

    // initslash = len(src)>0 and src[0:1]=='/'
    push_len(chunks, current, src, line);
    iconst(chunks, current, 0, line);
    op(chunks, current, Op::I32_GT_S, line);
    push_substr(
        chunks,
        current,
        src,
        |c, cur| iconst(c, cur, 0, line),
        |c, cur| iconst(c, cur, 1, line),
        line,
    );
    sconst(chunks, current, "/", line);
    str_eq(chunks, current, line);
    op(chunks, current, Op::I32_AND, line);
    lset(chunks, current, initslash, line);

    // parts = src.split('/'); comps = []
    lget(chunks, current, src, line);
    sconst(chunks, current, "/", line);
    strings::emit_split(&mut chunks[current], line);
    lset(chunks, current, parts, line);
    push_arr_len(chunks, current, parts, line);
    lset(chunks, current, n, line);
    call_import(chunks, current, "ecma:array", "new", 0, line);
    lset(chunks, current, comps, line);
    iconst(chunks, current, 0, line);
    lset(chunks, current, i, line);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, i, line);
    lget(chunks, current, n, line);
    op(chunks, current, Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line); // i>=n → break
    // comp = parts[i]
    lget(chunks, current, parts, line);
    lget(chunks, current, i, line);
    op(chunks, current, Op::ARRAY_GET, line);
    lset(chunks, current, comp, line);
    // if not (comp=='' or comp=='.'):
    lget(chunks, current, comp, line);
    sconst(chunks, current, "", line);
    str_eq(chunks, current, line);
    lget(chunks, current, comp, line);
    sconst(chunks, current, ".", line);
    str_eq(chunks, current, line);
    op(chunks, current, Op::I32_OR, line);
    op(chunks, current, Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    {
        // if comp != '..': comps.append(comp)
        lget(chunks, current, comp, line);
        sconst(chunks, current, "..", line);
        str_eq(chunks, current, line);
        op(chunks, current, Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        lget(chunks, current, comps, line);
        lget(chunks, current, comp, line);
        collections::emit_push(chunks, current, line);
        op(chunks, current, Op::DROP, line);
        chunks[current].emit_else(line);
        {
            // comp == '..'
            push_arr_len(chunks, current, comps, line);
            lset(chunks, current, clen, line);
            // if initslash==0 and clen==0: comps.append('..')
            lget(chunks, current, initslash, line);
            op(chunks, current, Op::I32_EQZ, line);
            lget(chunks, current, clen, line);
            op(chunks, current, Op::I32_EQZ, line);
            op(chunks, current, Op::I32_AND, line);
            chunks[current].emit_if(line);
            lget(chunks, current, comps, line);
            sconst(chunks, current, "..", line);
            collections::emit_push(chunks, current, line);
            op(chunks, current, Op::DROP, line);
            chunks[current].emit_else(line);
            {
                // elif clen>0: last=comps[-1]; if last=='..' append else pop
                lget(chunks, current, clen, line);
                iconst(chunks, current, 0, line);
                op(chunks, current, Op::I32_GT_S, line);
                chunks[current].emit_if(line);
                lget(chunks, current, comps, line);
                lget(chunks, current, clen, line);
                iconst(chunks, current, 1, line);
                op(chunks, current, Op::I32_SUB, line);
                op(chunks, current, Op::ARRAY_GET, line);
                lset(chunks, current, last, line);
                lget(chunks, current, last, line);
                sconst(chunks, current, "..", line);
                str_eq(chunks, current, line);
                chunks[current].emit_if(line);
                lget(chunks, current, comps, line);
                sconst(chunks, current, "..", line);
                collections::emit_push(chunks, current, line);
                op(chunks, current, Op::DROP, line);
                chunks[current].emit_else(line);
                lget(chunks, current, comps, line);
                collections::emit_pop(chunks, current, line);
                op(chunks, current, Op::DROP, line);
                chunks[current].emit_end(line);
                chunks[current].emit_end(line);
            }
            chunks[current].emit_end(line);
        }
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);
    // i += 1
    lget(chunks, current, i, line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_ADD, line);
    lset(chunks, current, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    // path = comps.join('/')
    lget(chunks, current, comps, line);
    sconst(chunks, current, "/", line);
    call_import(chunks, current, "ecma:array", "join", 2, line);
    lset(chunks, current, out, line);
    // if initslash: path = '/' + path
    lget(chunks, current, initslash, line);
    chunks[current].emit_if(line);
    sconst(chunks, current, "/", line);
    lget(chunks, current, out, line);
    concat2(chunks, current, line);
    lset(chunks, current, out, line);
    chunks[current].emit_end(line);
    // if path == '': path = '.'
    push_len(chunks, current, out, line);
    op(chunks, current, Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    sconst(chunks, current, ".", line);
    lset(chunks, current, out, line);
    chunks[current].emit_end(line);
}

/// Emit `parts(normpath(src_slot))` — normalize, split on `/`, drop empty
/// components — storing the resulting array in `out_slot`.
fn emit_norm_parts_to(chunks: &mut [Chunk], current: usize, src: u16, out: u16, line: u32) {
    let norm = chunks[current].alloc_scratch(1);
    emit_normpath_to(chunks, current, src, norm, line);
    emit_split_nonempty_to(chunks, current, norm, out, line);
}

/// Emit `[c for c in slot.split('/') if c != '']`, storing the array in `out`.
fn emit_split_nonempty_to(chunks: &mut [Chunk], current: usize, slot: u16, out: u16, line: u32) {
    let raw = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let k = chunks[current].alloc_scratch(1);
    let e = chunks[current].alloc_scratch(1);

    lget(chunks, current, slot, line);
    sconst(chunks, current, "/", line);
    strings::emit_split(&mut chunks[current], line);
    lset(chunks, current, raw, line);
    push_arr_len(chunks, current, raw, line);
    lset(chunks, current, n, line);
    call_import(chunks, current, "ecma:array", "new", 0, line);
    lset(chunks, current, out, line);
    iconst(chunks, current, 0, line);
    lset(chunks, current, k, line);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, k, line);
    lget(chunks, current, n, line);
    op(chunks, current, Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    lget(chunks, current, raw, line);
    lget(chunks, current, k, line);
    op(chunks, current, Op::ARRAY_GET, line);
    lset(chunks, current, e, line);
    // if e != '': out.push(e)
    lget(chunks, current, e, line);
    sconst(chunks, current, "", line);
    str_eq(chunks, current, line);
    op(chunks, current, Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    lget(chunks, current, out, line);
    lget(chunks, current, e, line);
    collections::emit_push(chunks, current, line);
    op(chunks, current, Op::DROP, line);
    chunks[current].emit_end(line);
    lget(chunks, current, k, line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_ADD, line);
    lset(chunks, current, k, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
}

// ── Public adapters ────────────────────────────────────────────────────────

/// `os.path.join(*parts)` — POSIX join. An absolute segment resets the result;
/// otherwise a single `/` separates non-empty pieces. Unrolled over `argc`.
pub fn emit_join(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        sconst(chunks, current, "", line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let result = chunks[current].alloc_scratch(1);
    lget(chunks, current, base, line);
    lset(chunks, current, result, line);

    for off in 1..argc as u16 {
        let seg = base + off;
        // if len(seg)>0 and seg[0:1]=='/': result = seg
        push_len(chunks, current, seg, line);
        iconst(chunks, current, 0, line);
        op(chunks, current, Op::I32_GT_S, line);
        push_substr(
            chunks,
            current,
            seg,
            |c, cur| iconst(c, cur, 0, line),
            |c, cur| iconst(c, cur, 1, line),
            line,
        );
        sconst(chunks, current, "/", line);
        str_eq(chunks, current, line);
        op(chunks, current, Op::I32_AND, line);
        chunks[current].emit_if(line);
        lget(chunks, current, seg, line);
        lset(chunks, current, result, line);
        chunks[current].emit_else(line);
        // elif result=='' or result[-1]=='/': result += seg
        push_len(chunks, current, result, line);
        op(chunks, current, Op::I32_EQZ, line);
        push_substr(
            chunks,
            current,
            result,
            |c, cur| {
                push_len(c, cur, result, line);
                iconst(c, cur, 1, line);
                op(c, cur, Op::I32_SUB, line);
            },
            |c, cur| push_len(c, cur, result, line),
            line,
        );
        sconst(chunks, current, "/", line);
        str_eq(chunks, current, line);
        op(chunks, current, Op::I32_OR, line);
        chunks[current].emit_if(line);
        lget(chunks, current, result, line);
        lget(chunks, current, seg, line);
        concat2(chunks, current, line);
        lset(chunks, current, result, line);
        chunks[current].emit_else(line);
        // else: result += '/' + seg
        lget(chunks, current, result, line);
        sconst(chunks, current, "/", line);
        concat2(chunks, current, line);
        lget(chunks, current, seg, line);
        concat2(chunks, current, line);
        lset(chunks, current, result, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    }
    lget(chunks, current, result, line);
}

/// `os.path.split(p)` → `(head, tail)`; `head` has trailing slashes stripped.
pub fn emit_split(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let s = base;
    let i = chunks[current].alloc_scratch(1);
    let tail = chunks[current].alloc_scratch(1);
    let head = chunks[current].alloc_scratch(1);
    // i = rfind('/') + 1
    push_rfind(chunks, current, s, "/", line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_ADD, line);
    lset(chunks, current, i, line);
    // tail = s[i:]
    push_substr(
        chunks,
        current,
        s,
        |c, cur| lget(c, cur, i, line),
        |c, cur| iconst(c, cur, END, line),
        line,
    );
    lset(chunks, current, tail, line);
    // head = strip_head(s[0:i])
    push_substr(
        chunks,
        current,
        s,
        |c, cur| iconst(c, cur, 0, line),
        |c, cur| lget(c, cur, i, line),
        line,
    );
    lset(chunks, current, head, line);
    emit_strip_head(chunks, current, head, line);
    lset(chunks, current, head, line);
    // (head, tail)
    lget(chunks, current, head, line);
    lget(chunks, current, tail, line);
    tuples::emit_tuple(chunks, current, 2, line);
}

/// `os.path.splitext(p)` → `(root, ext)`. A leading run of dots in the basename
/// does not count as an extension separator.
pub fn emit_splitext(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let s = base;
    let sep = chunks[current].alloc_scratch(1);
    let dot = chunks[current].alloc_scratch(1);
    let fi = chunks[current].alloc_scratch(1);
    let done = chunks[current].alloc_scratch(1);
    let head = chunks[current].alloc_scratch(1);
    let tail = chunks[current].alloc_scratch(1);

    push_rfind(chunks, current, s, "/", line);
    lset(chunks, current, sep, line);
    push_rfind(chunks, current, s, ".", line);
    lset(chunks, current, dot, line);
    // defaults: head = s, tail = ''
    lget(chunks, current, s, line);
    lset(chunks, current, head, line);
    sconst(chunks, current, "", line);
    lset(chunks, current, tail, line);
    iconst(chunks, current, 0, line);
    lset(chunks, current, done, line);

    // if dot > sep: scan sep+1..dot for a non-dot char
    lget(chunks, current, dot, line);
    lget(chunks, current, sep, line);
    op(chunks, current, Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    lget(chunks, current, sep, line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_ADD, line);
    lset(chunks, current, fi, line);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, done, line);
    chunks[current].emit_br_if(1, line); // found → break
    lget(chunks, current, fi, line);
    lget(chunks, current, dot, line);
    op(chunks, current, Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line); // fi>=dot → break (all dots)
    // if s[fi:fi+1] != '.': head=s[0:dot], tail=s[dot:], done=1
    push_substr(
        chunks,
        current,
        s,
        |c, cur| lget(c, cur, fi, line),
        |c, cur| {
            lget(c, cur, fi, line);
            iconst(c, cur, 1, line);
            op(c, cur, Op::I32_ADD, line);
        },
        line,
    );
    sconst(chunks, current, ".", line);
    str_eq(chunks, current, line);
    op(chunks, current, Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    push_substr(
        chunks,
        current,
        s,
        |c, cur| iconst(c, cur, 0, line),
        |c, cur| lget(c, cur, dot, line),
        line,
    );
    lset(chunks, current, head, line);
    push_substr(
        chunks,
        current,
        s,
        |c, cur| lget(c, cur, dot, line),
        |c, cur| iconst(c, cur, END, line),
        line,
    );
    lset(chunks, current, tail, line);
    iconst(chunks, current, 1, line);
    lset(chunks, current, done, line);
    chunks[current].emit_end(line);
    // fi += 1
    lget(chunks, current, fi, line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_ADD, line);
    lset(chunks, current, fi, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_end(line); // end `if dot > sep`

    lget(chunks, current, head, line);
    lget(chunks, current, tail, line);
    tuples::emit_tuple(chunks, current, 2, line);
}

/// `os.path.basename(p)` → everything after the last `/`.
pub fn emit_basename(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let s = base;
    let i = chunks[current].alloc_scratch(1);
    push_rfind(chunks, current, s, "/", line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_ADD, line);
    lset(chunks, current, i, line);
    push_substr(
        chunks,
        current,
        s,
        |c, cur| lget(c, cur, i, line),
        |c, cur| iconst(c, cur, END, line),
        line,
    );
}

/// `os.path.dirname(p)` → the leading directory (trailing slashes stripped).
pub fn emit_dirname(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let s = base;
    let i = chunks[current].alloc_scratch(1);
    let head = chunks[current].alloc_scratch(1);
    push_rfind(chunks, current, s, "/", line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_ADD, line);
    lset(chunks, current, i, line);
    push_substr(
        chunks,
        current,
        s,
        |c, cur| iconst(c, cur, 0, line),
        |c, cur| lget(c, cur, i, line),
        line,
    );
    lset(chunks, current, head, line);
    emit_strip_head(chunks, current, head, line);
}

/// `os.path.isabs(p)` → `p` starts with `/`.
pub fn emit_isabs(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let s = base;
    push_len(chunks, current, s, line);
    iconst(chunks, current, 0, line);
    op(chunks, current, Op::I32_GT_S, line);
    push_substr(
        chunks,
        current,
        s,
        |c, cur| iconst(c, cur, 0, line),
        |c, cur| iconst(c, cur, 1, line),
        line,
    );
    sconst(chunks, current, "/", line);
    str_eq(chunks, current, line);
    op(chunks, current, Op::I32_AND, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `os.path.normpath(p)`.
pub fn emit_normpath(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let out = chunks[current].alloc_scratch(1);
    emit_normpath_to(chunks, current, base, out, line);
    lget(chunks, current, out, line);
}

/// `os.path.realpath(p)` / `os.path.abspath(p)` — no FS access; normalize only.
pub fn emit_realpath(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_normpath(chunks, current, argc, line);
}

/// `os.path.normcase(s)` — POSIX identity.
pub fn emit_normcase(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    lget(chunks, current, base, line);
}

/// `os.path.expandvars(s)` — no environment substitution here; identity.
pub fn emit_expandvars(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    lget(chunks, current, base, line);
}

/// `os.path.expanduser(p)` — `~`-prefix → `/home` + remainder.
pub fn emit_expanduser(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let s = base;
    push_len(chunks, current, s, line);
    iconst(chunks, current, 0, line);
    op(chunks, current, Op::I32_GT_S, line);
    push_substr(
        chunks,
        current,
        s,
        |c, cur| iconst(c, cur, 0, line),
        |c, cur| iconst(c, cur, 1, line),
        line,
    );
    sconst(chunks, current, "~", line);
    str_eq(chunks, current, line);
    op(chunks, current, Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    sconst(chunks, current, "/home", line);
    push_substr(
        chunks,
        current,
        s,
        |c, cur| iconst(c, cur, 1, line),
        |c, cur| iconst(c, cur, END, line),
        line,
    );
    concat2(chunks, current, line);
    chunks[current].emit_else(line);
    lget(chunks, current, s, line);
    chunks[current].emit_end(line);
}

/// `os.path.islink(p)` — no symlinks in this sandbox; always `False`.
pub fn emit_islink(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    stash_args(chunks, current, 1, line);
    chunks[current].emit_bool_const(false, line);
}

/// `os.path.ismount(p)` → `normpath(p) == '/'`.
pub fn emit_ismount(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let out = chunks[current].alloc_scratch(1);
    emit_normpath_to(chunks, current, base, out, line);
    lget(chunks, current, out, line);
    sconst(chunks, current, "/", line);
    str_eq(chunks, current, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `os.path.relpath(path, start)` — relative path from `start` to `path`.
pub fn emit_relpath(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let path = base;
    let start = base + 1;
    let pp = chunks[current].alloc_scratch(1);
    let sp = chunks[current].alloc_scratch(1);
    let pn = chunks[current].alloc_scratch(1);
    let sn = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let j = chunks[current].alloc_scratch(1);
    let rel = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);

    // if path == start: return '.'
    lget(chunks, current, path, line);
    lget(chunks, current, start, line);
    str_eq(chunks, current, line);
    chunks[current].emit_if_value(line);
    sconst(chunks, current, ".", line);
    chunks[current].emit_else(line);

    emit_norm_parts_to(chunks, current, path, pp, line);
    emit_norm_parts_to(chunks, current, start, sp, line);
    push_arr_len(chunks, current, pp, line);
    lset(chunks, current, pn, line);
    push_arr_len(chunks, current, sp, line);
    lset(chunks, current, sn, line);

    // i = 0; while i<pn and i<sn and pp[i]==sp[i]: i++
    iconst(chunks, current, 0, line);
    lset(chunks, current, i, line);
    let b1 = chunks[current].emit_block(line);
    let l1 = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, i, line);
    lget(chunks, current, pn, line);
    op(chunks, current, Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    lget(chunks, current, i, line);
    lget(chunks, current, sn, line);
    op(chunks, current, Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    // pp[i]==sp[i]?
    lget(chunks, current, pp, line);
    lget(chunks, current, i, line);
    op(chunks, current, Op::ARRAY_GET, line);
    lget(chunks, current, sp, line);
    lget(chunks, current, i, line);
    op(chunks, current, Op::ARRAY_GET, line);
    str_eq(chunks, current, line);
    op(chunks, current, Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    lget(chunks, current, i, line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_ADD, line);
    lset(chunks, current, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(l1);
    chunks[current].emit_end(line);
    chunks[current].patch_block(b1);

    // rel = []; j=i; while j<sn: rel.append('..'); j++
    call_import(chunks, current, "ecma:array", "new", 0, line);
    lset(chunks, current, rel, line);
    lget(chunks, current, i, line);
    lset(chunks, current, j, line);
    let b2 = chunks[current].emit_block(line);
    let l2 = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, j, line);
    lget(chunks, current, sn, line);
    op(chunks, current, Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    lget(chunks, current, rel, line);
    sconst(chunks, current, "..", line);
    collections::emit_push(chunks, current, line);
    op(chunks, current, Op::DROP, line);
    lget(chunks, current, j, line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_ADD, line);
    lset(chunks, current, j, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(l2);
    chunks[current].emit_end(line);
    chunks[current].patch_block(b2);

    // j=i; while j<pn: rel.append(pp[j]); j++
    lget(chunks, current, i, line);
    lset(chunks, current, j, line);
    let b3 = chunks[current].emit_block(line);
    let l3 = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, j, line);
    lget(chunks, current, pn, line);
    op(chunks, current, Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    lget(chunks, current, rel, line);
    lget(chunks, current, pp, line);
    lget(chunks, current, j, line);
    op(chunks, current, Op::ARRAY_GET, line);
    collections::emit_push(chunks, current, line);
    op(chunks, current, Op::DROP, line);
    lget(chunks, current, j, line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_ADD, line);
    lset(chunks, current, j, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(l3);
    chunks[current].emit_end(line);
    chunks[current].patch_block(b3);

    // if len(rel)==0: '.' else '/'.join(rel)
    push_arr_len(chunks, current, rel, line);
    op(chunks, current, Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    sconst(chunks, current, ".", line);
    lset(chunks, current, out, line);
    chunks[current].emit_else(line);
    lget(chunks, current, rel, line);
    sconst(chunks, current, "/", line);
    call_import(chunks, current, "ecma:array", "join", 2, line);
    lset(chunks, current, out, line);
    chunks[current].emit_end(line);
    lget(chunks, current, out, line);

    chunks[current].emit_end(line); // end `if path == start`
}

/// `os.path.commonprefix(items)` — longest common *string* prefix (char-wise).
pub fn emit_commonprefix(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let items = base;
    let m = chunks[current].alloc_scratch(1);
    let first = chunks[current].alloc_scratch(1);
    let minlen = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let k = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    let ch = chunks[current].alloc_scratch(1);
    let it = chunks[current].alloc_scratch(1);
    let same = chunks[current].alloc_scratch(1);

    // result = ''
    sconst(chunks, current, "", line);
    lset(chunks, current, result, line);
    push_arr_len(chunks, current, items, line);
    lset(chunks, current, m, line);
    // if m>0:
    lget(chunks, current, m, line);
    iconst(chunks, current, 0, line);
    op(chunks, current, Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    // first = items[0]; minlen = len(first)
    lget(chunks, current, items, line);
    iconst(chunks, current, 0, line);
    op(chunks, current, Op::ARRAY_GET, line);
    lset(chunks, current, first, line);
    push_len(chunks, current, first, line);
    lset(chunks, current, minlen, line);
    // for it in items: minlen = min(minlen, len(it))
    iconst(chunks, current, 0, line);
    lset(chunks, current, k, line);
    let bl = chunks[current].emit_block(line);
    let ll = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, k, line);
    lget(chunks, current, m, line);
    op(chunks, current, Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    lget(chunks, current, items, line);
    lget(chunks, current, k, line);
    op(chunks, current, Op::ARRAY_GET, line);
    lset(chunks, current, it, line);
    push_len(chunks, current, it, line);
    lget(chunks, current, minlen, line);
    op(chunks, current, Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    push_len(chunks, current, it, line);
    lset(chunks, current, minlen, line);
    chunks[current].emit_end(line);
    lget(chunks, current, k, line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_ADD, line);
    lset(chunks, current, k, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(ll);
    chunks[current].emit_end(line);
    chunks[current].patch_block(bl);

    // i=0; while i<minlen: ch=first[i]; same over items; if diff break else result+=ch
    iconst(chunks, current, 0, line);
    lset(chunks, current, i, line);
    let b2 = chunks[current].emit_block(line);
    let l2 = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, i, line);
    lget(chunks, current, minlen, line);
    op(chunks, current, Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    // ch = first[i:i+1]
    push_substr(
        chunks,
        current,
        first,
        |c, cur| lget(c, cur, i, line),
        |c, cur| {
            lget(c, cur, i, line);
            iconst(c, cur, 1, line);
            op(c, cur, Op::I32_ADD, line);
        },
        line,
    );
    lset(chunks, current, ch, line);
    // same = 1; for it in items: if it[i:i+1]!=ch: same=0
    iconst(chunks, current, 1, line);
    lset(chunks, current, same, line);
    iconst(chunks, current, 0, line);
    lset(chunks, current, k, line);
    let b3 = chunks[current].emit_block(line);
    let l3 = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, k, line);
    lget(chunks, current, m, line);
    op(chunks, current, Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    lget(chunks, current, items, line);
    lget(chunks, current, k, line);
    op(chunks, current, Op::ARRAY_GET, line);
    lset(chunks, current, it, line);
    push_substr(
        chunks,
        current,
        it,
        |c, cur| lget(c, cur, i, line),
        |c, cur| {
            lget(c, cur, i, line);
            iconst(c, cur, 1, line);
            op(c, cur, Op::I32_ADD, line);
        },
        line,
    );
    lget(chunks, current, ch, line);
    str_eq(chunks, current, line);
    op(chunks, current, Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    iconst(chunks, current, 0, line);
    lset(chunks, current, same, line);
    chunks[current].emit_end(line);
    lget(chunks, current, k, line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_ADD, line);
    lset(chunks, current, k, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(l3);
    chunks[current].emit_end(line);
    chunks[current].patch_block(b3);
    // if not same: break; else result+=ch; i++
    lget(chunks, current, same, line);
    op(chunks, current, Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    lget(chunks, current, result, line);
    lget(chunks, current, ch, line);
    concat2(chunks, current, line);
    lset(chunks, current, result, line);
    lget(chunks, current, i, line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_ADD, line);
    lset(chunks, current, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(l2);
    chunks[current].emit_end(line);
    chunks[current].patch_block(b2);

    chunks[current].emit_end(line); // end `if m>0`
    lget(chunks, current, result, line);
}

/// `os.path.commonpath(items)` — longest common *path* prefix (component-wise).
pub fn emit_commonpath(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let items = base;
    let m = chunks[current].alloc_scratch(1);
    let absolute = chunks[current].alloc_scratch(1);
    let first0 = chunks[current].alloc_scratch(1);
    let common = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let keep = chunks[current].alloc_scratch(1);
    let c0 = chunks[current].alloc_scratch(1);
    let allmatch = chunks[current].alloc_scratch(1);
    let k = chunks[current].alloc_scratch(1);
    let pl = chunks[current].alloc_scratch(1);
    let it = chunks[current].alloc_scratch(1);
    let lists = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);

    push_arr_len(chunks, current, items, line);
    lset(chunks, current, m, line);
    // if m==0: '' else compute
    lget(chunks, current, m, line);
    op(chunks, current, Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    sconst(chunks, current, "", line);
    chunks[current].emit_else(line);

    // absolute = items[0][0:1] == '/'
    lget(chunks, current, items, line);
    iconst(chunks, current, 0, line);
    op(chunks, current, Op::ARRAY_GET, line);
    lset(chunks, current, first0, line);
    push_substr(
        chunks,
        current,
        first0,
        |c, cur| iconst(c, cur, 0, line),
        |c, cur| iconst(c, cur, 1, line),
        line,
    );
    sconst(chunks, current, "/", line);
    str_eq(chunks, current, line);
    lset(chunks, current, absolute, line);

    // lists = [ parts(it) for it in items ]  (split on '/', drop empties)
    call_import(chunks, current, "ecma:array", "new", 0, line);
    lset(chunks, current, lists, line);
    iconst(chunks, current, 0, line);
    lset(chunks, current, k, line);
    let bl = chunks[current].emit_block(line);
    let ll = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, k, line);
    lget(chunks, current, m, line);
    op(chunks, current, Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    lget(chunks, current, items, line);
    lget(chunks, current, k, line);
    op(chunks, current, Op::ARRAY_GET, line);
    lset(chunks, current, it, line);
    emit_split_nonempty_to(chunks, current, it, pl, line);
    lget(chunks, current, lists, line);
    lget(chunks, current, pl, line);
    collections::emit_push(chunks, current, line);
    op(chunks, current, Op::DROP, line);
    lget(chunks, current, k, line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_ADD, line);
    lset(chunks, current, k, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(ll);
    chunks[current].emit_end(line);
    chunks[current].patch_block(bl);

    // common = []; idx=0; keep=1
    call_import(chunks, current, "ecma:array", "new", 0, line);
    lset(chunks, current, common, line);
    iconst(chunks, current, 0, line);
    lset(chunks, current, idx, line);
    iconst(chunks, current, 1, line);
    lset(chunks, current, keep, line);

    let b2 = chunks[current].emit_block(line);
    let l2 = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, keep, line);
    op(chunks, current, Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line); // keep==0 → break
    // if idx >= len(lists[0]): keep=0
    lget(chunks, current, lists, line);
    iconst(chunks, current, 0, line);
    op(chunks, current, Op::ARRAY_GET, line);
    lset(chunks, current, pl, line); // pl = lists[0]
    lget(chunks, current, idx, line);
    push_arr_len(chunks, current, pl, line);
    op(chunks, current, Op::I32_GE_S, line);
    chunks[current].emit_if(line);
    iconst(chunks, current, 0, line);
    lset(chunks, current, keep, line);
    chunks[current].emit_else(line);
    // c0 = lists[0][idx]; allmatch=1
    lget(chunks, current, pl, line);
    lget(chunks, current, idx, line);
    op(chunks, current, Op::ARRAY_GET, line);
    lset(chunks, current, c0, line);
    iconst(chunks, current, 1, line);
    lset(chunks, current, allmatch, line);
    // for pl in lists: if idx>=len(pl) or pl[idx]!=c0: allmatch=0
    iconst(chunks, current, 0, line);
    lset(chunks, current, k, line);
    let b3 = chunks[current].emit_block(line);
    let l3 = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, k, line);
    lget(chunks, current, m, line);
    op(chunks, current, Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    lget(chunks, current, lists, line);
    lget(chunks, current, k, line);
    op(chunks, current, Op::ARRAY_GET, line);
    lset(chunks, current, pl, line);
    // cond = idx>=len(pl) OR pl[idx]!=c0
    lget(chunks, current, idx, line);
    push_arr_len(chunks, current, pl, line);
    op(chunks, current, Op::I32_GE_S, line);
    chunks[current].emit_if_value(line);
    iconst(chunks, current, 1, line);
    chunks[current].emit_else(line);
    lget(chunks, current, pl, line);
    lget(chunks, current, idx, line);
    op(chunks, current, Op::ARRAY_GET, line);
    lget(chunks, current, c0, line);
    str_eq(chunks, current, line);
    op(chunks, current, Op::I32_EQZ, line);
    chunks[current].emit_end(line);
    chunks[current].emit_if(line);
    iconst(chunks, current, 0, line);
    lset(chunks, current, allmatch, line);
    chunks[current].emit_end(line);
    lget(chunks, current, k, line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_ADD, line);
    lset(chunks, current, k, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(l3);
    chunks[current].emit_end(line);
    chunks[current].patch_block(b3);
    // if allmatch: common.append(c0); idx++ else keep=0
    lget(chunks, current, allmatch, line);
    chunks[current].emit_if(line);
    lget(chunks, current, common, line);
    lget(chunks, current, c0, line);
    collections::emit_push(chunks, current, line);
    op(chunks, current, Op::DROP, line);
    lget(chunks, current, idx, line);
    iconst(chunks, current, 1, line);
    op(chunks, current, Op::I32_ADD, line);
    lset(chunks, current, idx, line);
    chunks[current].emit_else(line);
    iconst(chunks, current, 0, line);
    lset(chunks, current, keep, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line); // end `if idx>=len(lists[0])`
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(l2);
    chunks[current].emit_end(line);
    chunks[current].patch_block(b2);

    // out = '/'.join(common); if absolute: out = '/' + out
    lget(chunks, current, common, line);
    sconst(chunks, current, "/", line);
    call_import(chunks, current, "ecma:array", "join", 2, line);
    lset(chunks, current, out, line);
    lget(chunks, current, absolute, line);
    chunks[current].emit_if(line);
    sconst(chunks, current, "/", line);
    lget(chunks, current, out, line);
    concat2(chunks, current, line);
    lset(chunks, current, out, line);
    chunks[current].emit_end(line);
    lget(chunks, current, out, line);

    chunks[current].emit_end(line); // end `if m==0`
}
