//! Python-specific string instance methods.
//!
//! These compose Python surface semantics on top of `ecma:string` primitives
//! (never mutating the shared ECMA host into a Python shape). Routed via
//! `common:python.str_*` from the Python profile `[value_methods]` table.
//!
//! The receiver and explicit arguments arrive pre-pushed on the stack (receiver
//! first), matching the `emit_common` value-method calling convention.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use vybe_compiler::primitives::{collections, ops, strings};

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    // Register on the CURRENT chunk (not chunks[0]) so `normalize_import_table`
    // remaps this CALL_IMPORT via the emitting chunk's own local table. A
    // chunks[0] index inside a non-root chunk collides with per-chunk imports
    // and resolves the wrong host fn. Matches shared `emit_import_call`.
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(argc, line);
}

/// Pop `argc` stack values (deepest first) into freshly allocated scratch slots
/// and return the base slot. `base + 0` is the receiver, `base + i` the i-th arg.
fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn push_i32(chunk: &mut Chunk, v: i32, line: u32) {
    chunk.emit_i32_const(v, line);
}

fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

/// `str.casefold()` — aggressive lowercase; `toLowerCase` covers the cases the
/// suite exercises. Receiver already on the stack.
pub fn emit_casefold(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "ecma:string", "toLowerCase", 1, line);
}

/// `s.removeprefix(p)` → `s[len(p):]` when `s.startswith(p)`, else `s`.
pub fn emit_removeprefix(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let s = base;
    let p = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
    call_import(chunks, current, "ecma:string", "startsWith", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    // s[len(p):]
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_i32_const(0x7FFF_FFFF, line);
    strings::emit_substring(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_end(line);
}

/// `s.removesuffix(x)` → `s[:len(s)-len(x)]` when `s.endswith(x)`, else `s`.
pub fn emit_removesuffix(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let s = base;
    let x = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    call_import(chunks, current, "ecma:string", "endsWith", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    // s[0 : len(s) - len(x)]
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_SUB, line);
    strings::emit_substring(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_end(line);
}

/// Python `s.expandtabs(tabsize=8)` — replace each tab with spaces up to the
/// next multiple of `tabsize`, tracking the column so alignment is correct, and
/// resetting the column on `\n`/`\r`. Builds the result by scanning one code
/// unit at a time.
pub fn emit_expandtabs(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let s = base;

    // tabsize (default 8); `repeat`/arithmetic coerce boxed args, but the column
    // math is i32, so unbox an explicit argument.
    let ts = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
        call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
        chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    } else {
        chunks[current].emit_i32_const(8, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, ts, line);

    let result = chunks[current].alloc_scratch(1);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    let col = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, col, line);
    let i = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    let ch = chunks[current].alloc_scratch(1);
    let pad = chunks[current].alloc_scratch(1);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    // break when i >= n
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    // ch = s[i:i+1]
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    strings::emit_substring(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch, line);
    // if ch == "\t"
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    chunks[current].emit_string_const("\t", line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
    chunks[current].emit_if(line);
    // pad = ts - (col % ts)
    chunks[current].emit_op_u16(Op::LOCAL_GET, ts, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, col, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ts, line);
    chunks[current].emit_op(Op::I32_REM_S, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pad, line);
    // result += " " * pad
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_string_const(" ", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pad, line);
    call_import(chunks, current, "ecma:string", "repeat", 2, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    // col += pad
    chunks[current].emit_op_u16(Op::LOCAL_GET, col, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pad, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, col, line);
    chunks[current].emit_else(line);
    // result += ch
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    // col = (ch == "\n" || ch == "\r") ? 0 : col + 1
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    chunks[current].emit_string_const("\n", line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    chunks[current].emit_string_const("\r", line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, col, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, col, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, col, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    // i++
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}

/// Python `s.strip([chars])` — with no argument trims surrounding whitespace;
/// with `chars`, removes any leading/trailing character contained in the set.
/// The two edges are trimmed by symmetric `while` loops that peel one code unit
/// at a time while it is a member of `chars` (`ecma:string.includes`).
pub fn emit_split(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let s = base;

    if argc <= 1 {
        // trimmed = s.trim(); an empty/all-whitespace input yields `[]`
        // (`ecma:regexp.split` on `''` would give `['']`), otherwise split on
        // `\s+`, which collapses interior whitespace runs to single boundaries.
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        call_import(chunks, current, "ecma:string", "trim", 1, line);
        let trimmed = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, trimmed, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, trimmed, line);
        strings::emit_length(&mut chunks[current], line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, trimmed, line);
        chunks[current].emit_string_const("\\s+", line);
        call_import(chunks, current, "ecma:regexp", "split", 2, line);
        chunks[current].emit_else(line);
        call_import(chunks, current, "ecma:array", "new", 0, line);
        chunks[current].emit_end(line);
        return;
    }

    let sep = base + 1;
    if argc == 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, sep, line);
        call_import(chunks, current, "ecma:string", "split", 2, line);
        return;
    }

    // argc >= 3: maxsplit — split fully, then re-join the tail beyond `n`.
    let n = base + 2;
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep, line);
    call_import(chunks, current, "ecma:string", "split", 2, line);
    let full = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, full, line);

    // head = full.slice(0, n)
    chunks[current].emit_op_u16(Op::LOCAL_GET, full, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    call_import(chunks, current, "ecma:array", "slice", 3, line);
    let head = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, head, line);

    // tail = full.slice(n)
    chunks[current].emit_op_u16(Op::LOCAL_GET, full, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_i32_const(0x7FFF_FFFF, line);
    call_import(chunks, current, "ecma:array", "slice", 3, line);
    let tail = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, tail, line);

    // if tail non-empty: head.push(tail.join(sep))
    chunks[current].emit_op_u16(Op::LOCAL_GET, tail, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, head, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tail, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep, line);
    call_import(chunks, current, "ecma:array", "join", 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, head, line);
}

/// Python `s.startswith(prefix[, start[, end]])` — test whether the slice
/// `s[start:end]` begins with `prefix`. `substring` (`wasm:js-string`) coerces
/// boxed offset arguments to i32 via `as_i32`, so they pass through unmodified.
/// The host `startsWith` yields a boolean that Python renders as `True`/`False`.
pub fn emit_startswith(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let s = base;
    let prefix = base + 1;

    // sub = s[start:end]
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 2, line);
    } else {
        chunks[current].emit_i32_const(0, line);
    }
    if argc >= 4 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 3, line);
    } else {
        chunks[current].emit_i32_const(0x7FFF_FFFF, line);
    }
    strings::emit_substring(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, prefix, line);
    call_import(chunks, current, "ecma:string", "startsWith", 2, line);
}

/// Python `s.endswith(suffix[, start[, end]])` — test whether the slice
/// `s[start:end]` ends with `suffix`.
pub fn emit_endswith(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let s = base;
    let suffix = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 2, line);
    } else {
        chunks[current].emit_i32_const(0, line);
    }
    if argc >= 4 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 3, line);
    } else {
        chunks[current].emit_i32_const(0x7FFF_FFFF, line);
    }
    strings::emit_substring(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, suffix, line);
    call_import(chunks, current, "ecma:string", "endsWith", 2, line);
}

/// Python `s.count(sub)` — number of non-overlapping occurrences, via
/// `len(s.split(sub)) - 1`. `ecma:array.length` already returns a raw `i32`, so
/// the count is computed directly with no boxed-number round-trip.
pub fn emit_count(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let s = base;
    let sub = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sub, line);
    call_import(chunks, current, "ecma:string", "split", 2, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
}

/// Python `s.replace(old, new[, count])`: no `count` → replace ALL
/// (`ecma:string.replaceAll`); with `count`, replace the first `count`
/// occurrences by peeling one leftmost match per iteration.
pub fn emit_replace(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // argc = receiver + old + new (+ count) → 3 or 4.
    let base = stash_args(chunks, current, argc, line);
    let s = base;
    let old = base + 1;
    let new = base + 2;

    if argc < 4 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, old, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, new, line);
        call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
        return;
    }

    // cnt = int(count)
    let cnt = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 3, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cnt, line);
    let k = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, k, line);
    let idx = chunks[current].alloc_scratch(1);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    // break when k >= cnt
    chunks[current].emit_op_u16(Op::LOCAL_GET, k, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cnt, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    // idx = s.indexOf(old)  (indexOf returns a boxed number)
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, old, line);
    strings::emit_index_of(&mut chunks[current], line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);
    // break when idx < 0 (no more matches)
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_br_if(1, line);
    // s = s[:idx] + new + s[idx+len(old):]
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    strings::emit_substring(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, new, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, old, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_i32_const(0x7FFF_FFFF, line);
    strings::emit_substring(&mut chunks[current], line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
    // k++
    chunks[current].emit_op_u16(Op::LOCAL_GET, k, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, k, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
}

/// Python `s.rsplit(sep[, maxsplit])` — like `split`, but when `maxsplit` limits
/// the count the splits are taken from the RIGHT (the head stays joined). Split
/// fully, keep the last `maxsplit` pieces, and re-join everything before them.
pub fn emit_rsplit(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let s = base;

    // No maxsplit → identical to a plain split.
    if argc < 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        if argc == 2 {
            chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
            call_import(chunks, current, "ecma:string", "split", 2, line);
        } else {
            call_import(chunks, current, "ecma:string", "trim", 1, line);
            chunks[current].emit_string_const("\\s+", line);
            call_import(chunks, current, "ecma:regexp", "split", 2, line);
        }
        return;
    }

    let sep = base + 1;
    // full = s.split(sep)
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep, line);
    call_import(chunks, current, "ecma:string", "split", 2, line);
    let full = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, full, line);

    // split point = max(0, len(full) - maxsplit)
    let sp = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, full, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 2, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sp, line);
    // clamp negatives to 0: sp = sp * (sp > 0)
    chunks[current].emit_op_u16(Op::LOCAL_GET, sp, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sp, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_op(Op::I32_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sp, line);

    // tail = full.slice(sp)  (the last maxsplit pieces, kept whole)
    let tail = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, full, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sp, line);
    chunks[current].emit_i32_const(0x7FFF_FFFF, line);
    call_import(chunks, current, "ecma:array", "slice", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, tail, line);

    // if sp > 0: result = concat([full.slice(0,sp).join(sep)], tail); else full
    let hd = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sp, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    // hd = [full.slice(0,sp).join(sep)]  — stash the array so `push` doesn't
    // consume the only reference before `concat` reads it.
    call_import(chunks, current, "ecma:array", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, hd, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, hd, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, full, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sp, line);
    call_import(chunks, current, "ecma:array", "slice", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep, line);
    call_import(chunks, current, "ecma:array", "join", 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, hd, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tail, line);
    call_import(chunks, current, "ecma:array", "concat", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, full, line);
    chunks[current].emit_end(line);
}

/// Python `s.splitlines([keepends])` — split on line boundaries (`\n`, `\r`,
/// `\r\n`). Terminators are stripped unless `keepends` is truthy, and a trailing
/// terminator does not produce a final empty element. Scans one code unit at a
/// time so `\r\n` counts as a single boundary.
pub fn emit_splitlines(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let s = base;

    let keepends = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
    } else {
        chunks[current].emit_i32_const(0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, keepends, line);

    let result = chunks[current].alloc_scratch(1);
    call_import(chunks, current, "ecma:array", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    let i = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let start = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
    let ch = chunks[current].alloc_scratch(1);
    let termlen = chunks[current].alloc_scratch(1);
    let lineend = chunks[current].alloc_scratch(1);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    // break when i >= n
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    // ch = s[i:i+1]
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    strings::emit_substring(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch, line);
    // nl_or_cr = (ch == "\n") | (ch == "\r")
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    chunks[current].emit_string_const("\n", line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    chunks[current].emit_string_const("\r", line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    // termlen = 1 + (ch == "\r" && s[i+1:i+2] == "\n")  → \r\n is one boundary
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    chunks[current].emit_string_const("\r", line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    strings::emit_substring(&mut chunks[current], line);
    chunks[current].emit_string_const("\n", line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, termlen, line);
    // lineend = i + keepends * termlen
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keepends, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, termlen, line);
    chunks[current].emit_op(Op::I32_MUL, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lineend, line);
    // result.push(s[start:lineend])
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lineend, line);
    strings::emit_substring(&mut chunks[current], line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // i += termlen; start = i
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, termlen, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
    chunks[current].emit_else(line);
    // i += 1
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_end(line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    // trailing content with no final terminator → one last line
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    strings::emit_substring(&mut chunks[current], line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}

/// `str.maketrans(x[, y[, z]])` → a Map from single-character key to its
/// replacement (a string, or null to delete the character).
///
/// Three forms, all of which the suite uses:
///   * `maketrans(from, to)`      — positional character pairing
///   * `maketrans(from, to, del)` — plus characters mapped to null (deleted)
///   * `maketrans(dict)`          — `{ord('a'): 'AA', ord('b'): None}`; integer
///     keys are code points, so they are converted back to characters here and
///     `translate` can stay a plain per-character lookup.
pub fn emit_maketrans(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let table = chunks[current].alloc_scratch(1);
    call_import(chunks, current, "ecma:map", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, table, line);

    if argc == 1 {
        // Dict form: re-key `ord(c) -> repl` as `c -> repl`.
        let entries = chunks[current].alloc_scratch(1);
        let i = chunks[current].alloc_scratch(1);
        let n = chunks[current].alloc_scratch(1);
        let pair = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
        call_import(chunks, current, "ecma:object", "entries", 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, entries, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, entries, line);
        chunks[current].emit_op(Op::ARRAY_LENGTH, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

        let block = chunks[current].emit_block(line);
        let lp = chunks[current].emit_loop_s(line).0;
        chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_br_if(1, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, entries, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, pair, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, table, line);
        // key: fromCodePoint(Number(pair[0]))
        chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
        call_import(chunks, current, "ecma:string", "fromCodePoint", 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        call_import(chunks, current, "ecma:map", "set", 3, line);
        chunks[current].emit_op(Op::DROP, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].patch_loop(lp);
        chunks[current].emit_end(line);
        chunks[current].patch_block(block);
    } else {
        // from/to pairing, then the optional delete set.
        emit_pair_chars(chunks, current, table, base, base + 1, false, line);
        if argc >= 3 {
            emit_pair_chars(chunks, current, table, base + 2, base + 2, true, line);
        }
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, table, line);
}

/// Map `src[i] -> dst[i]` for every index of `src`. When `delete` is set the
/// value stored is null, which `translate` treats as "drop this character".
fn emit_pair_chars(
    chunks: &mut [Chunk],
    current: usize,
    table: u16,
    src: u16,
    dst: u16,
    delete: bool,
    line: u32,
) {
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, table, line);
    emit_char_at(chunks, current, src, i, line);
    if delete {
        let k = chunks[current].add_constant(vybe_runtime::Value::Null);
        chunks[current].emit_op_u16(Op::CONST, k, line);
    } else {
        emit_char_at(chunks, current, dst, i, line);
    }
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
}

/// Push `s[i:i+1]`.
fn emit_char_at(chunks: &mut [Chunk], current: usize, s: u16, i: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    strings::emit_substring(&mut chunks[current], line);
}

/// `s.translate(table)` — per-character lookup in the `maketrans` Map. A missing
/// key keeps the character; a null value deletes it.
pub fn emit_translate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let s = base;
    let table = base + 1;

    let result = chunks[current].alloc_scratch(1);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    let i = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    let ch = chunks[current].alloc_scratch(1);
    let repl = chunks[current].alloc_scratch(1);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    emit_char_at(chunks, current, s, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, table, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    call_import(chunks, current, "ecma:map", "has", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, table, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    call_import(chunks, current, "ecma:map", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, repl, line);
    // null → delete the character.
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}

/// `str.istitle()` — the string equals its title-cased form AND contains at
/// least one cased character (so `"123".istitle()` is False even though it is
/// trivially equal to its own title-case).
pub fn emit_istitle(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);

    // s == title(s)
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    emit_title_of(chunks, current, s, line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);

    // ... and the string has a cased character: upper(s) != lower(s).
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    call_import(chunks, current, "ecma:string", "toUpperCase", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    call_import(chunks, current, "ecma:string", "toLowerCase", 1, line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
    chunks[current].emit_op(Op::I32_EQZ, line);

    chunks[current].emit_op(Op::I32_AND, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Push the title-cased form of the string in `src`.
fn emit_title_of(chunks: &mut [Chunk], current: usize, src: u16, line: u32) {
    let result = chunks[current].alloc_scratch(1);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    let i = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    let ch = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    emit_char_at(chunks, current, src, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch, line);

    // A character is "cased" when upper and lower differ.
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    call_import(chunks, current, "ecma:string", "toUpperCase", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    call_import(chunks, current, "ecma:string", "toLowerCase", 1, line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
    chunks[current].emit_op(Op::I32_EQZ, line); // 1 when cased
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    call_import(chunks, current, "ecma:string", "toUpperCase", 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    call_import(chunks, current, "ecma:string", "toLowerCase", 1, line);
    chunks[current].emit_end(line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}

// ── str.find / str.rfind / str.rindex ──────────────────────────────────
//
// Python string indices count CODE POINTS (`unifiedstringplan.md` Axis 1,
// Python's unit is `scalar`), so these must agree with `len`, `s[i]` and
// `s[a:b]`. They were bound straight to `common:strings.index_of` /
// `common:str_last_index_of`, which are UTF-16 — `"a😀b".find("b")` answered
// 3 instead of 2.
//
// Those raw bindings also ignored `argc`: `str.find(sub, start[, end])` left
// the extra arguments on the stack, so `"abcabc".find("b", 2)` produced
// `-1abcabc` (stack garbage) rather than 4. Stashing by `argc` fixes that as a
// side effect of needing the bounds at all.
//
// `end` defaults to the code-point length, `start` to 0; both wrap from the end
// and clamp, as `s[start:end]` does.

/// Python `str.find` / `str.rfind` / `str.index` / `str.rindex` on a string.
/// Stack: `[s, sub]`, `[s, sub, start]` or `[s, sub, start, end]` → `[i32]`.
/// `last` picks the final occurrence; `raises` turns the `-1` miss into
/// `ValueError`, which is the only difference between `find` and `index`.
pub fn emit_str_search(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
    last: bool,
    raises: bool,
) {
    let chunk = &mut chunks[current];
    let s_slot = alloc_local(chunk);
    let sub_slot = alloc_local(chunk);
    let start_slot = alloc_local(chunk);
    let end_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let idx_slot = alloc_local(chunk);

    // Stack (bottom→top): s, sub, [start, [end]] — pop in reverse.
    if argc >= 4 {
        vybe_compiler::primitives::convert::emit_to_int(chunk, line);
        lset(chunk, end_slot, line);
    }
    if argc >= 3 {
        vybe_compiler::primitives::convert::emit_to_int(chunk, line);
        lset(chunk, start_slot, line);
    }
    lset(chunk, sub_slot, line);
    lset(chunk, s_slot, line);
    if argc < 3 {
        push_i32(chunk, 0, line);
        lset(chunk, start_slot, line);
    }

    lget(chunk, s_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::strings::emit_scalar_length(chunks, current, line);
    let chunk = &mut chunks[current];
    lset(chunk, len_slot, line);
    if argc < 4 {
        lget(chunk, len_slot, line);
        lset(chunk, end_slot, line);
    }
    clamp_index(chunk, start_slot, len_slot, line);
    clamp_index(chunk, end_slot, len_slot, line);

    // Search the window, then shift the answer back into whole-string space.
    lget(chunk, s_slot, line);
    lget(chunk, start_slot, line);
    lget(chunk, end_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::strings::emit_scalar_substring(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, sub_slot, line);
    let _ = chunk;
    if last {
        vybe_compiler::primitives::strings::emit_scalar_last_index_of(chunks, current, line);
    } else {
        vybe_compiler::primitives::strings::emit_scalar_index_of(chunks, current, line);
    }
    let chunk = &mut chunks[current];
    lset(chunk, idx_slot, line);

    lget(chunk, idx_slot, line);
    push_i32(chunk, 0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    if raises {
        chunk.emit_if(line);
        // `emit_exception_new_finalize` wants `[obj, obj, msg]`.
        chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
        chunk.emit_dup(line);
        chunk.emit_string_const("substring not found", line);
        vybe_compiler::primitives::errors::emit_exception_new_finalize(chunk, "ValueError", line);
        vybe_compiler::primitives::errors::emit_throw(chunk, line);
        chunk.emit_end(line);
        lget(chunk, idx_slot, line);
        lget(chunk, start_slot, line);
        chunk.emit_op(Op::I32_ADD, line);
    } else {
        chunk.emit_if_value(line);
        push_i32(chunk, -1, line);
        chunk.emit_else(line);
        lget(chunk, idx_slot, line);
        lget(chunk, start_slot, line);
        chunk.emit_op(Op::I32_ADD, line);
        chunk.emit_end(line);
    }
}

/// `slot < 0 → slot += len`, then clamp into `[0, len]` — Python slice bounds.
fn clamp_index(chunk: &mut Chunk, slot: u16, len_slot: u16, line: u32) {
    lget(chunk, slot, line);
    push_i32(chunk, 0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, len_slot, line);
    lget(chunk, slot, line);
    chunk.emit_op(Op::I32_ADD, line);
    lset(chunk, slot, line);
    chunk.emit_end(line);
    lget(chunk, slot, line);
    push_i32(chunk, 0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_i32(chunk, 0, line);
    lset(chunk, slot, line);
    chunk.emit_end(line);
    lget(chunk, slot, line);
    lget(chunk, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, len_slot, line);
    lset(chunk, slot, line);
    chunk.emit_end(line);
}
