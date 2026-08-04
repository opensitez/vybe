//! Kotlin's own string extensions (`kotlin.text`) — the surface the JDK does
//! not have: `isBlank`, `removePrefix`, `substringAfter`, `trimIndent`,
//! `lines`, the `OrNull` parsers, and friends.
//!
//! JDK spellings (`java.lang.StringBuilder`, `String.compareTo`) live in
//! `platforms/jvm`; these are Kotlin-only names, so they live with the
//! language. Everything composes `ecma:string` host fns and the shared
//! `strings`/`collections` primitives.

use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::{collections, ops, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn get(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn truthy(chunks: &mut [Chunk], current: usize, line: u32) {
    ops::emit_dyn_to_bool(&mut chunks[current], line);
}

fn host(chunks: &mut [Chunk], current: usize, module: &str, func: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import(module, func);
    chunks[current].emit_call(idx, argc, line);
}

fn bool_out(chunks: &mut [Chunk], current: usize, line: u32) {
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `isBlank()` — trimmed length is zero.
pub fn emit_is_blank(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    host(chunks, current, "ecma:string", "trim", 1, line);
    strings::emit_length(&mut chunks[current], line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_eq(&mut chunks[current], line);
    bool_out(chunks, current, line);
}

pub fn emit_is_not_blank(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_is_blank(chunks, current, argc, line);
    ops::emit_dyn_not(&mut chunks[current], line);
    bool_out(chunks, current, line);
}

/// `isNullOrEmpty()` / `isNullOrBlank()`.
pub fn emit_is_null_or_empty(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_null_or(chunks, current, /*blank:*/ false, line);
}

pub fn emit_is_null_or_blank(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_null_or(chunks, current, /*blank:*/ true, line);
}

fn emit_null_or(chunks: &mut Vec<Chunk>, current: usize, blank: bool, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    set(chunks, current, v, line);
    get(chunks, current, v, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    bool_out(chunks, current, line);
    chunks[current].emit_else(line);
    get(chunks, current, v, line);
    if blank {
        emit_is_blank(chunks, current, 1, line);
    } else {
        strings::emit_length(&mut chunks[current], line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        ops::emit_dyn_eq(&mut chunks[current], line);
        bool_out(chunks, current, line);
    }
    chunks[current].emit_end(line);
}

/// `removePrefix(p)` / `removeSuffix(s)` — strip only when present.
pub fn emit_remove_prefix(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_remove_edge(chunks, current, /*prefix:*/ true, line);
}

pub fn emit_remove_suffix(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_remove_edge(chunks, current, /*prefix:*/ false, line);
}

fn emit_remove_edge(chunks: &mut Vec<Chunk>, current: usize, prefix: bool, line: u32) {
    let p = chunks[current].alloc_scratch(1);
    let s = chunks[current].alloc_scratch(1);
    set(chunks, current, p, line);
    set(chunks, current, s, line);
    get(chunks, current, s, line);
    get(chunks, current, p, line);
    host(
        chunks,
        current,
        "ecma:string",
        if prefix { "startsWith" } else { "endsWith" },
        2,
        line,
    );
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    if prefix {
        get(chunks, current, s, line);
        get(chunks, current, p, line);
        strings::emit_length(&mut chunks[current], line);
        host(chunks, current, "ecma:string", "slice", 2, line);
    } else {
        get(chunks, current, s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        get(chunks, current, s, line);
        strings::emit_length(&mut chunks[current], line);
        get(chunks, current, p, line);
        strings::emit_length(&mut chunks[current], line);
        chunks[current].emit_op(Op::F64_SUB, line);
        host(chunks, current, "ecma:string", "slice", 3, line);
    }
    chunks[current].emit_else(line);
    get(chunks, current, s, line);
    chunks[current].emit_end(line);
}

/// The `substringBefore`/`After`(`Last`) family. Missing delimiter → the
/// default when given, else the receiver itself (Kotlin's contract).
pub fn emit_substring_around(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    after: bool,
    last: bool,
    line: u32,
) {
    let default = if argc >= 3 {
        let d = chunks[current].alloc_scratch(1);
        set(chunks, current, d, line);
        Some(d)
    } else {
        None
    };
    let delim = chunks[current].alloc_scratch(1);
    let s = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    set(chunks, current, delim, line);
    set(chunks, current, s, line);
    get(chunks, current, s, line);
    get(chunks, current, delim, line);
    host(
        chunks,
        current,
        "ecma:string",
        if last { "lastIndexOf" } else { "indexOf" },
        2,
        line,
    );
    set(chunks, current, idx, line);
    get(chunks, current, idx, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_lt(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    match default {
        Some(d) => get(chunks, current, d, line),
        None => get(chunks, current, s, line),
    }
    chunks[current].emit_else(line);
    if after {
        get(chunks, current, s, line);
        get(chunks, current, idx, line);
        get(chunks, current, delim, line);
        strings::emit_length(&mut chunks[current], line);
        ops::emit_dyn_add(&mut chunks[current], line);
        host(chunks, current, "ecma:string", "slice", 2, line);
    } else {
        get(chunks, current, s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        get(chunks, current, idx, line);
        host(chunks, current, "ecma:string", "slice", 3, line);
    }
    chunks[current].emit_end(line);
}

/// `lines()` — split on `\n`, tolerating `\r\n` (Kotlin splits on both).
pub fn emit_lines(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    chunks[current].emit_string_const("\r\n", line);
    chunks[current].emit_string_const("\n", line);
    host(chunks, current, "ecma:string", "replaceAll", 3, line);
    chunks[current].emit_string_const("\n", line);
    host(chunks, current, "ecma:string", "split", 2, line);
}

/// `replaceRange(a, b, replacement)` — end-exclusive.
pub fn emit_replace_range(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let rep = chunks[current].alloc_scratch(1);
    let b = chunks[current].alloc_scratch(1);
    let a = chunks[current].alloc_scratch(1);
    let s = chunks[current].alloc_scratch(1);
    set(chunks, current, rep, line);
    set(chunks, current, b, line);
    set(chunks, current, a, line);
    set(chunks, current, s, line);
    get(chunks, current, s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    get(chunks, current, a, line);
    host(chunks, current, "ecma:string", "slice", 3, line);
    get(chunks, current, rep, line);
    strings::emit_str_concat(&mut chunks[current], line);
    get(chunks, current, s, line);
    get(chunks, current, b, line);
    host(chunks, current, "ecma:string", "slice", 2, line);
    strings::emit_str_concat(&mut chunks[current], line);
}

/// `compareTo(other)` → negative / 0 / positive.
pub fn emit_compare_to(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let b = chunks[current].alloc_scratch(1);
    let a = chunks[current].alloc_scratch(1);
    set(chunks, current, b, line);
    set(chunks, current, a, line);
    get(chunks, current, a, line);
    get(chunks, current, b, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    chunks[current].emit_else(line);
    get(chunks, current, a, line);
    get(chunks, current, b, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `equals(other, ignoreCase = …)` / `contains(needle, ignoreCase = …)` /
/// `startsWith`/`endsWith` with the ignore-case flag — one lowering: fold
/// case on both sides when the flag is truthy, then the plain op.
pub fn emit_ignore_case_op(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    op: &'static str,
    line: u32,
) {
    let flag = if argc >= 3 {
        let f = chunks[current].alloc_scratch(1);
        set(chunks, current, f, line);
        Some(f)
    } else {
        None
    };
    let b = chunks[current].alloc_scratch(1);
    let a = chunks[current].alloc_scratch(1);
    set(chunks, current, b, line);
    set(chunks, current, a, line);
    if let Some(f) = flag {
        get(chunks, current, f, line);
        truthy(chunks, current, line);
        chunks[current].emit_if(line);
        get(chunks, current, a, line);
        strings::emit_to_lower(&mut chunks[current], line);
        set(chunks, current, a, line);
        get(chunks, current, b, line);
        strings::emit_to_lower(&mut chunks[current], line);
        set(chunks, current, b, line);
        chunks[current].emit_end(line);
    }
    get(chunks, current, a, line);
    get(chunks, current, b, line);
    if op == "equals" {
        ops::emit_dyn_eq(&mut chunks[current], line);
        bool_out(chunks, current, line);
    } else {
        host(chunks, current, "ecma:string", op, 2, line);
    }
}

/// `indexOfAny(chars)` / `lastIndexOfAny(chars)`.
pub fn emit_index_of_any(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, last: bool, line: u32) {
    let chars = chunks[current].alloc_scratch(1);
    let s = chunks[current].alloc_scratch(1);
    let best = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let found = chunks[current].alloc_scratch(1);
    set(chunks, current, chars, line);
    set(chunks, current, s, line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    set(chunks, current, best, line);
    let state =
        vybe_compiler::primitives::loops::emit_for_in_start(chunks, current, chars, idx, line);
    set(chunks, current, elem, line);
    get(chunks, current, s, line);
    get(chunks, current, elem, line);
    host(
        chunks,
        current,
        "ecma:string",
        if last { "lastIndexOf" } else { "indexOf" },
        2,
        line,
    );
    set(chunks, current, found, line);
    get(chunks, current, found, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_ge(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if(line);
    // keep the smallest hit (or largest for the `last` form)
    get(chunks, current, best, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_lt(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    get(chunks, current, found, line);
    get(chunks, current, best, line);
    if last {
        ops::emit_dyn_gt(&mut chunks[current], line);
    } else {
        ops::emit_dyn_lt(&mut chunks[current], line);
    }
    chunks[current].emit_end(line);
    truthy(chunks, current, line);
    chunks[current].emit_if(line);
    get(chunks, current, found, line);
    set(chunks, current, best, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    vybe_compiler::primitives::loops::emit_for_in_end(chunks, current, idx, state, line);
    get(chunks, current, best, line);
}

/// `toBoolean()` — `equalsIgnoreCase("true")`; `toBooleanStrictOrNull()` —
/// exactly `"true"`/`"false"`, else null.
pub fn emit_to_boolean(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    strings::emit_to_lower(&mut chunks[current], line);
    chunks[current].emit_string_const("true", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    bool_out(chunks, current, line);
}

pub fn emit_to_boolean_strict_or_null(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let s = chunks[current].alloc_scratch(1);
    set(chunks, current, s, line);
    get(chunks, current, s, line);
    chunks[current].emit_string_const("true", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    bool_out(chunks, current, line);
    chunks[current].emit_else(line);
    get(chunks, current, s, line);
    chunks[current].emit_string_const("false", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    bool_out(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// Char classification on one-char strings. `isDigit` is a range test;
/// `isLetter` is "case-folding changes it OR it is a letter-ish non-cased
/// char" — approximated as the two case folds differing, which covers the
/// bicameral scripts the corpus uses.
pub fn emit_is_digit(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let c = chunks[current].alloc_scratch(1);
    set(chunks, current, c, line);
    get(chunks, current, c, line);
    chunks[current].emit_string_const("0", line);
    ops::emit_dyn_ge(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, c, line);
    chunks[current].emit_string_const("9", line);
    ops::emit_dyn_le(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
    bool_out(chunks, current, line);
}

pub fn emit_is_letter(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let c = chunks[current].alloc_scratch(1);
    set(chunks, current, c, line);
    get(chunks, current, c, line);
    strings::emit_to_lower(&mut chunks[current], line);
    get(chunks, current, c, line);
    strings::emit_to_upper(&mut chunks[current], line);
    ops::emit_dyn_ne(&mut chunks[current], line);
    bool_out(chunks, current, line);
}

/// `trimIndent()` — drop a first/last blank line, remove the common leading
/// whitespace of the non-blank lines.
pub fn emit_trim_indent(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_trim_lines(chunks, current, None, line);
}

/// `trimMargin(prefix = "|")` — drop leading whitespace up to and including
/// the margin char on every line.
pub fn emit_trim_margin(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let margin = if argc >= 2 {
        let m = chunks[current].alloc_scratch(1);
        set(chunks, current, m, line);
        Some(m)
    } else {
        None
    };
    emit_trim_lines(chunks, current, Some(margin), line);
}

/// The shared line loop for `trimIndent` (margin `None`) and `trimMargin`.
fn emit_trim_lines(
    chunks: &mut Vec<Chunk>,
    current: usize,
    margin: Option<Option<u16>>,
    line: u32,
) {
    let s = chunks[current].alloc_scratch(1);
    set(chunks, current, s, line);
    let lines_arr = chunks[current].alloc_scratch(1);
    get(chunks, current, s, line);
    emit_lines(chunks, current, 1, line);
    set(chunks, current, lines_arr, line);

    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);

    match margin {
        None => {
            // trimIndent: first pass — the minimum indent of non-blank lines.
            let min_indent = chunks[current].alloc_scratch(1);
            let indent = chunks[current].alloc_scratch(1);
            core_wasm::i32_const(&mut chunks[current], line, -1);
            set(chunks, current, min_indent, line);
            let state = vybe_compiler::primitives::loops::emit_for_in_start(
                chunks, current, lines_arr, idx, line,
            );
            set(chunks, current, elem, line);
            get(chunks, current, elem, line);
            emit_is_blank(chunks, current, 1, line);
            truthy(chunks, current, line);
            chunks[current].emit_op(Op::I32_EQZ, line);
            chunks[current].emit_if(line);
            // indent = len - len(trimStart)
            get(chunks, current, elem, line);
            strings::emit_length(&mut chunks[current], line);
            get(chunks, current, elem, line);
            host(chunks, current, "ecma:string", "trimStart", 1, line);
            strings::emit_length(&mut chunks[current], line);
            chunks[current].emit_op(Op::F64_SUB, line);
            set(chunks, current, indent, line);
            get(chunks, current, min_indent, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            ops::emit_dyn_lt(&mut chunks[current], line);
            truthy(chunks, current, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_bool_const(true, line);
            chunks[current].emit_else(line);
            get(chunks, current, indent, line);
            get(chunks, current, min_indent, line);
            ops::emit_dyn_lt(&mut chunks[current], line);
            chunks[current].emit_end(line);
            truthy(chunks, current, line);
            chunks[current].emit_if(line);
            get(chunks, current, indent, line);
            set(chunks, current, min_indent, line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
            vybe_compiler::primitives::loops::emit_for_in_end(chunks, current, idx, state, line);

            // second pass: strip that many chars off every non-blank line.
            collections::emit_array_new(chunks, current, 0, line);
            set(chunks, current, out, line);
            let state = vybe_compiler::primitives::loops::emit_for_in_start(
                chunks, current, lines_arr, idx, line,
            );
            set(chunks, current, elem, line);
            get(chunks, current, out, line);
            get(chunks, current, elem, line);
            emit_is_blank(chunks, current, 1, line);
            truthy(chunks, current, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_string_const("", line);
            chunks[current].emit_else(line);
            get(chunks, current, elem, line);
            get(chunks, current, min_indent, line);
            host(chunks, current, "ecma:string", "slice", 2, line);
            chunks[current].emit_end(line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            vybe_compiler::primitives::loops::emit_for_in_end(chunks, current, idx, state, line);
        }
        Some(margin) => {
            // trimMargin: strip up to and past the margin char per line.
            let m = chunks[current].alloc_scratch(1);
            match margin {
                Some(slot) => {
                    get(chunks, current, slot, line);
                    set(chunks, current, m, line);
                }
                None => {
                    chunks[current].emit_string_const("|", line);
                    set(chunks, current, m, line);
                }
            }
            collections::emit_array_new(chunks, current, 0, line);
            set(chunks, current, out, line);
            let pos = chunks[current].alloc_scratch(1);
            let trimmed = chunks[current].alloc_scratch(1);
            let state = vybe_compiler::primitives::loops::emit_for_in_start(
                chunks, current, lines_arr, idx, line,
            );
            set(chunks, current, elem, line);
            get(chunks, current, elem, line);
            host(chunks, current, "ecma:string", "trimStart", 1, line);
            set(chunks, current, trimmed, line);
            get(chunks, current, trimmed, line);
            get(chunks, current, m, line);
            host(chunks, current, "ecma:string", "startsWith", 2, line);
            truthy(chunks, current, line);
            chunks[current].emit_if(line);
            get(chunks, current, out, line);
            get(chunks, current, trimmed, line);
            get(chunks, current, m, line);
            strings::emit_length(&mut chunks[current], line);
            host(chunks, current, "ecma:string", "slice", 2, line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_end(line);
            let _ = pos;
            vybe_compiler::primitives::loops::emit_for_in_end(chunks, current, idx, state, line);
        }
    }

    // Drop a leading/trailing blank line (the raw-string shape), then join.
    emit_strip_edge_blanks(chunks, current, out, line);
    get(chunks, current, out, line);
    chunks[current].emit_string_const("\n", line);
    collections::emit_join(chunks, current, line);
}

/// Remove a first and last EMPTY line from the array in `arr` (in place via a
/// rebuilt array left in the same slot).
fn emit_strip_edge_blanks(chunks: &mut Vec<Chunk>, current: usize, arr: u16, line: u32) {
    // first
    get(chunks, current, arr, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_string_const("", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if(line);
    get(chunks, current, arr, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    get(chunks, current, arr, line);
    collections::emit_len(chunks, current, line);
    collections::emit_slice(chunks, current, line);
    set(chunks, current, arr, line);
    chunks[current].emit_end(line);
    // last
    get(chunks, current, arr, line);
    collections::emit_len(chunks, current, line);
    truthy(chunks, current, line);
    chunks[current].emit_if(line);
    get(chunks, current, arr, line);
    get(chunks, current, arr, line);
    collections::emit_len(chunks, current, line);
    let f = chunks[current].add_import("wasm:js-number", "toF64");
    chunks[current].emit_call(f, 1, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_string_const("", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if(line);
    get(chunks, current, arr, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    get(chunks, current, arr, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_call(f, 1, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    collections::emit_slice(chunks, current, line);
    set(chunks, current, arr, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `reversed()` — a string reverses its characters; a list its elements.
pub fn emit_reversed_any(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    set(chunks, current, v, line);
    get(chunks, current, v, line);
    host(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, v, line);
    chunks[current].emit_string_const("", line);
    host(chunks, current, "ecma:string", "split", 2, line);
    collections::emit_reverse(chunks, current, line);
    chunks[current].emit_string_const("", line);
    collections::emit_join(chunks, current, line);
    chunks[current].emit_else(line);
    get(chunks, current, v, line);
    collections::emit_reverse(chunks, current, line);
    chunks[current].emit_end(line);
}

/// `regionMatches(thisOffset, other, otherOffset, length, ignoreCase = false)`.
pub fn emit_region_matches(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let flag = if argc >= 6 {
        let f = chunks[current].alloc_scratch(1);
        set(chunks, current, f, line);
        Some(f)
    } else {
        None
    };
    let len = chunks[current].alloc_scratch(1);
    let ooff = chunks[current].alloc_scratch(1);
    let other = chunks[current].alloc_scratch(1);
    let toff = chunks[current].alloc_scratch(1);
    let s = chunks[current].alloc_scratch(1);
    set(chunks, current, len, line);
    set(chunks, current, ooff, line);
    set(chunks, current, other, line);
    set(chunks, current, toff, line);
    set(chunks, current, s, line);
    let a = chunks[current].alloc_scratch(1);
    let b = chunks[current].alloc_scratch(1);
    get(chunks, current, s, line);
    get(chunks, current, toff, line);
    get(chunks, current, toff, line);
    get(chunks, current, len, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    host(chunks, current, "ecma:string", "slice", 3, line);
    set(chunks, current, a, line);
    get(chunks, current, other, line);
    get(chunks, current, ooff, line);
    get(chunks, current, ooff, line);
    get(chunks, current, len, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    host(chunks, current, "ecma:string", "slice", 3, line);
    set(chunks, current, b, line);
    if let Some(f) = flag {
        get(chunks, current, f, line);
        truthy(chunks, current, line);
        chunks[current].emit_if(line);
        get(chunks, current, a, line);
        strings::emit_to_lower(&mut chunks[current], line);
        set(chunks, current, a, line);
        get(chunks, current, b, line);
        strings::emit_to_lower(&mut chunks[current], line);
        set(chunks, current, b, line);
        chunks[current].emit_end(line);
    }
    get(chunks, current, a, line);
    get(chunks, current, b, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    bool_out(chunks, current, line);
}
