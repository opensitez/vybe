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
        // Kotlin overloads the second argument: a NUMBER is `startIndex`
        // (`startsWith(prefix, 3)`), anything truthy else is `ignoreCase`.
        if op == "startsWith" || op == "endsWith" {
            get(chunks, current, f, line);
            host(chunks, current, "ecma:value", "typeof", 1, line);
            chunks[current].emit_string_const("number", line);
            ops::emit_dyn_eq(&mut chunks[current], line);
            truthy(chunks, current, line);
            chunks[current].emit_if_value(line);
            get(chunks, current, a, line);
            get(chunks, current, b, line);
            get(chunks, current, f, line);
            host(chunks, current, "ecma:string", op, 3, line);
            chunks[current].emit_else(line);
            emit_case_folded_op(chunks, current, a, b, f, op, line);
            chunks[current].emit_end(line);
            return;
        }
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

/// The ignoreCase leg shared by the startsWith/endsWith overload split.
fn emit_case_folded_op(
    chunks: &mut Vec<Chunk>,
    current: usize,
    a: u16,
    b: u16,
    f: u16,
    op: &str,
    line: u32,
) {
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
    get(chunks, current, a, line);
    get(chunks, current, b, line);
    host(chunks, current, "ecma:string", op, 2, line);
}

/// `indexOfAny(chars)` / `lastIndexOfAny(chars)`.
pub fn emit_index_of_any(chunks: &mut Vec<Chunk>, current: usize, argc: u8, last: bool, line: u32) {
    if argc >= 3 {
        // (chars, startIndex): search the suffix, re-offset hits.
        let start = chunks[current].alloc_scratch(1);
        let chars0 = chunks[current].alloc_scratch(1);
        let s0 = chunks[current].alloc_scratch(1);
        set(chunks, current, start, line);
        set(chunks, current, chars0, line);
        set(chunks, current, s0, line);
        get(chunks, current, s0, line);
        get(chunks, current, start, line);
        host(chunks, current, "ecma:string", "slice", 2, line);
        get(chunks, current, chars0, line);
        emit_index_of_any(chunks, current, 2, last, line);
        let hit = chunks[current].alloc_scratch(1);
        set(chunks, current, hit, line);
        get(chunks, current, hit, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        ops::emit_dyn_lt(&mut chunks[current], line);
        truthy(chunks, current, line);
        chunks[current].emit_if_value(line);
        core_wasm::i32_const(&mut chunks[current], line, -1);
        chunks[current].emit_else(line);
        get(chunks, current, hit, line);
        get(chunks, current, start, line);
        ops::emit_dyn_add(&mut chunks[current], line);
        chunks[current].emit_end(line);
        return;
    }
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
    // Kotlin's `reversed()` returns an INDEPENDENT list.
    get(chunks, current, v, line);
    collections::emit_clone(chunks, current, line);
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

/// `indexOf(needle, from)` / `lastIndexOf(needle, from)` — the offset was
/// silently dropped by the 2-arg host path.
pub fn emit_index_of_from(chunks: &mut Vec<Chunk>, current: usize, argc: u8, last: bool, line: u32) {
    let from = if argc >= 3 {
        let f = chunks[current].alloc_scratch(1);
        set(chunks, current, f, line);
        Some(f)
    } else {
        None
    };
    let needle = chunks[current].alloc_scratch(1);
    let s = chunks[current].alloc_scratch(1);
    set(chunks, current, needle, line);
    set(chunks, current, s, line);
    match from {
        Some(f) if !last => {
            // Search the SUFFIX and re-offset the hit.
            let idx = chunks[current].alloc_scratch(1);
            get(chunks, current, s, line);
            get(chunks, current, f, line);
            host(chunks, current, "ecma:string", "slice", 2, line);
            get(chunks, current, needle, line);
            host(chunks, current, "ecma:string", "indexOf", 2, line);
            set(chunks, current, idx, line);
            get(chunks, current, idx, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            ops::emit_dyn_lt(&mut chunks[current], line);
            truthy(chunks, current, line);
            chunks[current].emit_if_value(line);
            core_wasm::i32_const(&mut chunks[current], line, -1);
            chunks[current].emit_else(line);
            get(chunks, current, idx, line);
            get(chunks, current, f, line);
            ops::emit_dyn_add(&mut chunks[current], line);
            chunks[current].emit_end(line);
        }
        _ => {
            get(chunks, current, s, line);
            get(chunks, current, needle, line);
            host(
                chunks,
                current,
                "ecma:string",
                if last { "lastIndexOf" } else { "indexOf" },
                2,
                line,
            );
        }
    }
}

/// `chunked(n)` / `windowed(size, step = 1)` at RUNTIME, for strings and
/// lists. A string chunks to STRINGS; a list to lists. `windowed` keeps only
/// FULL windows (Kotlin's default `partialWindows = false`).
pub fn emit_chunked_windowed(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    windowed: bool,
    line: u32,
) {
    emit_chunked_windowed_ex(chunks, current, argc, windowed, /*partial:*/ !windowed, line);
}

/// `partial`: keep the tail window shorter than `size` (chunked always does;
/// windowed only with `partialWindows = true`).
pub fn emit_chunked_windowed_ex(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    windowed: bool,
    partial: bool,
    line: u32,
) {
    // [recv, size, (step)] — chunked's step == size.
    let step = chunks[current].alloc_scratch(1);
    let size = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    if windowed && argc >= 3 {
        set(chunks, current, step, line);
        set(chunks, current, size, line);
    } else {
        set(chunks, current, size, line);
        if windowed {
            core_wasm::i32_const(&mut chunks[current], line, 1);
        } else {
            get(chunks, current, size, line);
        }
        set(chunks, current, step, line);
    }
    set(chunks, current, recv, line);

    let is_str = chunks[current].alloc_scratch(1);
    get(chunks, current, recv, line);
    host(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    set(chunks, current, is_str, line);

    let view = chunks[current].alloc_scratch(1);
    get(chunks, current, is_str, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, recv, line);
    chunks[current].emit_string_const("", line);
    host(chunks, current, "ecma:string", "split", 2, line);
    chunks[current].emit_else(line);
    get(chunks, current, recv, line);
    chunks[current].emit_end(line);
    set(chunks, current, view, line);

    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let piece = chunks[current].alloc_scratch(1);
    let upper = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(chunks, current, i, line);
    get(chunks, current, view, line);
    collections::emit_len(chunks, current, line);
    set(chunks, current, len, line);

    let _block = chunks[current].emit_block(line);
    let (_loop, _) = chunks[current].emit_loop_s(line);
    get(chunks, current, i, line);
    get(chunks, current, len, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    // upper = min(i + size, len); windowed drops partials.
    get(chunks, current, i, line);
    get(chunks, current, size, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(chunks, current, upper, line);
    get(chunks, current, upper, line);
    get(chunks, current, len, line);
    ops::emit_dyn_gt(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if(line);
    if !partial {
        // no partial window: jump out
        get(chunks, current, len, line);
        set(chunks, current, i, line);
        chunks[current].emit_end(line);
        get(chunks, current, i, line);
        get(chunks, current, len, line);
        ops::emit_dyn_lt(&mut chunks[current], line);
        truthy(chunks, current, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_br_if(1, line);
    } else {
        get(chunks, current, len, line);
        set(chunks, current, upper, line);
        chunks[current].emit_end(line);
    }

    get(chunks, current, view, line);
    get(chunks, current, i, line);
    get(chunks, current, upper, line);
    collections::emit_slice(chunks, current, line);
    set(chunks, current, piece, line);
    get(chunks, current, out, line);
    get(chunks, current, is_str, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, piece, line);
    chunks[current].emit_string_const("", line);
    collections::emit_join(chunks, current, line);
    chunks[current].emit_else(line);
    get(chunks, current, piece, line);
    chunks[current].emit_end(line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(chunks, current, i, line);
    get(chunks, current, step, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(chunks, current, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    get(chunks, current, out, line);
}

/// `slice(rangeValue)` where the range arrives MATERIALIZED as an array
/// (`IntRange(a, b)`, a range held in a variable): from = first element,
/// to = last element + 1.
pub fn emit_slice_range_value(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let r = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    set(chunks, current, r, line);
    set(chunks, current, recv, line);
    let from = chunks[current].alloc_scratch(1);
    let to = chunks[current].alloc_scratch(1);
    get(chunks, current, r, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    set(chunks, current, from, line);
    get(chunks, current, r, line);
    get(chunks, current, r, line);
    collections::emit_len(chunks, current, line);
    let f = chunks[current].add_import("wasm:js-number", "toF64");
    chunks[current].emit_call(f, 1, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    collections::emit_get(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(chunks, current, to, line);
    get(chunks, current, recv, line);
    get(chunks, current, from, line);
    get(chunks, current, to, line);
    crate::emitter::hof::emit_slice_any(chunks, current, 3, line);
}

/// `indexOf(x)` / `lastIndexOf(x)` for ANY receiver — string→host, list→shared.
pub fn emit_index_of_any_recv(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    last: bool,
    line: u32,
) {
    let needle = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    set(chunks, current, needle, line);
    set(chunks, current, recv, line);
    get(chunks, current, recv, line);
    host(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, recv, line);
    get(chunks, current, needle, line);
    host(
        chunks,
        current,
        "ecma:string",
        if last { "lastIndexOf" } else { "indexOf" },
        2,
        line,
    );
    chunks[current].emit_else(line);
    get(chunks, current, recv, line);
    get(chunks, current, needle, line);
    if last {
        collections::emit_last_index_of(chunks, current, line);
    } else {
        collections::emit_index_of(chunks, current, line);
    }
    chunks[current].emit_end(line);
}

/// `toByteOrNull()` / `toShortOrNull()` — the int parse plus a bounds check.
pub fn emit_to_bounded_or_null(
    chunks: &mut Vec<Chunk>,
    current: usize,
    lo: i32,
    hi: i32,
    line: u32,
) {
    crate::emitter::numbers::emit_to_int_or_null(chunks, current, 1, line);
    let v = chunks[current].alloc_scratch(1);
    set(chunks, current, v, line);
    // NOT `REF_IS_NULL`: a NUMBER is not a ref and reads as null there.
    get(chunks, current, v, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    get(chunks, current, v, line);
    core_wasm::i32_const(&mut chunks[current], line, lo);
    ops::emit_dyn_ge(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, v, line);
    core_wasm::i32_const(&mut chunks[current], line, hi);
    ops::emit_dyn_le(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, v, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `trim('-')` / `trimStart(chars…)` / `trimEnd(chars…)` — strip the GIVEN
/// characters. Bare `trim()` never reaches here (host trim serves argc 1).
pub fn emit_trim_chars(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    start: bool,
    end: bool,
    line: u32,
) {
    let n = argc.saturating_sub(1);
    let base = chunks[current].alloc_scratch(n.max(1) as u16);
    collections::emit_pack_n(chunks, current, n as u16, base, line);
    let set_arr = chunks[current].alloc_scratch(1);
    set(chunks, current, set_arr, line);
    let s = chunks[current].alloc_scratch(1);
    set(chunks, current, s, line);

    if start {
        emit_strip_edge(chunks, current, s, set_arr, /*front:*/ true, line);
    }
    if end {
        emit_strip_edge(chunks, current, s, set_arr, /*front:*/ false, line);
    }
    get(chunks, current, s, line);
}

fn emit_strip_edge(
    chunks: &mut Vec<Chunk>,
    current: usize,
    s: u16,
    set_arr: u16,
    front: bool,
    line: u32,
) {
    let ch = chunks[current].alloc_scratch(1);
    let _block = chunks[current].emit_block(line);
    let (_loop, _) = chunks[current].emit_loop_s(line);
    get(chunks, current, s, line);
    strings::emit_length(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(chunks, current, s, line);
    if front {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        host(chunks, current, "ecma:string", "slice", 3, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, -1);
        host(chunks, current, "ecma:string", "slice", 2, line);
    }
    set(chunks, current, ch, line);
    get(chunks, current, set_arr, line);
    get(chunks, current, ch, line);
    collections::emit_contains(chunks, current, line);
    truthy(chunks, current, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(chunks, current, s, line);
    if front {
        core_wasm::i32_const(&mut chunks[current], line, 1);
        host(chunks, current, "ecma:string", "slice", 2, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        core_wasm::i32_const(&mut chunks[current], line, -1);
        host(chunks, current, "ecma:string", "slice", 3, line);
    }
    set(chunks, current, s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `split(sep, limit = n)` — Kotlin's limit keeps the REMAINDER as the last
/// element; the ECMA host limit truncates it away.
pub fn emit_split_limit(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc <= 2 {
        host(chunks, current, "ecma:string", "split", 2, line);
        return;
    }
    let limit = chunks[current].alloc_scratch(1);
    let sep = chunks[current].alloc_scratch(1);
    let s = chunks[current].alloc_scratch(1);
    set(chunks, current, limit, line);
    set(chunks, current, sep, line);
    set(chunks, current, s, line);
    get(chunks, current, limit, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_le(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, s, line);
    get(chunks, current, sep, line);
    host(chunks, current, "ecma:string", "split", 2, line);
    chunks[current].emit_else(line);
    // parts = split; out = parts[0..limit-1] + join(parts[limit-1..], sep)
    let parts = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let cut = chunks[current].alloc_scratch(1);
    get(chunks, current, s, line);
    get(chunks, current, sep, line);
    host(chunks, current, "ecma:string", "split", 2, line);
    set(chunks, current, parts, line);
    get(chunks, current, limit, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    dyn_sub_str(chunks, current, line);
    set(chunks, current, cut, line);
    get(chunks, current, parts, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    get(chunks, current, cut, line);
    collections::emit_slice(chunks, current, line);
    set(chunks, current, out, line);
    get(chunks, current, out, line);
    get(chunks, current, parts, line);
    get(chunks, current, cut, line);
    get(chunks, current, parts, line);
    collections::emit_len(chunks, current, line);
    collections::emit_slice(chunks, current, line);
    get(chunks, current, sep, line);
    collections::emit_join(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(chunks, current, out, line);
    chunks[current].emit_end(line);
}

fn dyn_sub_str(chunks: &mut [Chunk], current: usize, line: u32) {
    let b = chunks[current].alloc_scratch(1);
    set(chunks, current, b, line);
    let f = chunks[current].add_import("wasm:js-number", "toF64");
    chunks[current].emit_call(f, 1, line);
    get(chunks, current, b, line);
    chunks[current].emit_call(f, 1, line);
    chunks[current].emit_op(Op::F64_SUB, line);
}

/// `s[i]` / `substring(a, b)` with Kotlin's throw-on-out-of-bounds contract.
pub fn emit_char_at_throwing(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let i = chunks[current].alloc_scratch(1);
    let s = chunks[current].alloc_scratch(1);
    set(chunks, current, i, line);
    set(chunks, current, s, line);
    emit_bounds_or_throw(chunks, current, s, i, i, /*index:*/ true, line);
    get(chunks, current, s, line);
    get(chunks, current, i, line);
    get(chunks, current, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    ops::emit_dyn_add(&mut chunks[current], line);
    host(chunks, current, "ecma:string", "slice", 3, line);
}

pub fn emit_substring_throwing(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc >= 3 {
        let b = chunks[current].alloc_scratch(1);
        let a = chunks[current].alloc_scratch(1);
        let s = chunks[current].alloc_scratch(1);
        set(chunks, current, b, line);
        set(chunks, current, a, line);
        set(chunks, current, s, line);
        emit_bounds_or_throw(chunks, current, s, a, b, /*index:*/ false, line);
        get(chunks, current, s, line);
        get(chunks, current, a, line);
        get(chunks, current, b, line);
        host(chunks, current, "ecma:string", "slice", 3, line);
    } else {
        let a = chunks[current].alloc_scratch(1);
        let s = chunks[current].alloc_scratch(1);
        set(chunks, current, a, line);
        set(chunks, current, s, line);
        emit_bounds_or_throw(chunks, current, s, a, a, /*index:*/ false, line);
        get(chunks, current, s, line);
        get(chunks, current, a, line);
        host(chunks, current, "ecma:string", "slice", 2, line);
    }
}

/// Throw `IndexOutOfBoundsException` unless `0 <= a`, `b <= len` (and for an
/// INDEX read, `a < len`).
fn emit_bounds_or_throw(
    chunks: &mut Vec<Chunk>,
    current: usize,
    s: u16,
    a: u16,
    b: u16,
    index: bool,
    line: u32,
) {
    let exc = if index {
        "StringIndexOutOfBoundsException"
    } else {
        "IndexOutOfBoundsException"
    };
    let bad = chunks[current].alloc_scratch(1);
    get(chunks, current, a, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_lt(&mut chunks[current], line);
    truthy(chunks, current, line);
    set(chunks, current, bad, line);
    get(chunks, current, b, line);
    get(chunks, current, s, line);
    strings::emit_length(&mut chunks[current], line);
    if index {
        ops::emit_dyn_ge(&mut chunks[current], line);
    } else {
        ops::emit_dyn_gt(&mut chunks[current], line);
    }
    truthy(chunks, current, line);
    get(chunks, current, bad, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("index out of bounds", line);
    crate::emitter::nullability::emit_exception(chunks, current, 1, exc, line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

/// `contains(x)` for ANY receiver — string→substring test, else the shared
/// collection contains.
pub fn emit_contains_any(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let needle = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    set(chunks, current, needle, line);
    set(chunks, current, recv, line);
    get(chunks, current, recv, line);
    host(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, recv, line);
    get(chunks, current, needle, line);
    host(chunks, current, "ecma:string", "includes", 2, line);
    chunks[current].emit_else(line);
    get(chunks, current, recv, line);
    get(chunks, current, needle, line);
    collections::emit_contains(chunks, current, line);
    chunks[current].emit_end(line);
}

/// `toInt(radix)` / `toIntOrNull(radix)` — `ecma:global.parseInt` with the
/// radix; NaN maps to null for the OrNull form and throws for `toInt`.
pub fn emit_parse_radix(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, or_null: bool, line: u32) {
    let radix = chunks[current].alloc_scratch(1);
    let s = chunks[current].alloc_scratch(1);
    let v = chunks[current].alloc_scratch(1);
    set(chunks, current, radix, line);
    set(chunks, current, s, line);
    // Kotlin rejects radix outside 2..36 with an IllegalArgumentException,
    // and the OrNull form still throws for a bad RADIX (it is the VALUE that
    // may be null).
    get(chunks, current, radix, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    ops::emit_dyn_lt(&mut chunks[current], line);
    truthy(chunks, current, line);
    let bad = chunks[current].alloc_scratch(1);
    set(chunks, current, bad, line);
    get(chunks, current, radix, line);
    core_wasm::i32_const(&mut chunks[current], line, 36);
    ops::emit_dyn_gt(&mut chunks[current], line);
    truthy(chunks, current, line);
    get(chunks, current, bad, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("radix out of range", line);
    crate::emitter::nullability::emit_exception(
        chunks,
        current,
        1,
        "IllegalArgumentException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    get(chunks, current, s, line);
    get(chunks, current, radix, line);
    host(chunks, current, "ecma:global", "parseInt", 2, line);
    set(chunks, current, v, line);
    get(chunks, current, v, line);
    host(chunks, current, "ecma:number", "isNaN", 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    if or_null {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    } else {
        chunks[current].emit_string_const("invalid number", line);
        crate::emitter::nullability::emit_exception(
            chunks,
            current,
            1,
            "NumberFormatException",
            line,
        );
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    chunks[current].emit_else(line);
    get(chunks, current, v, line);
    chunks[current].emit_end(line);
}

/// Kotlin-strict `toIntOrNull()`: leading `+` allowed, anything that does not
/// round-trip (`"12L"`, `"2x"`) is null — the lenient shared parse accepted
/// junk suffixes.
pub fn emit_strict_int_or_null(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let s = chunks[current].alloc_scratch(1);
    set(chunks, current, s, line);
    // strip a single leading '+'
    get(chunks, current, s, line);
    chunks[current].emit_string_const("+", line);
    host(chunks, current, "ecma:string", "startsWith", 2, line);
    truthy(chunks, current, line);
    chunks[current].emit_if(line);
    get(chunks, current, s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    host(chunks, current, "ecma:string", "slice", 2, line);
    set(chunks, current, s, line);
    chunks[current].emit_end(line);

    let v = chunks[current].alloc_scratch(1);
    get(chunks, current, s, line);
    crate::emitter::numbers::emit_to_int_or_null(chunks, current, 1, line);
    set(chunks, current, v, line);
    get(chunks, current, v, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    // round-trip: String(v) must equal the (plus-stripped) source
    get(chunks, current, v, line);
    host(chunks, current, "ecma:string", "String", 1, line);
    get(chunks, current, s, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, v, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}
