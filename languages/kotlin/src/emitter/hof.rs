//! Kotlin's lambda-taking collection extensions (`kotlin.collections`).
//!
//! These are KOTLIN's own stdlib surface, not JDK classes — which is why they
//! live in the language crate and not `platforms/jvm`. Each adapter receives
//! the receiver deepest on the stack, then the source arguments, with
//! `argc` counting receiver + args (the shared instance-target convention). A
//! lambda argument arrives as an ordinary closure value; calling it goes
//! through the shared callable primitive, so Kotlin HOFs follow the same
//! machinery as Java streams and delegate/function-reference invocation.
//!
//! Result collections are built with the shared primitives: arrays via
//! `collections::*`, maps via `dict::*` (`ARRAY_SET` keeps `__keys` insertion
//! order), and Pairs via `tuples::emit_tuple`, which stamps the tag the
//! renderer's `(a, b)` bracket decision reads.

use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::{callable, collections, dict, loops, ops, sets, tuples};
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

/// `[a, b] → [a - b]` over dynamic numbers: both through `wasm:js-number.toF64`.
fn dyn_sub(chunks: &mut [Chunk], current: usize, line: u32) {
    let b = chunks[current].alloc_scratch(1);
    set(chunks, current, b, line);
    let f = chunks[current].add_import("wasm:js-number", "toF64");
    chunks[current].emit_call(f, 1, line);
    get(chunks, current, b, line);
    chunks[current].emit_call(f, 1, line);
    chunks[current].emit_op(Op::F64_SUB, line);
}

/// `[a, b] → [a / b]` over dynamic numbers.
fn dyn_div(chunks: &mut [Chunk], current: usize, line: u32) {
    let b = chunks[current].alloc_scratch(1);
    set(chunks, current, b, line);
    let f = chunks[current].add_import("wasm:js-number", "toF64");
    chunks[current].emit_call(f, 1, line);
    get(chunks, current, b, line);
    chunks[current].emit_call(f, 1, line);
    chunks[current].emit_op(Op::F64_DIV, line);
}

/// `fn(args…)` with every operand in a local. Leaves the result on the stack.
fn call_fn(chunks: &mut [Chunk], current: usize, fn_slot: u16, args: &[u16], line: u32) {
    get(chunks, current, fn_slot, line);
    for &a in args {
        get(chunks, current, a, line);
    }
    callable::emit_direct_invoke(chunks, current, args.len() as u8, line);
}

/// `arr[i]` from locals. Leaves the element on the stack.
fn elem_at(chunks: &mut [Chunk], current: usize, arr: u16, i: u16, line: u32) {
    get(chunks, current, arr, line);
    get(chunks, current, i, line);
    collections::emit_get(chunks, current, line);
}

/// `arr.length` from a local. Leaves the length on the stack.
fn len_of(chunks: &mut [Chunk], current: usize, arr: u16, line: u32) {
    get(chunks, current, arr, line);
    collections::emit_len(chunks, current, line);
}

/// Iterate `arr`, leaving each ELEMENT in `elem` for `body`.
fn for_each(
    chunks: &mut Vec<Chunk>,
    current: usize,
    arr: u16,
    idx: u16,
    elem: u16,
    line: u32,
    body: impl FnOnce(&mut Vec<Chunk>),
) {
    let state = loops::emit_for_in_start(chunks, current, arr, idx, line);
    set(chunks, current, elem, line);
    body(chunks);
    loops::emit_for_in_end(chunks, current, idx, state, line);
}

/// Throw `NoSuchElementException(msg)` — what `first()`/`last()`/`single()`/
/// `reduce()` do on an empty receiver (kotlin.collections contract).
fn throw_no_such_element(chunks: &mut Vec<Chunk>, current: usize, msg: &str, line: u32) {
    chunks[current].emit_string_const(msg, line);
    crate::emitter::nullability::emit_exception(chunks, current, 1, "NoSuchElementException", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
}

/// Pop `[receiver, fn]` into locals. Returns `(arr, f)`.
fn pop_recv_fn(chunks: &mut [Chunk], current: usize, line: u32) -> (u16, u16) {
    let f = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, arr, line);
    (arr, f)
}

// ── Predicates over the whole receiver ──────────────────────────────────────

/// `takeWhile { }` / `dropWhile { }` — one loop, direction by flag.
pub fn emit_take_while(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_while_split(chunks, current, /*take:*/ true, line);
}

pub fn emit_drop_while(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_while_split(chunks, current, /*take:*/ false, line);
}

fn emit_while_split(chunks: &mut Vec<Chunk>, current: usize, take: bool, line: u32) {
    let (recv, f) = pop_recv_fn(chunks, current, line);
    let is_str = chunks[current].alloc_scratch(1);
    get(chunks, current, recv, line);
    let type_of = chunks[current].add_import("ecma:value", "typeof");
    chunks[current].emit_call(type_of, 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    set(chunks, current, is_str, line);
    let arr = chunks[current].alloc_scratch(1);
    get(chunks, current, recv, line);
    crate::emitter::collections::emit_dict_as_list(chunks, current, line);
    set(chunks, current, arr, line);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let active = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    chunks[current].emit_bool_const(true, line);
    set(chunks, current, active, line);

    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        get(chunks, current, active, line);
        truthy(chunks, current, line);
        chunks[current].emit_if(line);
        call_fn(chunks, current, f, &[elem], line);
        truthy(chunks, current, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        chunks[current].emit_bool_const(false, line);
        set(chunks, current, active, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);

        // takeWhile appends while ACTIVE; dropWhile appends once INACTIVE.
        get(chunks, current, active, line);
        truthy(chunks, current, line);
        if !take {
            chunks[current].emit_op(Op::I32_EQZ, line);
        }
        chunks[current].emit_if(line);
        get(chunks, current, out, line);
        get(chunks, current, elem, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, is_str, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, out, line);
    chunks[current].emit_string_const("", line);
    collections::emit_join(chunks, current, line);
    chunks[current].emit_else(line);
    get(chunks, current, out, line);
    chunks[current].emit_end(line);
}

/// `count()` / `count { }`.
pub fn emit_count(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc <= 1 {
        // Bare count(): length — of the array, or of a dict's keys.
        emit_size_any(chunks, current, line);
        return;
    }
    let (arr, f) = pop_recv_fn(chunks, current, line);
    let arr = emit_entries_if_dict(chunks, current, arr, line);
    let n = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(chunks, current, n, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[elem], line);
        truthy(chunks, current, line);
        chunks[current].emit_if(line);
        get(chunks, current, n, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        ops::emit_dyn_add(&mut chunks[current], line);
        set(chunks, current, n, line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, n, line);
}

/// `none()` / `none { }`.
pub fn emit_none(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_count(chunks, current, argc, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `sumOf { }`.
pub fn emit_sum_of(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr, f) = pop_recv_fn(chunks, current, line);
    let acc = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(chunks, current, acc, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        get(chunks, current, acc, line);
        call_fn(chunks, current, f, &[elem], line);
        ops::emit_dyn_add(&mut chunks[current], line);
        set(chunks, current, acc, line);
    });
    get(chunks, current, acc, line);
}

/// `minByOrNull { }` / `maxByOrNull { }`.
pub fn emit_min_by_or_null(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_extreme_by(chunks, current, /*min:*/ true, line);
}

pub fn emit_max_by_or_null(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_extreme_by(chunks, current, /*min:*/ false, line);
}

fn emit_extreme_by(chunks: &mut Vec<Chunk>, current: usize, min: bool, line: u32) {
    let (recv, f) = pop_recv_fn(chunks, current, line);
    let arr = emit_list_view(chunks, current, recv, line);
    let best = chunks[current].alloc_scratch(1);
    let best_key = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let started = chunks[current].alloc_scratch(1);

    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    set(chunks, current, best, line);
    chunks[current].emit_bool_const(false, line);
    set(chunks, current, started, line);

    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[elem], line);
        set(chunks, current, key, line);

        get(chunks, current, started, line);
        truthy(chunks, current, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_bool_const(true, line);
        chunks[current].emit_else(line);
        get(chunks, current, key, line);
        get(chunks, current, best_key, line);
        if min {
            ops::emit_dyn_lt(&mut chunks[current], line);
        } else {
            ops::emit_dyn_gt(&mut chunks[current], line);
        }
        chunks[current].emit_end(line);
        truthy(chunks, current, line);
        chunks[current].emit_if(line);
        get(chunks, current, elem, line);
        set(chunks, current, best, line);
        get(chunks, current, key, line);
        set(chunks, current, best_key, line);
        chunks[current].emit_bool_const(true, line);
        set(chunks, current, started, line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, best, line);
}

/// `indexOfFirst { }` / `indexOfLast { }` / `findLast { }`.
pub fn emit_index_of_first(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_scan_matches(chunks, current, ScanResult::FirstIndex, line);
}

pub fn emit_index_of_last(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_scan_matches(chunks, current, ScanResult::LastIndex, line);
}

pub fn emit_find_last(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_scan_matches(chunks, current, ScanResult::LastElem, line);
}

enum ScanResult {
    FirstIndex,
    LastIndex,
    FirstElem,
    LastElem,
}

/// `find { }` / `findLast { }` — element or null, over any receiver.
pub fn emit_find(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_scan_matches(chunks, current, ScanResult::FirstElem, line);
}

pub fn emit_find_last_any(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_scan_matches(chunks, current, ScanResult::LastElem, line);
}

fn emit_scan_matches(chunks: &mut Vec<Chunk>, current: usize, kind: ScanResult, line: u32) {
    let (recv, f) = pop_recv_fn(chunks, current, line);
    let arr = emit_list_view(chunks, current, recv, line);
    let result = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let done = chunks[current].alloc_scratch(1);
    match kind {
        ScanResult::FirstIndex | ScanResult::LastIndex => {
            core_wasm::i32_const(&mut chunks[current], line, -1);
        }
        ScanResult::FirstElem | ScanResult::LastElem => {
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        }
    }
    set(chunks, current, result, line);
    chunks[current].emit_bool_const(false, line);
    set(chunks, current, done, line);

    let first_only = matches!(kind, ScanResult::FirstIndex | ScanResult::FirstElem);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        if first_only {
            get(chunks, current, done, line);
            truthy(chunks, current, line);
            chunks[current].emit_op(Op::I32_EQZ, line);
        } else {
            chunks[current].emit_bool_const(true, line);
        }
        chunks[current].emit_if(line);
        call_fn(chunks, current, f, &[elem], line);
        truthy(chunks, current, line);
        chunks[current].emit_if(line);
        match kind {
            ScanResult::FirstIndex | ScanResult::LastIndex => {
                get(chunks, current, idx, line);
            }
            ScanResult::FirstElem | ScanResult::LastElem => {
                get(chunks, current, elem, line);
            }
        }
        set(chunks, current, result, line);
        chunks[current].emit_bool_const(true, line);
        set(chunks, current, done, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, result, line);
}

// ── first / last / single, with their throw-on-empty contract ───────────────

/// `first()` / `first { }` — throws `NoSuchElementException` when nothing
/// matches, which is Kotlin's contract and the reason `null` was wrong.
pub fn emit_first(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_edge(
        chunks,
        current,
        argc,
        Edge::First,
        /*or_null:*/ false,
        line,
    );
}

pub fn emit_first_or_null(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_edge(
        chunks,
        current,
        argc,
        Edge::First,
        /*or_null:*/ true,
        line,
    );
}

pub fn emit_last(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_edge(
        chunks,
        current,
        argc,
        Edge::Last,
        /*or_null:*/ false,
        line,
    );
}

pub fn emit_last_or_null(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_edge(
        chunks,
        current,
        argc,
        Edge::Last,
        /*or_null:*/ true,
        line,
    );
}

enum Edge {
    First,
    Last,
}

fn emit_edge(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    edge: Edge,
    or_null: bool,
    line: u32,
) {
    // With a predicate: filter semantics, but only the edge element is kept.
    let f = if argc >= 2 {
        let f = chunks[current].alloc_scratch(1);
        set(chunks, current, f, line);
        Some(f)
    } else {
        None
    };
    let arr0 = chunks[current].alloc_scratch(1);
    set(chunks, current, arr0, line);
    let arr = emit_list_view(chunks, current, arr0, line);

    let result = chunks[current].alloc_scratch(1);
    let found = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    set(chunks, current, result, line);
    chunks[current].emit_bool_const(false, line);
    set(chunks, current, found, line);

    let keep_first = matches!(edge, Edge::First);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        if keep_first {
            get(chunks, current, found, line);
            truthy(chunks, current, line);
            chunks[current].emit_op(Op::I32_EQZ, line);
        } else {
            chunks[current].emit_bool_const(true, line);
        }
        chunks[current].emit_if(line);
        if let Some(f) = f {
            call_fn(chunks, current, f, &[elem], line);
            truthy(chunks, current, line);
        } else {
            chunks[current].emit_bool_const(true, line);
            truthy(chunks, current, line);
        }
        chunks[current].emit_if(line);
        get(chunks, current, elem, line);
        set(chunks, current, result, line);
        chunks[current].emit_bool_const(true, line);
        set(chunks, current, found, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    });

    get(chunks, current, found, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, result, line);
    chunks[current].emit_else(line);
    if or_null {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    } else {
        throw_no_such_element(chunks, current, "Collection is empty.", line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    chunks[current].emit_end(line);
}

/// `single()` / `singleOrNull()`.
pub fn emit_single(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_single_impl(chunks, current, /*or_null:*/ false, line);
}

pub fn emit_single_or_null(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_single_impl(chunks, current, /*or_null:*/ true, line);
}

fn emit_single_impl(chunks: &mut Vec<Chunk>, current: usize, or_null: bool, line: u32) {
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, arr, line);
    len_of(chunks, current, arr, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, arr, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    if or_null {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    } else {
        throw_no_such_element(
            chunks,
            current,
            "Collection has not exactly one element.",
            line,
        );
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    chunks[current].emit_end(line);
}

// ── Element access with fallback ────────────────────────────────────────────

/// `getOrNull(i)` / `elementAtOrNull(i)` — `null` out of bounds.
pub fn emit_get_or_null(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let i = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, i, line);
    set(chunks, current, arr, line);
    get(chunks, current, arr, line);
    let is_arr = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_arr, 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    emit_in_bounds(chunks, current, arr, i, line);
    chunks[current].emit_if_value(line);
    elem_at(chunks, current, arr, i, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    // Map receiver: the key's value or null.
    get(chunks, current, arr, line);
    get(chunks, current, i, line);
    dict::emit_get_dynamic(chunks, current, line);
    chunks[current].emit_end(line);
}

/// `getOrElse(i) { }` / `elementAtOrElse(i) { }` — `fn(i)` out of bounds.
pub fn emit_get_or_else(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, i, line);
    set(chunks, current, arr, line);
    get(chunks, current, arr, line);
    let is_arr = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_arr, 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    emit_in_bounds(chunks, current, arr, i, line);
    chunks[current].emit_if_value(line);
    elem_at(chunks, current, arr, i, line);
    chunks[current].emit_else(line);
    call_fn(chunks, current, f, &[i], line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    // Map receiver: existing value, else the supplier — Kotlin's
    // `map.getOrElse(k) { fallback }`.
    get(chunks, current, arr, line);
    get(chunks, current, i, line);
    dict::emit_method_has(chunks, current, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, arr, line);
    get(chunks, current, i, line);
    dict::emit_get_dynamic(chunks, current, line);
    chunks[current].emit_else(line);
    call_fn(chunks, current, f, &[], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// Push i32 `1` when `0 <= i < arr.length`.
fn emit_in_bounds(chunks: &mut Vec<Chunk>, current: usize, arr: u16, i: u16, line: u32) {
    get(chunks, current, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_ge(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, i, line);
    len_of(chunks, current, arr, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
}

// ── Transformations producing a new list ────────────────────────────────────

/// `filterNot { }`.
pub fn emit_filter_not(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr, f) = pop_recv_fn(chunks, current, line);
    let arr = emit_entries_if_dict(chunks, current, arr, line);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[elem], line);
        truthy(chunks, current, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        get(chunks, current, out, line);
        get(chunks, current, elem, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, out, line);
}

/// `filterIndexed { i, v -> }` — a string filters characters and returns a
/// STRING.
pub fn emit_filter_indexed(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (recv, f) = pop_recv_fn(chunks, current, line);
    let is_str = chunks[current].alloc_scratch(1);
    get(chunks, current, recv, line);
    let type_of = chunks[current].add_import("ecma:value", "typeof");
    chunks[current].emit_call(type_of, 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    set(chunks, current, is_str, line);
    let arr = chunks[current].alloc_scratch(1);
    get(chunks, current, recv, line);
    crate::emitter::collections::emit_dict_as_list(chunks, current, line);
    set(chunks, current, arr, line);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[idx, elem], line);
        truthy(chunks, current, line);
        chunks[current].emit_if(line);
        get(chunks, current, out, line);
        get(chunks, current, elem, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, is_str, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, out, line);
    chunks[current].emit_string_const("", line);
    collections::emit_join(chunks, current, line);
    chunks[current].emit_else(line);
    get(chunks, current, out, line);
    chunks[current].emit_end(line);
}

/// `filterNotNull()`.
pub fn emit_filter_not_null(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, arr, line);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        get(chunks, current, elem, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        get(chunks, current, out, line);
        get(chunks, current, elem, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, out, line);
}

/// `mapNotNull { }`.
pub fn emit_map_not_null(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr, f) = pop_recv_fn(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let mapped = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[elem], line);
        set(chunks, current, mapped, line);
        get(chunks, current, mapped, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        get(chunks, current, out, line);
        get(chunks, current, mapped, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, out, line);
}

/// `mapIndexed { i, v -> }`.
pub fn emit_map_indexed(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr, f) = pop_recv_fn(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        get(chunks, current, out, line);
        call_fn(chunks, current, f, &[idx, elem], line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    });
    get(chunks, current, out, line);
}

/// `forEachIndexed { i, v -> }` — Unit result.
pub fn emit_for_each_indexed(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr, f) = pop_recv_fn(chunks, current, line);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[idx, elem], line);
        chunks[current].emit_op(Op::DROP, line);
    });
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `onEach { }` — forEach that returns the receiver.
pub fn emit_on_each(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr, f) = pop_recv_fn(chunks, current, line);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[elem], line);
        chunks[current].emit_op(Op::DROP, line);
    });
    get(chunks, current, arr, line);
}

/// `distinctBy { }` — first element per distinct key.
pub fn emit_distinct_by(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr, f) = pop_recv_fn(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    let seen = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, seen, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[elem], line);
        set(chunks, current, key, line);
        get(chunks, current, seen, line);
        get(chunks, current, key, line);
        collections::emit_contains(chunks, current, line);
        truthy(chunks, current, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        get(chunks, current, seen, line);
        get(chunks, current, key, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        get(chunks, current, out, line);
        get(chunks, current, elem, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, out, line);
}

// ── Folds ───────────────────────────────────────────────────────────────────

/// `foldRight(init) { e, acc -> }` — Kotlin's lambda takes `(element, acc)`,
/// the REVERSE of JS `reduceRight`, so the shared `__array_reduce_right`
/// cannot serve it.
pub fn emit_fold_right(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    let acc = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, acc, line);
    set(chunks, current, arr, line);
    emit_fold_right_loop(chunks, current, arr, acc, f, line);
}

/// `reduceRight { e, acc -> }` — seeded with the LAST element.
pub fn emit_reduce_right(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr, f) = pop_recv_fn(chunks, current, line);
    len_of(chunks, current, arr, line);
    truthy(chunks, current, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("Empty collection can't be reduced.", line);
    crate::emitter::nullability::emit_exception(
        chunks,
        current,
        1,
        "UnsupportedOperationException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    let acc = chunks[current].alloc_scratch(1);
    // acc = arr[len-1]; then fold the prefix.
    get(chunks, current, arr, line);
    len_of(chunks, current, arr, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    dyn_sub(chunks, current, line);
    collections::emit_get(chunks, current, line);
    set(chunks, current, acc, line);
    get(chunks, current, arr, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    len_of(chunks, current, arr, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    dyn_sub(chunks, current, line);
    collections::emit_slice(chunks, current, line);
    let prefix = chunks[current].alloc_scratch(1);
    set(chunks, current, prefix, line);
    emit_fold_right_loop(chunks, current, prefix, acc, f, line);
}

fn emit_fold_right_loop(
    chunks: &mut Vec<Chunk>,
    current: usize,
    arr: u16,
    acc: u16,
    f: u16,
    line: u32,
) {
    // Iterate a REVERSED copy so the shared forward loop serves both folds.
    let rev = chunks[current].alloc_scratch(1);
    get(chunks, current, arr, line);
    collections::emit_reverse(chunks, current, line);
    set(chunks, current, rev, line);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    for_each(chunks, current, rev, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[elem, acc], line);
        set(chunks, current, acc, line);
    });
    get(chunks, current, acc, line);
}

/// `foldIndexed(init) { i, acc, e -> }`.
pub fn emit_fold_indexed(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    let acc = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, acc, line);
    set(chunks, current, arr, line);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[idx, acc, elem], line);
        set(chunks, current, acc, line);
    });
    get(chunks, current, acc, line);
}

/// `reduceOrNull { }` — `null` on empty where `reduce` throws.
pub fn emit_reduce_or_null(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr, f) = pop_recv_fn(chunks, current, line);
    len_of(chunks, current, arr, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    emit_reduce_loop(chunks, current, arr, f, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

fn emit_reduce_loop(chunks: &mut Vec<Chunk>, current: usize, arr: u16, f: u16, line: u32) {
    let acc = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    get(chunks, current, arr, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    set(chunks, current, acc, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        get(chunks, current, idx, line);
        truthy(chunks, current, line);
        chunks[current].emit_if(line);
        call_fn(chunks, current, f, &[acc, elem], line);
        set(chunks, current, acc, line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, acc, line);
}

/// `runningFold(init) { acc, e -> }` / `scan(init) { acc, e -> }`.
pub fn emit_running_fold(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    let acc = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, acc, line);
    set(chunks, current, arr, line);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    get(chunks, current, out, line);
    get(chunks, current, acc, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[acc, elem], line);
        set(chunks, current, acc, line);
        get(chunks, current, out, line);
        get(chunks, current, acc, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    });
    get(chunks, current, out, line);
}

/// `runningReduce { acc, e -> }`.
pub fn emit_running_reduce(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr, f) = pop_recv_fn(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    let acc = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        get(chunks, current, idx, line);
        truthy(chunks, current, line);
        chunks[current].emit_if(line);
        call_fn(chunks, current, f, &[acc, elem], line);
        set(chunks, current, acc, line);
        chunks[current].emit_else(line);
        get(chunks, current, elem, line);
        set(chunks, current, acc, line);
        chunks[current].emit_end(line);
        get(chunks, current, out, line);
        get(chunks, current, acc, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    });
    get(chunks, current, out, line);
}

// ── Grouping and associating (map-producing) ────────────────────────────────

/// `groupBy { }` → `{key: [elems]}` in first-seen key order.
pub fn emit_group_by(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr, f) = pop_recv_fn(chunks, current, line);
    emit_group_by_into_new(chunks, current, arr, f, line);
}

/// `groupByTo(dest) { }`.
pub fn emit_group_by_to(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    // 3-arg form carries a VALUE transform: groupByTo(dest, keySel, valueSel).
    let vf = if argc >= 4 {
        let vf = chunks[current].alloc_scratch(1);
        set(chunks, current, vf, line);
        Some(vf)
    } else {
        None
    };
    let f = chunks[current].alloc_scratch(1);
    let dest = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, dest, line);
    if let Some(vf) = vf {
        // Transform the elements first; the grouping loop then buckets the
        // transformed values while keys come from the ORIGINAL elements —
        // so run both selectors per element inline instead.
        set(chunks, current, arr, line);
        let idx = chunks[current].alloc_scratch(1);
        let elem = chunks[current].alloc_scratch(1);
        let key = chunks[current].alloc_scratch(1);
        let bucket = chunks[current].alloc_scratch(1);
        for_each(chunks, current, arr, idx, elem, line, |chunks| {
            call_fn(chunks, current, f, &[elem], line);
            set(chunks, current, key, line);
            get(chunks, current, dest, line);
            get(chunks, current, key, line);
            dict::emit_get_dynamic(chunks, current, line);
            set(chunks, current, bucket, line);
            get(chunks, current, bucket, line);
            chunks[current].emit_op(Op::REF_IS_NULL, line);
            chunks[current].emit_if(line);
            collections::emit_array_new(chunks, current, 0, line);
            set(chunks, current, bucket, line);
            get(chunks, current, dest, line);
            get(chunks, current, key, line);
            get(chunks, current, bucket, line);
            crate::emitter::maps::emit_dict_set_tracked(chunks, current, line);
            chunks[current].emit_end(line);
            get(chunks, current, bucket, line);
            call_fn(chunks, current, vf, &[elem], line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
        });
        get(chunks, current, dest, line);
        return;
    }
    set(chunks, current, arr, line);
    emit_group_by_loop(chunks, current, arr, f, dest, line);
    get(chunks, current, dest, line);
}

fn emit_group_by_into_new(chunks: &mut Vec<Chunk>, current: usize, arr: u16, f: u16, line: u32) {
    let out = chunks[current].alloc_scratch(1);
    dict::emit_new(chunks, current, line);
    set(chunks, current, out, line);
    emit_group_by_loop(chunks, current, arr, f, out, line);
    get(chunks, current, out, line);
}

fn emit_group_by_loop(
    chunks: &mut Vec<Chunk>,
    current: usize,
    arr: u16,
    f: u16,
    out: u16,
    line: u32,
) {
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let bucket = chunks[current].alloc_scratch(1);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[elem], line);
        set(chunks, current, key, line);
        // bucket = out[key], created on first sight.
        get(chunks, current, out, line);
        get(chunks, current, key, line);
        dict::emit_get_dynamic(chunks, current, line);
        set(chunks, current, bucket, line);
        get(chunks, current, bucket, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        collections::emit_array_new(chunks, current, 0, line);
        set(chunks, current, bucket, line);
        get(chunks, current, out, line);
        get(chunks, current, key, line);
        get(chunks, current, bucket, line);
        crate::emitter::maps::emit_dict_set_tracked(chunks, current, line);
        chunks[current].emit_end(line);
        get(chunks, current, bucket, line);
        get(chunks, current, elem, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    });
}

/// `associateBy { }` → `{fn(e): e}`, last key wins. The two-lambda overload
/// `associateBy(keySelector, valueTransform)` maps both sides.
pub fn emit_associate_by(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc >= 3 {
        let vf = chunks[current].alloc_scratch(1);
        set(chunks, current, vf, line);
        let (arr, kf) = pop_recv_fn(chunks, current, line);
        let out = chunks[current].alloc_scratch(1);
        let idx = chunks[current].alloc_scratch(1);
        let elem = chunks[current].alloc_scratch(1);
        dict::emit_new(chunks, current, line);
        set(chunks, current, out, line);
        for_each(chunks, current, arr, idx, elem, line, |chunks| {
            get(chunks, current, out, line);
            call_fn(chunks, current, kf, &[elem], line);
            call_fn(chunks, current, vf, &[elem], line);
            crate::emitter::maps::emit_dict_set_tracked(chunks, current, line);
        });
        get(chunks, current, out, line);
        return;
    }
    let (arr, f) = pop_recv_fn(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    dict::emit_new(chunks, current, line);
    set(chunks, current, out, line);
    emit_associate_loop(chunks, current, arr, f, out, AssocKind::ByKey, line);
    get(chunks, current, out, line);
}

/// `mapIndexedNotNull { i, v -> }` — indexed map, nulls dropped.
pub fn emit_map_indexed_not_null(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr, f) = pop_recv_fn(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let mapped = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[idx, elem], line);
        set(chunks, current, mapped, line);
        // NOT `REF_IS_NULL`: a NUMBER result reads as null there, so every
        // kept element was dropped. Compare against a pushed null instead.
        get(chunks, current, mapped, line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        truthy(chunks, current, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        get(chunks, current, out, line);
        get(chunks, current, mapped, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, out, line);
}

/// `associateByTo(dest) { }`.
pub fn emit_associate_by_to(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    let dest = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, dest, line);
    set(chunks, current, arr, line);
    emit_associate_loop(chunks, current, arr, f, dest, AssocKind::ByKey, line);
    get(chunks, current, dest, line);
}

/// `associateWith { }` → `{e: fn(e)}`.
pub fn emit_associate_with(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr, f) = pop_recv_fn(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    dict::emit_new(chunks, current, line);
    set(chunks, current, out, line);
    emit_associate_loop(chunks, current, arr, f, out, AssocKind::WithValue, line);
    get(chunks, current, out, line);
}

/// `associate { it to … }` → the lambda returns a Pair `[k, v]`.
pub fn emit_associate(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr, f) = pop_recv_fn(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    dict::emit_new(chunks, current, line);
    set(chunks, current, out, line);
    emit_associate_loop(chunks, current, arr, f, out, AssocKind::PairResult, line);
    get(chunks, current, out, line);
}

enum AssocKind {
    ByKey,
    WithValue,
    PairResult,
}

fn emit_associate_loop(
    chunks: &mut Vec<Chunk>,
    current: usize,
    arr: u16,
    f: u16,
    out: u16,
    kind: AssocKind,
    line: u32,
) {
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let r = chunks[current].alloc_scratch(1);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[elem], line);
        set(chunks, current, r, line);
        get(chunks, current, out, line);
        match kind {
            AssocKind::ByKey => {
                get(chunks, current, r, line);
                get(chunks, current, elem, line);
            }
            AssocKind::WithValue => {
                get(chunks, current, elem, line);
                get(chunks, current, r, line);
            }
            AssocKind::PairResult => {
                get(chunks, current, r, line);
                core_wasm::i32_const(&mut chunks[current], line, 0);
                collections::emit_get(chunks, current, line);
                get(chunks, current, r, line);
                core_wasm::i32_const(&mut chunks[current], line, 1);
                collections::emit_get(chunks, current, line);
            }
        }
        crate::emitter::maps::emit_dict_set_tracked(chunks, current, line);
    });
}

// ── Pair producers ──────────────────────────────────────────────────────────

/// Build a tagged Pair from two locals, with `first`/`second` properties so
/// zip results answer both destructuring AND the property spelling.
fn emit_pair_with_props(chunks: &mut Vec<Chunk>, current: usize, a: u16, b: u16, line: u32) {
    get(chunks, current, a, line);
    get(chunks, current, b, line);
    tuples::emit_tuple(chunks, current, 2, line);
    for (prop, slot) in [("first", a), ("second", b)] {
        chunks[current].emit_dup(line);
        get(chunks, current, slot, line);
        let k =
            chunks[current].add_constant(vybe_runtime::Value::String(std::sync::Arc::from(prop)));
        chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
    }
}

/// `zip(other)` → list of tagged Pairs, truncated to the shorter side.
/// `zip(other) { a, b -> }` maps the pair through the transform instead.
pub fn emit_zip(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let f = if argc >= 3 {
        let f = chunks[current].alloc_scratch(1);
        set(chunks, current, f, line);
        Some(f)
    } else {
        None
    };
    let other = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, other, line);
    set(chunks, current, arr, line);
    if let Some(f) = f {
        return emit_zip_transform(chunks, current, arr, other, f, line);
    }
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        get(chunks, current, idx, line);
        len_of(chunks, current, other, line);
        ops::emit_dyn_lt(&mut chunks[current], line);
        truthy(chunks, current, line);
        chunks[current].emit_if(line);
        get(chunks, current, out, line);
        let b = chunks[current].alloc_scratch(1);
        elem_at(chunks, current, other, idx, line);
        set(chunks, current, b, line);
        emit_pair_with_props(chunks, current, elem, b, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, out, line);
}

fn emit_zip_transform(
    chunks: &mut Vec<Chunk>,
    current: usize,
    arr: u16,
    other: u16,
    f: u16,
    line: u32,
) {
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let b = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        get(chunks, current, idx, line);
        len_of(chunks, current, other, line);
        ops::emit_dyn_lt(&mut chunks[current], line);
        truthy(chunks, current, line);
        chunks[current].emit_if(line);
        elem_at(chunks, current, other, idx, line);
        set(chunks, current, b, line);
        get(chunks, current, out, line);
        call_fn(chunks, current, f, &[elem, b], line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, out, line);
}

/// `zipWithNext()` → Pairs of consecutive elements.
pub fn emit_zip_with_next(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let f = if argc >= 2 {
        let f = chunks[current].alloc_scratch(1);
        set(chunks, current, f, line);
        Some(f)
    } else {
        None
    };
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, arr, line);
    let _ = &f;
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let next_i = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        get(chunks, current, idx, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        ops::emit_dyn_add(&mut chunks[current], line);
        set(chunks, current, next_i, line);
        get(chunks, current, next_i, line);
        len_of(chunks, current, arr, line);
        ops::emit_dyn_lt(&mut chunks[current], line);
        truthy(chunks, current, line);
        chunks[current].emit_if(line);
        get(chunks, current, out, line);
        let b = chunks[current].alloc_scratch(1);
        elem_at(chunks, current, arr, next_i, line);
        set(chunks, current, b, line);
        match f {
            Some(f) => call_fn(chunks, current, f, &[elem, b], line),
            None => emit_pair_with_props(chunks, current, elem, b, line),
        }
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, out, line);
}

/// `unzip()` → Pair of lists from a list of Pairs.
pub fn emit_unzip(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, arr, line);
    let firsts = chunks[current].alloc_scratch(1);
    let seconds = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, firsts, line);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, seconds, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        get(chunks, current, firsts, line);
        get(chunks, current, elem, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        collections::emit_get(chunks, current, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        get(chunks, current, seconds, line);
        get(chunks, current, elem, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        collections::emit_get(chunks, current, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    });
    emit_pair_with_props(chunks, current, firsts, seconds, line);
}

/// `withIndex()` → list of `(index, value)` Pairs — destructures like
/// Kotlin's `IndexedValue`.
pub fn emit_with_index(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, arr, line);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        get(chunks, current, out, line);
        get(chunks, current, idx, line);
        get(chunks, current, elem, line);
        tuples::emit_tuple(chunks, current, 2, line);
        // Kotlin's `IndexedValue` exposes `.index`/`.value`; the tuple also
        // destructures positionally, so both spellings read the same object.
        for (prop, slot) in [("index", idx), ("value", elem)] {
            chunks[current].emit_dup(line);
            get(chunks, current, slot, line);
            let k = chunks[current]
                .add_constant(vybe_runtime::Value::String(std::sync::Arc::from(prop)));
            chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
        }
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    });
    get(chunks, current, out, line);
}

/// `partition { }` → `Pair(matching, rest)`.
pub fn emit_partition(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr0, f) = pop_recv_fn(chunks, current, line);
    // Sets/maps are dicts — iterate the list VIEW, not the raw receiver.
    let arr = emit_list_view(chunks, current, arr0, line);
    let yes = chunks[current].alloc_scratch(1);
    let no = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, yes, line);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, no, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[elem], line);
        truthy(chunks, current, line);
        chunks[current].emit_if_value(line);
        get(chunks, current, yes, line);
        chunks[current].emit_else(line);
        get(chunks, current, no, line);
        chunks[current].emit_end(line);
        get(chunks, current, elem, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    });
    emit_pair_with_props(chunks, current, yes, no, line);
}

// ── Non-HOF list ops that still need composition ────────────────────────────

/// `average()`.
pub fn emit_average(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, arr, line);
    get(chunks, current, arr, line);
    collections::emit_sum(chunks, current, line);
    len_of(chunks, current, arr, line);
    dyn_div(chunks, current, line);
}

/// `takeLast(n)` / `dropLast(n)`.
pub fn emit_take_last(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let n = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, n, line);
    set(chunks, current, arr, line);
    // slice(len - n, len), floored at 0.
    get(chunks, current, arr, line);
    emit_len_minus_clamped(chunks, current, arr, n, line);
    len_of(chunks, current, arr, line);
    collections::emit_slice(chunks, current, line);
}

pub fn emit_drop_last(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let n = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, n, line);
    set(chunks, current, arr, line);
    // slice(0, len - n), floored at 0.
    get(chunks, current, arr, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    emit_len_minus_clamped(chunks, current, arr, n, line);
    collections::emit_slice(chunks, current, line);
}

/// Push `max(len - n, 0)`.
fn emit_len_minus_clamped(chunks: &mut Vec<Chunk>, current: usize, arr: u16, n: u16, line: u32) {
    let cut = chunks[current].alloc_scratch(1);
    len_of(chunks, current, arr, line);
    get(chunks, current, n, line);
    dyn_sub(chunks, current, line);
    set(chunks, current, cut, line);
    get(chunks, current, cut, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_lt(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_else(line);
    get(chunks, current, cut, line);
    chunks[current].emit_end(line);
}

/// `flatten()` — one level.
pub fn emit_flatten(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let arr0 = chunks[current].alloc_scratch(1);
    set(chunks, current, arr0, line);
    // Sets/maps at EITHER level are dicts — iterate list views.
    let arr = emit_list_view(chunks, current, arr0, line);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem0 = chunks[current].alloc_scratch(1);
    let jdx = chunks[current].alloc_scratch(1);
    let inner = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for_each(chunks, current, arr, idx, elem0, line, |chunks| {
        let elem = emit_list_view(chunks, current, elem0, line);
        for_each(chunks, current, elem, jdx, inner, line, |chunks| {
            get(chunks, current, out, line);
            get(chunks, current, inner, line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
        });
    });
    get(chunks, current, out, line);
}

/// `orEmpty()` — `null` receiver becomes `[]`.
pub fn emit_or_empty(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    set(chunks, current, v, line);
    get(chunks, current, v, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_else(line);
    get(chunks, current, v, line);
    chunks[current].emit_end(line);
}

/// `sortedDescending()`.
pub fn emit_sorted_descending(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    collections::emit_sorted(chunks, current, line);
    collections::emit_reverse(chunks, current, line);
}

/// `toMap()` — a list of Pairs into an insertion-ordered map; on a map
/// receiver, an independent copy.
pub fn emit_to_map(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let recv = chunks[current].alloc_scratch(1);
    set(chunks, current, recv, line);
    get(chunks, current, recv, line);
    let is_arr = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_arr, 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, recv, line);
    emit_pairs_to_map(chunks, current, argc, line);
    chunks[current].emit_else(line);
    get(chunks, current, recv, line);
    crate::emitter::maps::emit_copy_map(chunks, current, argc, line);
    chunks[current].emit_end(line);
}

fn emit_pairs_to_map(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, arr, line);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    dict::emit_new(chunks, current, line);
    set(chunks, current, out, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        get(chunks, current, out, line);
        get(chunks, current, elem, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        collections::emit_get(chunks, current, line);
        get(chunks, current, elem, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        collections::emit_get(chunks, current, line);
        crate::emitter::maps::emit_dict_set_tracked(chunks, current, line);
    });
    get(chunks, current, out, line);
}

// ── Receiver-kind helpers ───────────────────────────────────────────────────

/// When `arr_slot` holds a dict (Map/Set — carries `__keys`), replace it with
/// a LIST view: a Set's keys, a Map's `[k, v]` entry pairs. A predicate over a
/// Map iterates entries in Kotlin, and one loop shape serves all receivers.
fn emit_entries_if_dict(chunks: &mut Vec<Chunk>, current: usize, arr_slot: u16, line: u32) -> u16 {
    let out = chunks[current].alloc_scratch(1);
    get(chunks, current, arr_slot, line);
    let idx = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(idx, 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, arr_slot, line);
    chunks[current].emit_else(line);
    get(chunks, current, arr_slot, line);
    crate::emitter::collections::emit_dict_as_list(chunks, current, line);
    chunks[current].emit_end(line);
    set(chunks, current, out, line);
    out
}

/// A list view of ANY receiver: arrays pass through, Sets become their keys,
/// Maps their entry objects.
fn emit_list_view(chunks: &mut Vec<Chunk>, current: usize, slot: u16, line: u32) -> u16 {
    emit_entries_if_dict(chunks, current, slot, line)
}

/// `size`/`length`/`count()` for a DYNAMIC receiver — the runtime half of the
/// `len` slot (builtinslotplan §2c): string → UTF-16 length, array → element
/// count, dict-backed Map/Set → key count. A statically-typed receiver never
/// reaches this; the slot binding emits the direct op.
pub fn emit_size_any(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    set(chunks, current, v, line);
    get(chunks, current, v, line);
    let type_of = chunks[current].add_import("ecma:value", "typeof");
    chunks[current].emit_call(type_of, 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, v, line);
    vybe_compiler::primitives::strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(chunks, current, v, line);
    let idx = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(idx, 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, v, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_else(line);
    // A StringBuilder carries its text in `__buffer`; its `length` is the
    // buffer's, not a key count.
    get(chunks, current, v, line);
    let buf_key = chunks[current].add_constant(vybe_runtime::Value::String(std::sync::Arc::from(
        "__buffer",
    )));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, buf_key, line);
    let buf = chunks[current].alloc_scratch(1);
    set(chunks, current, buf, line);
    get(chunks, current, buf, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, buf, line);
    vybe_compiler::primitives::strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(chunks, current, v, line);
    let tag = chunks[current].add_import("ecma:object", "toStringTag");
    chunks[current].emit_call(tag, 1, line);
    chunks[current].emit_string_const("[object Set]", line);
    let eq = chunks[current].add_import("wasm:js-string", "equals");
    chunks[current].emit_call(eq, 2, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, v, line);
    sets::emit_size(chunks, current, line);
    chunks[current].emit_else(line);
    get(chunks, current, v, line);
    dict::emit_method_size(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

// ── Grouping ────────────────────────────────────────────────────────────────

/// The properties a `groupingBy { }` receiver carries: the source list and
/// the key selector. `Grouping<T, K>` is nothing but this pair — Kotlin's own
/// definition is "a source paired with a keyOf function" — and the terminal
/// ops (`eachCount`, `fold`, `reduce`, `aggregate`) read both.
const GROUPING_SRC: &str = "__kt_grouping_src";
const GROUPING_FN: &str = "__kt_grouping_fn";

/// `groupingBy { }`.
pub fn emit_grouping_by(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr, f) = pop_recv_fn(chunks, current, line);
    chunks[current].emit_struct_new(0, 0, line);
    for (prop, slot) in [(GROUPING_SRC, arr), (GROUPING_FN, f)] {
        chunks[current].emit_dup(line);
        get(chunks, current, slot, line);
        let k =
            chunks[current].add_constant(vybe_runtime::Value::String(std::sync::Arc::from(prop)));
        chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
    }
}

/// Pop `[grouping]` into `(src, keyfn)` locals.
fn pop_grouping(chunks: &mut Vec<Chunk>, current: usize, line: u32) -> (u16, u16) {
    let g = chunks[current].alloc_scratch(1);
    let src = chunks[current].alloc_scratch(1);
    let f = chunks[current].alloc_scratch(1);
    set(chunks, current, g, line);
    for (prop, slot) in [(GROUPING_SRC, src), (GROUPING_FN, f)] {
        get(chunks, current, g, line);
        let k =
            chunks[current].add_constant(vybe_runtime::Value::String(std::sync::Arc::from(prop)));
        chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
        set(chunks, current, slot, line);
    }
    (src, f)
}

/// `eachCount()`.
pub fn emit_each_count(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (src, f) = pop_grouping(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    dict::emit_new(chunks, current, line);
    set(chunks, current, out, line);
    for_each(chunks, current, src, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[elem], line);
        set(chunks, current, key, line);
        get(chunks, current, out, line);
        get(chunks, current, key, line);
        dict::emit_get_dynamic(chunks, current, line);
        set(chunks, current, n, line);
        get(chunks, current, n, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        set(chunks, current, n, line);
        chunks[current].emit_end(line);
        get(chunks, current, out, line);
        get(chunks, current, key, line);
        get(chunks, current, n, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        ops::emit_dyn_add(&mut chunks[current], line);
        crate::emitter::maps::emit_dict_set_tracked(chunks, current, line);
    });
    get(chunks, current, out, line);
}

/// `fold(init) { acc, e -> }` on a grouping.
pub fn emit_grouping_fold(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    let init = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, init, line);
    let (src, keyfn) = pop_grouping(chunks, current, line);
    emit_grouping_accumulate(
        chunks,
        current,
        src,
        keyfn,
        GroupAcc::Fold { init, f },
        line,
    );
}

/// `reduce { key, acc, e -> }` on a grouping.
pub fn emit_grouping_reduce(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    let (src, keyfn) = pop_grouping(chunks, current, line);
    emit_grouping_accumulate(chunks, current, src, keyfn, GroupAcc::Reduce { f }, line);
}

/// `aggregate { key, acc, e, first -> }` on a grouping.
pub fn emit_grouping_aggregate(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    let (src, keyfn) = pop_grouping(chunks, current, line);
    emit_grouping_accumulate(chunks, current, src, keyfn, GroupAcc::Aggregate { f }, line);
}

enum GroupAcc {
    Fold { init: u16, f: u16 },
    Reduce { f: u16 },
    Aggregate { f: u16 },
}

fn emit_grouping_accumulate(
    chunks: &mut Vec<Chunk>,
    current: usize,
    src: u16,
    keyfn: u16,
    kind: GroupAcc,
    line: u32,
) {
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let acc = chunks[current].alloc_scratch(1);
    let fresh = chunks[current].alloc_scratch(1);
    dict::emit_new(chunks, current, line);
    set(chunks, current, out, line);
    for_each(chunks, current, src, idx, elem, line, |chunks| {
        call_fn(chunks, current, keyfn, &[elem], line);
        set(chunks, current, key, line);
        // fresh = !out.has(key); acc = out[key] (null when fresh)
        get(chunks, current, out, line);
        get(chunks, current, key, line);
        dict::emit_method_has(chunks, current, line);
        truthy(chunks, current, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        set(chunks, current, fresh, line);
        get(chunks, current, out, line);
        get(chunks, current, key, line);
        dict::emit_get_dynamic(chunks, current, line);
        set(chunks, current, acc, line);

        get(chunks, current, out, line);
        get(chunks, current, key, line);
        match kind {
            GroupAcc::Fold { init, f } => {
                // acc starts at init on first sight of the key.
                get(chunks, current, fresh, line);
                truthy(chunks, current, line);
                chunks[current].emit_if(line);
                get(chunks, current, init, line);
                set(chunks, current, acc, line);
                chunks[current].emit_end(line);
                call_fn(chunks, current, f, &[acc, elem], line);
            }
            GroupAcc::Reduce { f } => {
                // First element IS the accumulator; later ones reduce in.
                get(chunks, current, fresh, line);
                truthy(chunks, current, line);
                chunks[current].emit_if_value(line);
                get(chunks, current, elem, line);
                chunks[current].emit_else(line);
                call_fn(chunks, current, f, &[key, acc, elem], line);
                chunks[current].emit_end(line);
            }
            GroupAcc::Aggregate { f } => {
                get(chunks, current, fresh, line);
                ops::emit_i32_to_bool(&mut chunks[current], line);
                let flag = chunks[current].alloc_scratch(1);
                set(chunks, current, flag, line);
                call_fn(chunks, current, f, &[key, acc, elem, flag], line);
            }
        }
        crate::emitter::maps::emit_dict_set_tracked(chunks, current, line);
    });
    get(chunks, current, out, line);
}

// ── Factories ───────────────────────────────────────────────────────────────

/// `listOfNotNull(…)` — the non-null arguments, in order.
pub fn emit_list_of_not_null(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let args = chunks[current].alloc_scratch(argc.max(1) as u16);
    for i in (0..argc).rev() {
        set(chunks, current, args + i as u16, line);
    }
    let out = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for i in 0..argc {
        get(chunks, current, args + i as u16, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        get(chunks, current, out, line);
        get(chunks, current, args + i as u16, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    }
    get(chunks, current, out, line);
}

/// `arrayOfNulls(n)` — `n` nulls.
pub fn emit_null_array(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_sized_array(chunks, current, argc, /*null_fill:*/ true, line);
}

/// `IntArray(n)` / `DoubleArray(n)` — `n` zeros. With an init lambda,
/// `IntArray(n) { i -> }`.
pub fn emit_zeroed_array(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_sized_array(chunks, current, argc, /*null_fill:*/ false, line);
}

fn emit_sized_array(chunks: &mut Vec<Chunk>, current: usize, argc: u8, null_fill: bool, line: u32) {
    let f = if argc >= 2 {
        let f = chunks[current].alloc_scratch(1);
        set(chunks, current, f, line);
        Some(f)
    } else {
        None
    };
    let n = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    set(chunks, current, n, line);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(chunks, current, i, line);
    let _block = chunks[current].emit_block(line);
    let (_loop, _) = chunks[current].emit_loop_s(line);
    get(chunks, current, i, line);
    get(chunks, current, n, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(chunks, current, out, line);
    if let Some(f) = f {
        call_fn(chunks, current, f, &[i], line);
    } else if null_fill {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(chunks, current, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(chunks, current, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    get(chunks, current, out, line);
}

/// `distinct()` — first occurrence of each value, order preserved.
pub fn emit_distinct(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, arr, line);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        get(chunks, current, out, line);
        get(chunks, current, elem, line);
        collections::emit_contains(chunks, current, line);
        truthy(chunks, current, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        get(chunks, current, out, line);
        get(chunks, current, elem, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, out, line);
}

/// `toSortedSet()` — sorted elements, as a Kotlin Set.
pub fn emit_to_sorted_set(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    crate::emitter::collections::emit_dict_as_list(chunks, current, line);
    collections::emit_sorted(chunks, current, line);
    crate::emitter::collections::emit_to_set(chunks, current, 1, line);
}

/// `sortedSetOf(v…)` — the arguments, sorted, as a Kotlin Set.
pub fn emit_sorted_set_of(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(argc.max(1) as u16);
    collections::emit_pack_n(chunks, current, argc as u16, base, line);
    collections::emit_sorted(chunks, current, line);
    crate::emitter::collections::emit_to_set(chunks, current, 1, line);
}

/// `sortedMapOf(k to v, …)` — the pairs, as a map with sorted keys.
pub fn emit_sorted_map_of(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(argc.max(1) as u16);
    collections::emit_pack_n(chunks, current, argc as u16, base, line);
    emit_to_map(chunks, current, 1, line);
    crate::emitter::maps::emit_to_sorted_map(chunks, current, 1, line);
}

// ── Receiver-dispatching HOFs (`x.filter { }` for every collection kind) ────
//
// Kotlin's `filter` returns a Map on a Map and a List on a List or Set; `map`
// and `forEach` iterate entries on a Map. No compile-time HOF over one shape
// can do that, so the walker rewrites these three (plus `filterNot`) to the
// adapters below, which branch on the receiver at runtime.

/// Push i32 `1` when the value in `slot` carries the Kotlin Set marker.
fn is_set_marked(chunks: &mut Vec<Chunk>, current: usize, slot: u16, line: u32) {
    get(chunks, current, slot, line);
    let k = chunks[current].add_constant(vybe_runtime::Value::String(std::sync::Arc::from(
        crate::emitter::tostring::SET_MARKER,
    )));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
}

fn is_ecma_set(chunks: &mut Vec<Chunk>, current: usize, slot: u16, line: u32) {
    get(chunks, current, slot, line);
    let tag = chunks[current].add_import("ecma:object", "toStringTag");
    chunks[current].emit_call(tag, 1, line);
    chunks[current].emit_string_const("[object Set]", line);
    let eq = chunks[current].add_import("wasm:js-string", "equals");
    chunks[current].emit_call(eq, 2, line);
}

fn materialize_set_values(chunks: &mut Vec<Chunk>, current: usize, slot: u16, line: u32) -> u16 {
    let values = chunks[current].alloc_scratch(1);
    get(chunks, current, slot, line);
    crate::emitter::collections::emit_to_list(chunks, current, 1, line);
    set(chunks, current, values, line);
    values
}

/// The filter loop over the list in `arr`. Leaves the result list.
fn filter_loop(chunks: &mut Vec<Chunk>, current: usize, arr: u16, f: u16, invert: bool, line: u32) {
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[elem], line);
        truthy(chunks, current, line);
        if invert {
            chunks[current].emit_op(Op::I32_EQZ, line);
        }
        chunks[current].emit_if(line);
        get(chunks, current, out, line);
        get(chunks, current, elem, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, out, line);
}

fn emit_kt_filter_impl(chunks: &mut Vec<Chunk>, current: usize, invert: bool, line: u32) {
    let (recv, f) = pop_recv_fn(chunks, current, line);
    // `"abc".filter { }` filters CHARACTERS and returns a STRING.
    get(chunks, current, recv, line);
    let type_of = chunks[current].add_import("ecma:value", "typeof");
    chunks[current].emit_call(type_of, 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    let chars = chunks[current].alloc_scratch(1);
    get(chunks, current, recv, line);
    chunks[current].emit_string_const("", line);
    let split = chunks[current].add_import("ecma:string", "split");
    chunks[current].emit_call(split, 2, line);
    set(chunks, current, chars, line);
    filter_loop(chunks, current, chars, f, invert, line);
    chunks[current].emit_string_const("", line);
    collections::emit_join(chunks, current, line);
    chunks[current].emit_else(line);
    get(chunks, current, recv, line);
    let is_arr = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_arr, 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    filter_loop(chunks, current, recv, f, invert, line);
    chunks[current].emit_else(line);
    is_set_marked(chunks, current, recv, line);
    chunks[current].emit_if_value(line);
    // Set.filter → List of matching ELEMENTS (the values view — `__keys`
    // holds string spellings plus the marker itself).
    let keys = materialize_set_values(chunks, current, recv, line);
    filter_loop(chunks, current, keys, f, invert, line);
    chunks[current].emit_else(line);
    is_ecma_set(chunks, current, recv, line);
    chunks[current].emit_if_value(line);
    let values = materialize_set_values(chunks, current, recv, line);
    filter_loop(chunks, current, values, f, invert, line);
    chunks[current].emit_else(line);
    // Map.filter → Map of matching entries.
    get(chunks, current, recv, line);
    get(chunks, current, f, line);
    if invert {
        crate::emitter::maps::emit_map_filter_not(chunks, current, 2, line);
    } else {
        crate::emitter::maps::emit_map_filter(chunks, current, 2, line);
    }
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_kt_filter(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_kt_filter_impl(chunks, current, false, line);
}

pub fn emit_kt_filter_not(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_kt_filter_impl(chunks, current, true, line);
}

/// The map loop over the list in `arr`. Leaves the result list.
fn map_loop(chunks: &mut Vec<Chunk>, current: usize, arr: u16, f: u16, line: u32) {
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        get(chunks, current, out, line);
        call_fn(chunks, current, f, &[elem], line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    });
    get(chunks, current, out, line);
}

pub fn emit_kt_map_hof(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (recv, f) = pop_recv_fn(chunks, current, line);
    // `"abc".map { }` maps over the characters (result: List).
    get(chunks, current, recv, line);
    let type_of = chunks[current].add_import("ecma:value", "typeof");
    chunks[current].emit_call(type_of, 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    let chars = chunks[current].alloc_scratch(1);
    get(chunks, current, recv, line);
    chunks[current].emit_string_const("", line);
    let split = chunks[current].add_import("ecma:string", "split");
    chunks[current].emit_call(split, 2, line);
    set(chunks, current, chars, line);
    map_loop(chunks, current, chars, f, line);
    chunks[current].emit_else(line);
    get(chunks, current, recv, line);
    let is_arr = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_arr, 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    map_loop(chunks, current, recv, f, line);
    chunks[current].emit_else(line);
    is_set_marked(chunks, current, recv, line);
    chunks[current].emit_if_value(line);
    let keys = materialize_set_values(chunks, current, recv, line);
    map_loop(chunks, current, keys, f, line);
    chunks[current].emit_else(line);
    is_ecma_set(chunks, current, recv, line);
    chunks[current].emit_if_value(line);
    let values = materialize_set_values(chunks, current, recv, line);
    map_loop(chunks, current, values, f, line);
    chunks[current].emit_else(line);
    get(chunks, current, recv, line);
    get(chunks, current, f, line);
    crate::emitter::maps::emit_map_to_list(chunks, current, 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_kt_for_each(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (recv, f) = pop_recv_fn(chunks, current, line);
    get(chunks, current, recv, line);
    let is_arr = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_arr, 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    for_each(chunks, current, recv, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[elem], line);
        chunks[current].emit_op(Op::DROP, line);
    });
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    is_set_marked(chunks, current, recv, line);
    chunks[current].emit_if_value(line);
    let view = materialize_set_values(chunks, current, recv, line);
    let idx2 = chunks[current].alloc_scratch(1);
    let elem2 = chunks[current].alloc_scratch(1);
    for_each(chunks, current, view, idx2, elem2, line, |chunks| {
        call_fn(chunks, current, f, &[elem2], line);
        chunks[current].emit_op(Op::DROP, line);
    });
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    is_ecma_set(chunks, current, recv, line);
    chunks[current].emit_if_value(line);
    let view_set = materialize_set_values(chunks, current, recv, line);
    let idx_set = chunks[current].alloc_scratch(1);
    let elem_set = chunks[current].alloc_scratch(1);
    for_each(
        chunks,
        current,
        view_set,
        idx_set,
        elem_set,
        line,
        |chunks| {
            call_fn(chunks, current, f, &[elem_set], line);
            chunks[current].emit_op(Op::DROP, line);
        },
    );
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    get(chunks, current, recv, line);
    get(chunks, current, f, line);
    crate::emitter::maps::emit_map_for_each(chunks, current, 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `reduce { }` — Kotlin throws `UnsupportedOperationException` on an empty
/// receiver where the shared array HOF answered `null`.
pub fn emit_reduce_throwing(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr, f) = pop_recv_fn(chunks, current, line);
    // Sets and Maps reduce over their list view.
    let arr = emit_list_view(chunks, current, arr, line);
    len_of(chunks, current, arr, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    emit_reduce_loop(chunks, current, arr, f, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("Empty collection can't be reduced.", line);
    crate::emitter::nullability::emit_exception(
        chunks,
        current,
        1,
        "UnsupportedOperationException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

/// `slice(range)` for strings AND lists — `"abcd".slice(1..2)` is `"bc"`,
/// a list slices to a list. `[recv, from, to)`.
pub fn emit_slice_any(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let to = chunks[current].alloc_scratch(1);
    let from = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    set(chunks, current, to, line);
    set(chunks, current, from, line);
    set(chunks, current, recv, line);
    get(chunks, current, recv, line);
    let type_of = chunks[current].add_import("ecma:value", "typeof");
    chunks[current].emit_call(type_of, 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    // Kotlin's slice THROWS out of range; ECMA's clamps.
    {
        let bad = chunks[current].alloc_scratch(1);
        get(chunks, current, from, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        ops::emit_dyn_lt(&mut chunks[current], line);
        truthy(chunks, current, line);
        set(chunks, current, bad, line);
        get(chunks, current, to, line);
        get(chunks, current, recv, line);
        vybe_compiler::primitives::strings::emit_length(&mut chunks[current], line);
        ops::emit_dyn_gt(&mut chunks[current], line);
        truthy(chunks, current, line);
        get(chunks, current, bad, line);
        chunks[current].emit_op(Op::I32_OR, line);
        chunks[current].emit_if(line);
        chunks[current].emit_string_const("index out of bounds", line);
        crate::emitter::nullability::emit_exception(
            chunks,
            current,
            1,
            "IndexOutOfBoundsException",
            line,
        );
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_end(line);
    }
    get(chunks, current, recv, line);
    get(chunks, current, from, line);
    get(chunks, current, to, line);
    let sl = chunks[current].add_import("ecma:string", "slice");
    chunks[current].emit_call(sl, 3, line);
    chunks[current].emit_else(line);
    get(chunks, current, recv, line);
    get(chunks, current, from, line);
    get(chunks, current, to, line);
    collections::emit_slice(chunks, current, line);
    chunks[current].emit_end(line);
}

/// `takeLastWhile { }` / `dropLastWhile { }` — the same split, from the END.
pub fn emit_last_while_split(chunks: &mut Vec<Chunk>, current: usize, take: bool, line: u32) {
    let (recv, f) = pop_recv_fn(chunks, current, line);
    let is_str = chunks[current].alloc_scratch(1);
    get(chunks, current, recv, line);
    let type_of = chunks[current].add_import("ecma:value", "typeof");
    chunks[current].emit_call(type_of, 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    set(chunks, current, is_str, line);
    let arr = chunks[current].alloc_scratch(1);
    get(chunks, current, recv, line);
    crate::emitter::collections::emit_dict_as_list(chunks, current, line);
    collections::emit_reverse(chunks, current, line);
    set(chunks, current, arr, line);

    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let active = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    chunks[current].emit_bool_const(true, line);
    set(chunks, current, active, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        get(chunks, current, active, line);
        truthy(chunks, current, line);
        chunks[current].emit_if(line);
        call_fn(chunks, current, f, &[elem], line);
        truthy(chunks, current, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        chunks[current].emit_bool_const(false, line);
        set(chunks, current, active, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        get(chunks, current, active, line);
        truthy(chunks, current, line);
        if !take {
            chunks[current].emit_op(Op::I32_EQZ, line);
        }
        chunks[current].emit_if(line);
        get(chunks, current, out, line);
        get(chunks, current, elem, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    });
    // Built backwards — restore source order.
    get(chunks, current, out, line);
    collections::emit_reverse(chunks, current, line);
    set(chunks, current, out, line);
    get(chunks, current, is_str, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, out, line);
    chunks[current].emit_string_const("", line);
    collections::emit_join(chunks, current, line);
    chunks[current].emit_else(line);
    get(chunks, current, out, line);
    chunks[current].emit_end(line);
}

/// `foldRightIndexed(init) { i, v, acc -> }` — walk-time reversed the last
/// two args like `fold`, so the adapter receives `(f, init)`.
pub fn emit_fold_right_indexed(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    let acc = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, acc, line);
    set(chunks, current, arr, line);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    len_of(chunks, current, arr, line);
    set(chunks, current, n, line);
    // i = n-1 down to 0
    get(chunks, current, n, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    dyn_sub(chunks, current, line);
    set(chunks, current, i, line);
    let _block = chunks[current].emit_block(line);
    let (_loop, _) = chunks[current].emit_loop_s(line);
    get(chunks, current, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_ge(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    elem_at(chunks, current, arr, i, line);
    set(chunks, current, elem, line);
    call_fn(chunks, current, f, &[i, elem, acc], line);
    set(chunks, current, acc, line);
    get(chunks, current, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    dyn_sub(chunks, current, line);
    set(chunks, current, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    get(chunks, current, acc, line);
}

/// `runningFoldIndexed(init) { i, acc, e -> }` — prefix results, indexed.
pub fn emit_running_fold_indexed(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    let acc = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, acc, line);
    set(chunks, current, arr, line);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    get(chunks, current, out, line);
    get(chunks, current, acc, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    for_each(chunks, current, arr, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[idx, acc, elem], line);
        set(chunks, current, acc, line);
        get(chunks, current, out, line);
        get(chunks, current, acc, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    });
    get(chunks, current, out, line);
}

/// `Array(n) { init }` / `IntArray(n) { init }` / `List(n) { init }` — size
/// plus per-index initializer. The 1-arg numeric-array form zero-fills.
/// Stack: [n, (f)] → [array].
pub fn emit_array_init(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let f = if argc >= 2 {
        let f = chunks[current].alloc_scratch(1);
        set(chunks, current, f, line);
        Some(f)
    } else {
        None
    };
    let n = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    set(chunks, current, n, line);
    collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(chunks, current, i, line);

    let _block = chunks[current].emit_block(line);
    let (_loop, _) = chunks[current].emit_loop_s(line);
    get(chunks, current, i, line);
    get(chunks, current, n, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    get(chunks, current, out, line);
    if let Some(f) = f {
        call_fn(chunks, current, f, &[i], line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(chunks, current, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(chunks, current, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    get(chunks, current, out, line);
}

/// Kotlin `array.fill(value, from = 0, to = size)` — argument order differs
/// from Java's `Arrays.fill(a, from, to, value)`, which the jvm adapter
/// expects; routing Kotlin's member spelling there filled the wrong range.
/// Stack: [arr, value, (from), (to)] → [] (fills in place, answers Unit).
pub fn emit_fill(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let to = chunks[current].alloc_scratch(1);
    let from = chunks[current].alloc_scratch(1);
    let v = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    let has_range = argc >= 4;
    if has_range {
        set(chunks, current, to, line);
        set(chunks, current, from, line);
    }
    set(chunks, current, v, line);
    set(chunks, current, arr, line);
    get(chunks, current, arr, line);
    get(chunks, current, v, line);
    if has_range {
        get(chunks, current, from, line);
        get(chunks, current, to, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        len_of(chunks, current, arr, line);
    }
    collections::emit_fill(chunks, current, line);
}

/// `removeAll { p }` / `retainAll { p }` / `removeIf { p }` — predicate
/// forms. Answers whether anything was removed.
/// Stack: [recv, f] → [bool].
pub fn emit_remove_if(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_mutate_if(chunks, current, argc, /*invert:*/ false, line);
}

pub fn emit_retain_if(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_mutate_if(chunks, current, argc, /*invert:*/ true, line);
}

fn emit_mutate_if(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, invert: bool, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    let values = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let changed = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, recv, line);
    chunks[current].emit_bool_const(false, line);
    set(chunks, current, changed, line);
    // Snapshot the element view first — removal mutates the receiver.
    get(chunks, current, recv, line);
    crate::emitter::collections::emit_dict_as_list(chunks, current, line);
    collections::emit_clone(chunks, current, line);
    set(chunks, current, values, line);
    for_each(chunks, current, values, idx, elem, line, |chunks| {
        call_fn(chunks, current, f, &[elem], line);
        truthy(chunks, current, line);
        if invert {
            chunks[current].emit_op(Op::I32_EQZ, line);
        }
        chunks[current].emit_if(line);
        get(chunks, current, recv, line);
        get(chunks, current, elem, line);
        crate::emitter::maps::emit_remove_any(chunks, current, 2, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_bool_const(true, line);
        set(chunks, current, changed, line);
        chunks[current].emit_end(line);
    });
    get(chunks, current, changed, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `x.ifEmpty { fallback }` — the receiver when non-empty, else the
/// lambda's value. Strings, lists and dict-backed sets all answer their own
/// emptiness through `kotlin.is_empty`.
pub fn emit_if_empty(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, recv, line);
    get(chunks, current, recv, line);
    crate::emitter::collections::emit_is_empty(chunks, current, 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    call_fn(chunks, current, f, &[], line);
    chunks[current].emit_else(line);
    get(chunks, current, recv, line);
    chunks[current].emit_end(line);
}
