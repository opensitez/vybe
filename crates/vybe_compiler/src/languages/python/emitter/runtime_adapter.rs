//! Python runtime-surface emitters.
//!
//! These are routed from the Python profile through `common:python.*`.
//! Keep Python-specific call shapes here instead of sending them through
//! the old runtime-helper function table.

use crate::emitter::{collections, target::Target};
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

/// Python `print(...)` — inline emitter that writes to `wasi:cli/stdout`
/// (like PHP `echo`), so `sep`/`end` and the missing trailing newline are
/// all expressible (the line-oriented `wasi:logging/logging.log` sink cannot
/// control separators or suppress the newline).
///
/// Argument convention set by the Python walker: when args are present,
/// `arg0 = sep`, `arg1 = end`, `args[2..] = items`. A bare `print()` compiles
/// with `argc == 0` and emits just the default end (`"\n"`). Each item is
/// converted to its Python display form via `emit_py_repr`.
pub fn emit_print(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let write_idx = chunk.add_import("wasi:cli/stdout", "write-via-stream");
    let rd_slot = chunk.alloc_scratch(1);
    let wr_slot = chunk.alloc_scratch(1);
    let result_slot = chunk.alloc_scratch(1);

    if argc == 0 {
        // Bare `print()` → just the default line terminator.
        chunk.emit_string_const("\n", line);
        chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    } else {
        // Args are on the stack in call order (top = last). Save into slots.
        let mut slots = Vec::with_capacity(argc as usize);
        for _ in 0..argc {
            let s = chunk.alloc_scratch(1);
            chunk.emit_op_u16(Op::LOCAL_SET, s, line);
            slots.push(s);
        }
        slots.reverse(); // slots[0]=sep, slots[1]=end, slots[2..]=items
        let sep_slot = slots[0];
        let end_slot = slots[1];

        // Build the output string: item0, sep, item1, sep, …, end.
        let mut part_count = 0usize;
        for (i, &item) in slots[2..].iter().enumerate() {
            if i > 0 {
                chunk.emit_op_u16(Op::LOCAL_GET, sep_slot, line);
                part_count += 1;
            }
            chunk.emit_op_u16(Op::LOCAL_GET, item, line);
            emit_py_repr(chunk, line);
            part_count += 1;
        }
        chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
        part_count += 1;

        crate::emitter::strings::emit_concat(chunk, part_count, line);
        chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    }

    crate::emitter::io::emit_write_stdout_with_imports(
        chunk,
        write_idx,
        rd_slot,
        wr_slot,
        line,
        |c| c.emit_op_u16(Op::LOCAL_GET, result_slot, line),
    );
}

/// Python `+` operator: array→concat, else→dynamic add.
pub fn emit_pyadd(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let is_array = chunk.add_import("ecma:array", "isArray");
    let concat = chunk.add_import("ecma:array", "concat");
    let b_slot = chunk.alloc_scratch(1);
    let a_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a_slot, line);
    // if isArray(a)
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_if_value(line);
    // array concat: concat(a, b)
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_call(concat, 2, line);
    chunk.emit_else(line);
    // dynamic add (string concat or numeric add)
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_end(line);
}

/// Python `str(x)` — the same display form `print` uses (dict → `{'k': v}`,
/// list → `[..]`, True/False/None), so `str({'a':1})` isn't `[object Object]`.
/// Stack: `[x]` → `[string]`.
pub fn emit_str(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        chunk.emit_string_const("", line);
        return;
    }
    emit_py_repr(chunk, line);
}

/// Inline Python repr: Bool→True/False, None→None, Array→[elem, ...], else passthrough.
fn emit_py_repr(chunk: &mut Chunk, line: u32) {
    let test_bool = chunk.add_import("wasm:js-boolean", "test");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let is_array = chunk.add_import("ecma:array", "isArray");
    let json_str = chunk.add_import("ecma:json", "stringify");
    let replace_all = chunk.add_import("ecma:string", "replaceAll");
    let scratch = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, scratch, line);

    // null → "None"
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("None", line);
    chunk.emit_else(line);

    // bool → "True"/"False"
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_call(test_bool, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("True", line);
    chunk.emit_else(line);
    chunk.emit_string_const("False", line);
    chunk.emit_end(line);
    chunk.emit_else(line);

    // array → JSON stringify then fix spacing + Python bool/None literals
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_call(json_str, 1, line);
    // Fix spacing (`: ` for dicts nested inside the list, then `, `)
    chunk.emit_string_const(":", line);
    chunk.emit_string_const(": ", line);
    chunk.emit_call(replace_all, 3, line);
    chunk.emit_string_const(",", line);
    chunk.emit_string_const(", ", line);
    chunk.emit_call(replace_all, 3, line);
    // Fix Python bool/None capitalization
    chunk.emit_string_const("true", line);
    chunk.emit_string_const("True", line);
    chunk.emit_call(replace_all, 3, line);
    chunk.emit_string_const("false", line);
    chunk.emit_string_const("False", line);
    chunk.emit_call(replace_all, 3, line);
    chunk.emit_string_const("null", line);
    chunk.emit_string_const("None", line);
    chunk.emit_call(replace_all, 3, line);
    // Python uses single quotes for strings inside lists
    chunk.emit_string_const("\"", line);
    chunk.emit_string_const("'", line);
    chunk.emit_call(replace_all, 3, line);
    chunk.emit_else(line);

    // Not array: a string/number coerces straight to string; anything else is
    // treated as a dict — `{'k': v, ...}` via JSON + Python spacing/quotes.
    let is_number = chunk.add_import("wasm:js-number", "test");
    let is_string = chunk.add_import("wasm:js-string", "test");
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_call(is_string, 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_call(is_number, 1, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if_value(line);
    // string / number → plain string form
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    crate::emitter::strings::emit_to_string(chunk, line);
    chunk.emit_else(line);
    // Error-like (own "message" AND own "stack") → Python `str(exc)` /
    // `print(exc)` is the message, not a struct dump. The flag avoids a false
    // positive on a plain dict that happens to have just one of those keys.
    let has_own = chunk.add_import("ecma:object", "hasOwn");
    let obj_get = chunk.add_import("ecma:object", "get");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let err_flag = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_string_const("message", line);
    chunk.emit_call(has_own, 2, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_string_const("stack", line);
    chunk.emit_call(has_own, 2, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op_u16(Op::LOCAL_SET, err_flag, line);
    chunk.emit_op_u16(Op::LOCAL_GET, err_flag, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_string_const("message", line);
    chunk.emit_call(obj_get, 2, line);
    chunk.emit_else(line);
    // dict → JSON stringify then Python-ify (`: ` sep, `, `, single quotes,
    // True/False/None).
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_call(json_str, 1, line);
    for (from, to) in [
        (":", ": "),
        (",", ", "),
        ("true", "True"),
        ("false", "False"),
        ("null", "None"),
        ("\"", "'"),
    ] {
        chunk.emit_string_const(from, line);
        chunk.emit_string_const(to, line);
        chunk.emit_call(replace_all, 3, line);
    }
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// Python `*` operator: array repeat, string repeat, or numeric multiply.
/// Stack: [a, b] → [result]
pub fn emit_pymul(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let is_array = chunk.add_import("ecma:array", "isArray");
    let str_repeat = chunk.add_import("ecma:string", "repeat");
    let test_str = chunk.add_import("wasm:js-string", "test");
    let b_slot = chunk.alloc_scratch(1);
    let a_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a_slot, line);

    // if isArray(a): array repeat via newWithLength(n).fill(arr).flat(1)
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    let new_arr = chunk.add_import("ecma:array", "newWithLength");
    chunk.emit_call(new_arr, 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    let fill = chunk.add_import("ecma:array", "fill");
    chunk.emit_call(fill, 2, line);
    chunk.emit_f64_const(1.0, line);
    let flat = chunk.add_import("ecma:array", "flat");
    chunk.emit_call(flat, 2, line);
    chunk.emit_else(line);

    // if string(a): string repeat
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(test_str, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_call(str_repeat, 2, line);
    chunk.emit_else(line);

    // numeric multiply
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// Python `.count(x)` — for arrays, count element occurrences.
/// Stack: [receiver, needle] → [count]
/// Uses ecma:array.filter + length to count matching elements.
pub fn emit_count(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let is_array = chunk.add_import("ecma:array", "isArray");
    let needle = chunk.alloc_scratch(1);
    let arr = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, needle, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr, line);

    // if isArray(arr): use filter to count matches
    chunk.emit_op_u16(Op::LOCAL_GET, arr, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_if_value(line);

    // arr.filter(e => e === needle).length
    // Use ecma:array.filter with a callback that compares to needle
    // For simplicity, use the indexOf count approach:
    // iterate and count via the runtime helper
    chunk.emit_op_u16(Op::LOCAL_GET, arr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, needle, line);
    let count_fn = chunk.add_import("ecma:array", "count");
    chunk.emit_call(count_fn, 2, line);

    chunk.emit_else(line);
    // string count: substring occurrences
    chunk.emit_op_u16(Op::LOCAL_GET, arr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, needle, line);
    let str_count = chunk.add_import("ecma:string", "count");
    chunk.emit_call(str_count, 2, line);
    chunk.emit_end(line);
}

/// Python `range(...)`.
///
/// The common one-argument form is emitted inline as a WASM loop. The
/// multi-argument forms still fall back to the shared runtime helper for
/// now because they need Python's nullable-argument reshaping semantics.
pub fn emit_range(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    collections::emit_range_targeted(chunks, current, argc, &Target::wasm(), line);
}

pub fn emit_helper(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    let global = match name {
        "python.hex" => "__vybe_pyhex",
        "python.oct" => "__vybe_pyoct",
        "python.bin" => "__vybe_pybin",
        "python.bytes" | "python.encode" => "__vybe_to_bytes",
        "python.enumerate" => "__vybe_enumerate",
        "python.zip" => "__vybe_zip",
        "python.map" => "__vybe_pymap",
        "python.filter" => "__vybe_pyfilter",
        "python.any" => "__vybe_pyany",
        "python.all" => "__vybe_pyall",
        "python.iter" => "__vybe_pyiter",
        "python.next" => "__vybe_pynext",
        "python.isinf" => "__vybe_isinf",
        "python.random_choice" => "__vybe_rand_choice",
        "python.random_shuffle" => "__vybe_rand_shuffle",
        "python.random_sample" => "__vybe_rand_sample",
        "python.instanceof" => "__vybe_instanceof",
        "python.callable" => "__vybe_callable",
        "python.id" => "__vybe_id",
        "python.hash" => "__vybe_hash",
        "python.regex_findall" => "__ecma_regexp_match_all_pat_first",
        "python.regex_sub" => "__ecma_regexp_replace_pat_first",
        "python.regex_split" => "__ecma_regexp_split_pat_first",
        "python.format_map" => "__vybe_format_map",
        "python.setdefault" => "__vybe_setdefault",
        "python.tostring" => "__vybe_tostring",
        _ => return false,
    };
    collections::emit_runtime_helper_call(chunks, current, global, argc, line);
    true
}
