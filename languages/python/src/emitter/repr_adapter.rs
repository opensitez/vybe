//! Python recursive `repr` — a self-recursive bytecode chunk.
//!
//! Renders a value to its Python `repr` form, recursing into containers so
//! NESTED tuples/lists/dicts display correctly (`[(1, 2), (3, 4)]`, not
//! `[[1, 2], [3, 4]]`). This replaces the old `emit_py_repr` array path, which
//! went through `ecma:json.stringify` — JSON flattens tuples to `[...]` and
//! cannot tell a tuple from a list.
//!
//! Pattern mirrors PHP's `build_php_json_normalize_helper`
//! (`languages/php/src/emitter/misc_adapter.rs`): a `create_function_chunk`
//! that recurses on itself via `REF_FUNC` + `CALL_REF`. NO addition to the
//! retiring `__vybe_*`/`__stdlib_*` bundle, and NO Python semantics in the
//! host — Python's display stays in the Python crate.
//!
//! The chunk's name is `__py_repr`; it is built once per module (deduped by
//! scanning `chunks` for that name), since `repr` runs on every `print`/`str`.

use std::sync::Arc;
use vybe_compiler::primitives::functions::create_function_chunk;
use vybe_compiler::primitives::tuples::{FIELDS_TAG, TUPLE_TAG, TYPENAME_TAG};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

const REPR_CHUNK: &str = "__py_repr";

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}
fn str_const(chunk: &mut Chunk, s: &str, line: u32) {
    chunk.emit_string_const(s, line);
}
/// String concatenation of the top two stack strings. `[a, b] -> [a+b]`.
fn concat(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "concat");
    chunk.emit_call(idx, 2, line);
}
/// Recurse: `[value] -> [repr_string]` by calling this same chunk.
fn recurse(chunk: &mut Chunk, self_idx: usize, line: u32) {
    // ref to self, then the value is already on the stack ABOVE where we want
    // the fn — so callers arrange `[fn, value]`. Here we assume value on top:
    // stash it, push fn, push value, call.
    let tmp = chunk.alloc_scratch(1);
    lset(chunk, tmp, line);
    chunk.emit_op_u16(Op::REF_FUNC, self_idx as u16, line);
    chunk.emit(0, line);
    lget(chunk, tmp, line);
    chunk.emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
}

/// Return the index of the `__py_repr` chunk, building it once per module.
pub fn ensure_py_repr_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    if let Some(idx) = chunks.iter().position(|c| c.name == REPR_CHUNK) {
        return idx;
    }
    build_py_repr_chunk(chunks, line)
}

fn build_py_repr_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let self_idx = chunks.len();
    let mut c = create_function_chunk(REPR_CHUNK, 1);
    c.alloc_scratch(1); // reserve arg slot 0

    let value = 0u16;
    let out = c.alloc_scratch(1);
    let i = c.alloc_scratch(1);
    let n = c.alloc_scratch(1);
    let keys = c.alloc_scratch(1);
    let key = c.alloc_scratch(1);
    let fields = c.alloc_scratch(1);
    let esc = c.alloc_scratch(1);

    // ── None ────────────────────────────────────────────────────────────
    lget(&mut c, value, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if(line);
    str_const(&mut c, "None", line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);

    // ── bool → True / False ─────────────────────────────────────────────
    lget(&mut c, value, line);
    {
        let idx = c.add_import("wasm:js-boolean", "test");
        c.emit_call(idx, 1, line);
    }
    c.emit_if(line);
    lget(&mut c, value, line);
    {
        let idx = c.add_import("wasm:js-boolean", "cast");
        c.emit_call(idx, 1, line);
    }
    c.emit_if_value(line);
    str_const(&mut c, "True", line);
    c.emit_else(line);
    str_const(&mut c, "False", line);
    c.emit_end(line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);

    // ── string → 'quoted' (repr form: control chars escaped, Python quoting) ──
    lget(&mut c, value, line);
    {
        let idx = c.add_import("wasm:js-string", "test");
        c.emit_call(idx, 1, line);
    }
    c.emit_if(line);
    {
        let replace_all = c.add_import("ecma:string", "replaceAll");
        let includes = c.add_import("ecma:string", "includes");

        // esc = value with `\` escaped first, then the control chars — Python
        // repr renders `\n`/`\r`/`\t`/`\\` literally, not as raw bytes.
        lget(&mut c, value, line);
        for (search, repl) in [("\\", "\\\\"), ("\n", "\\n"), ("\r", "\\r"), ("\t", "\\t")] {
            str_const(&mut c, search, line);
            str_const(&mut c, repl, line);
            c.emit_call(replace_all, 3, line);
        }
        lset(&mut c, esc, line);

        // Python quote choice: `'…'` normally, but `"…"` when the string holds a
        // single quote and no double quote (so no `\'` escaping is needed).
        lget(&mut c, esc, line);
        str_const(&mut c, "'", line);
        c.emit_call(includes, 2, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut c, line);
        lget(&mut c, esc, line);
        str_const(&mut c, "\"", line);
        c.emit_call(includes, 2, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut c, line);
        c.emit_op(Op::I32_EQZ, line); // no double quote
        c.emit_op(Op::I32_AND, line); // has ' AND no "
        c.emit_if_value(line);
        // "…" — double-quoted, single quotes left bare
        str_const(&mut c, "\"", line);
        lget(&mut c, esc, line);
        concat(&mut c, line);
        str_const(&mut c, "\"", line);
        concat(&mut c, line);
        c.emit_else(line);
        // '…' — single-quoted, escape any embedded single quote
        str_const(&mut c, "'", line);
        lget(&mut c, esc, line);
        str_const(&mut c, "'", line);
        str_const(&mut c, "\\'", line);
        c.emit_call(replace_all, 3, line);
        concat(&mut c, line);
        str_const(&mut c, "'", line);
        concat(&mut c, line);
        c.emit_end(line);
        c.emit_op(Op::RETURN, line);
    }
    c.emit_end(line);

    // ── number → its string form ────────────────────────────────────────
    lget(&mut c, value, line);
    {
        let idx = c.add_import("wasm:js-number", "test");
        c.emit_call(idx, 1, line);
    }
    c.emit_if(line);
    lget(&mut c, value, line);
    {
        let idx = c.add_import("ecma:number", "isNaN");
        c.emit_call(idx, 1, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut c, line);
    c.emit_if(line);
    str_const(&mut c, "nan", line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);

    lget(&mut c, value, line);
    {
        let idx = c.add_import("ecma:number", "isFinite");
        c.emit_call(idx, 1, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut c, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    lget(&mut c, value, line);
    {
        let idx = c.add_import("wasm:js-number", "toF64");
        c.emit_call(idx, 1, line);
    }
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_if(line);
    str_const(&mut c, "-inf", line);
    c.emit_else(line);
    str_const(&mut c, "inf", line);
    c.emit_end(line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);

    lget(&mut c, value, line);
    {
        let idx = c.add_import("ecma:string", "String");
        c.emit_call(idx, 1, line);
    }
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);

    // ── array → named tuple / tuple / list ──────────────────────────────
    lget(&mut c, value, line);
    {
        let idx = c.add_import("ecma:array", "isArray");
        c.emit_call(idx, 1, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut c, line);
    c.emit_if(line);
    {
        // n = value.length; i = 0
        lget(&mut c, value, line);
        c.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut c, n, line);

        // Named tuple? `__typename` present → `Name(f=v, …)`.
        lget(&mut c, value, line);
        struct_get(&mut c, TYPENAME_TAG, line);
        c.emit_op(Op::REF_IS_NULL, line);
        c.emit_op(Op::I32_EQZ, line);
        c.emit_if(line);
        {
            lget(&mut c, value, line);
            struct_get(&mut c, TYPENAME_TAG, line);
            str_const(&mut c, "(", line);
            concat(&mut c, line);
            lset(&mut c, out, line);
            lget(&mut c, value, line);
            struct_get(&mut c, FIELDS_TAG, line);
            lset(&mut c, fields, line);
            emit_i32_zero(&mut c, i, line);
            let lp = loop_start(&mut c, line);
            loop_break_if_ge(&mut c, i, n, line);
            // out += (i>0 ? ", " : "") + fields[i] + "=" + repr(value[i])
            emit_sep_comma(&mut c, out, i, line);
            lget(&mut c, out, line);
            lget(&mut c, fields, line);
            lget(&mut c, i, line);
            c.emit_op(Op::ARRAY_GET, line);
            concat(&mut c, line);
            str_const(&mut c, "=", line);
            concat(&mut c, line);
            lget(&mut c, value, line);
            lget(&mut c, i, line);
            c.emit_op(Op::ARRAY_GET, line);
            recurse(&mut c, self_idx, line);
            concat(&mut c, line);
            lset(&mut c, out, line);
            bump(&mut c, i, line);
            loop_end(&mut c, lp, line);
            lget(&mut c, out, line);
            str_const(&mut c, ")", line);
            concat(&mut c, line);
            c.emit_op(Op::RETURN, line);
        }
        c.emit_end(line);

        // Plain tuple? `__tuple` present → `(a, b)` / `(a,)`.
        lget(&mut c, value, line);
        struct_get(&mut c, TUPLE_TAG, line);
        c.emit_op(Op::REF_IS_NULL, line);
        c.emit_op(Op::I32_EQZ, line);
        c.emit_if(line);
        {
            emit_join(&mut c, self_idx, value, out, i, n, "(", ")", line);
            // 1-tuple gets a trailing comma: `(x,)`. `out` currently is `(x)`;
            // rebuild with the comma when n == 1.
            lget(&mut c, n, line);
            emit_i32_const(&mut c, 1, line);
            c.emit_op(Op::I32_EQ, line);
            c.emit_if(line);
            str_const(&mut c, "(", line);
            lget(&mut c, value, line);
            emit_i32_zero_val(&mut c, line);
            c.emit_op(Op::ARRAY_GET, line);
            recurse(&mut c, self_idx, line);
            concat(&mut c, line);
            str_const(&mut c, ",)", line);
            concat(&mut c, line);
            lset(&mut c, out, line);
            c.emit_end(line);
            lget(&mut c, out, line);
            c.emit_op(Op::RETURN, line);
        }
        c.emit_end(line);

        // Plain list → `[a, b]`.
        emit_join(&mut c, self_idx, value, out, i, n, "[", "]", line);
        lget(&mut c, out, line);
        c.emit_op(Op::RETURN, line);
    }
    c.emit_end(line);

    // ── range → `range(0, 3)` / `range(1, 10, 2)` ───────────────────────
    // A range is lazy and opaque, so `emit_range` stamps its bounds onto the
    // object for exactly this. CPython omits the step when it is 1.
    lget(&mut c, value, line);
    struct_get(&mut c, "__py_range_stop", line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    {
        let concat = c.add_import("wasm:js-string", "concat");
        let to_str = c.add_import("ecma:string", "String");

        str_const(&mut c, "range(", line);
        lget(&mut c, value, line);
        struct_get(&mut c, "__py_range_start", line);
        c.emit_call(to_str, 1, line);
        c.emit_call(concat, 2, line);
        str_const(&mut c, ", ", line);
        c.emit_call(concat, 2, line);
        lget(&mut c, value, line);
        struct_get(&mut c, "__py_range_stop", line);
        c.emit_call(to_str, 1, line);
        c.emit_call(concat, 2, line);
        lset(&mut c, out, line);

        // step != 1 → append ", <step>"
        lget(&mut c, value, line);
        struct_get(&mut c, "__py_range_step", line);
        {
            let to_f64 = c.add_import("wasm:js-number", "toF64");
            c.emit_call(to_f64, 1, line);
        }
        c.emit_f64_const(1.0, line);
        c.emit_op(Op::F64_NE, line);
        c.emit_if(line);
        lget(&mut c, out, line);
        str_const(&mut c, ", ", line);
        c.emit_call(concat, 2, line);
        lget(&mut c, value, line);
        struct_get(&mut c, "__py_range_step", line);
        c.emit_call(to_str, 1, line);
        c.emit_call(concat, 2, line);
        lset(&mut c, out, line);
        c.emit_end(line);

        lget(&mut c, out, line);
        str_const(&mut c, ")", line);
        c.emit_call(concat, 2, line);
        c.emit_op(Op::RETURN, line);
    }
    c.emit_end(line);

    // ── user object with __repr__ / __str__ ─────────────────────────────
    // Class methods live on the PROTOTYPE, not as own properties, so look them
    // up prototype-aware via `__vybe_js_get_method` (the same global the
    // str-form dispatch uses) rather than `struct_get`, which only sees own
    // properties and would miss a class-defined `__repr__`.
    let test_undef = c.add_import("wasm:js-undefined", "test");
    for method in ["__repr__", "__str__"] {
        let m = c.alloc_scratch(1);
        vybe_compiler::primitives::globals::emit_read(&mut c, "__vybe_js_get_method", line);
        lget(&mut c, value, line);
        str_const(&mut c, method, line);
        c.emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
        lset(&mut c, m, line);
        lget(&mut c, m, line);
        c.emit_call(test_undef, 1, line); // 1 if undefined
        c.emit_op(Op::I32_EQZ, line); // 1 if a method was found
        c.emit_if(line);
        lget(&mut c, m, line);
        lget(&mut c, value, line);
        c.emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
        c.emit_op(Op::RETURN, line);
        c.emit_end(line);
    }

    // ── set / frozenset → `{a, b, …}` (empty → `set()`) ─────────────────
    // Sets carry `__type == "Set"` (stamped by `ecma:set`); check this BEFORE
    // the generic class-instance branch, which would otherwise render the set
    // as `<Set object at 0x0>`. `ecma:array.from` materializes the members in
    // insertion order, then the list machinery renders them inside braces.
    {
        lget(&mut c, value, line);
        struct_get(&mut c, "__type", line);
        str_const(&mut c, "Set", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut c, line);
        c.emit_if(line);
        {
            let set_arr = c.alloc_scratch(1);
            // frozen = truthy(value.__frozenset) — `frozenset(...)` reprs as
            // `frozenset({...})` / `frozenset()`; a plain set as `{...}` / `set()`.
            let frozen = c.alloc_scratch(1);
            lget(&mut c, value, line);
            struct_get(&mut c, "__frozenset", line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut c, line);
            lset(&mut c, frozen, line);

            lget(&mut c, value, line);
            let from = c.add_import("ecma:array", "from");
            c.emit_call(from, 1, line);
            lset(&mut c, set_arr, line);
            lget(&mut c, set_arr, line);
            c.emit_op(Op::ARRAY_LENGTH, line);
            lset(&mut c, n, line);
            // empty set → `set()` / `frozenset()` (Python has no `{}` empty-set)
            lget(&mut c, n, line);
            emit_i32_const(&mut c, 0, line);
            c.emit_op(Op::I32_EQ, line);
            c.emit_if(line);
            lget(&mut c, frozen, line);
            c.emit_if_value(line);
            str_const(&mut c, "frozenset()", line);
            c.emit_else(line);
            str_const(&mut c, "set()", line);
            c.emit_end(line);
            c.emit_op(Op::RETURN, line);
            c.emit_end(line);
            // non-empty: brace-join, wrapping in `frozenset(...)` when frozen
            lget(&mut c, frozen, line);
            c.emit_if(line);
            emit_join(
                &mut c,
                self_idx,
                set_arr,
                out,
                i,
                n,
                "frozenset({",
                "})",
                line,
            );
            c.emit_else(line);
            emit_join(&mut c, self_idx, set_arr, out, i, n, "{", "}", line);
            c.emit_end(line);
            lget(&mut c, out, line);
            c.emit_op(Op::RETURN, line);
        }
        c.emit_end(line);
    }

    // ── class instance without __repr__ → `<ClassName object at 0x0>` ────
    // A class instance carries a `__type` stamp (a plain dict does not); without
    // this it falls to the dict branch and reprs as `{}`. Known gap: a
    // `dict`/`list` subclass also has `__type` and renders here rather than as
    // its container — indistinguishable structurally in this value model.
    {
        let has_own = c.add_import("ecma:object", "hasOwn");
        lget(&mut c, value, line);
        str_const(&mut c, "__type", line);
        c.emit_call(has_own, 2, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut c, line);
        c.emit_if(line);
        str_const(&mut c, "<", line);
        lget(&mut c, value, line);
        struct_get(&mut c, "__type", line);
        {
            let to_str = c.add_import("ecma:string", "String");
            c.emit_call(to_str, 1, line);
        }
        concat(&mut c, line);
        str_const(&mut c, " object at 0x0>", line);
        concat(&mut c, line);
        c.emit_op(Op::RETURN, line);
        c.emit_end(line);
    }

    // ── dict → {k: v, …} ─────────────────────────────────────────────────
    // Iterate ENTRIES ([k, v] pairs), not keys-then-index: a Map's key does not
    // round-trip through `d[key]` (int keys), and `entries` works for both
    // Map-backed dicts and Ordinary objects.
    lget(&mut c, value, line);
    {
        let idx = c.add_import("ecma:object", "entries");
        c.emit_call(idx, 1, line);
    }
    lset(&mut c, keys, line);
    lget(&mut c, keys, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    lset(&mut c, n, line);
    str_const(&mut c, "{", line);
    lset(&mut c, out, line);
    emit_i32_zero(&mut c, i, line);
    let lp = loop_start(&mut c, line);
    loop_break_if_ge(&mut c, i, n, line);
    emit_sep_comma(&mut c, out, i, line);
    // pair = entries[i]; out += repr(pair[0]) + ": " + repr(pair[1])
    lget(&mut c, keys, line);
    lget(&mut c, i, line);
    c.emit_op(Op::ARRAY_GET, line);
    lset(&mut c, key, line);
    lget(&mut c, out, line);
    lget(&mut c, key, line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::ARRAY_GET, line);
    recurse(&mut c, self_idx, line);
    concat(&mut c, line);
    str_const(&mut c, ": ", line);
    concat(&mut c, line);
    lget(&mut c, key, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::ARRAY_GET, line);
    recurse(&mut c, self_idx, line);
    concat(&mut c, line);
    lset(&mut c, out, line);
    bump(&mut c, i, line);
    loop_end(&mut c, lp, line);
    lget(&mut c, out, line);
    str_const(&mut c, "}", line);
    concat(&mut c, line);
    c.emit_op(Op::RETURN, line);

    chunks.push(c);
    self_idx
}

// ── small emit helpers ──────────────────────────────────────────────────

fn struct_get(chunk: &mut Chunk, key: &str, line: u32) {
    let k = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
}

fn emit_i32_const(chunk: &mut Chunk, v: i32, line: u32) {
    vybe_compiler::primitives::instructions::core_wasm::i32_const(chunk, line, v);
}
fn emit_i32_zero_val(chunk: &mut Chunk, line: u32) {
    emit_i32_const(chunk, 0, line);
}
fn emit_i32_zero(chunk: &mut Chunk, slot: u16, line: u32) {
    emit_i32_const(chunk, 0, line);
    lset(chunk, slot, line);
}
fn bump(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    emit_i32_const(chunk, 1, line);
    chunk.emit_op(Op::I32_ADD, line);
    lset(chunk, slot, line);
}

fn loop_start(chunk: &mut Chunk, line: u32) -> vybe_compiler::primitives::loops::LoopState {
    let block_patch = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    vybe_compiler::primitives::loops::LoopState {
        block_patch,
        loop_patch,
        body_block_patch: None,
    }
}
fn loop_end(chunk: &mut Chunk, state: vybe_compiler::primitives::loops::LoopState, line: u32) {
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(state.loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(state.block_patch);
}
/// Break out of the loop when `!(i < n)`.
fn loop_break_if_ge(chunk: &mut Chunk, i: u16, n: u16, line: u32) {
    lget(chunk, i, line);
    lget(chunk, n, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);
}
/// `out += ", "` when `i > 0`.
fn emit_sep_comma(chunk: &mut Chunk, out: u16, i: u16, line: u32) {
    lget(chunk, i, line);
    emit_i32_const(chunk, 0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);
    lget(chunk, out, line);
    str_const(chunk, ", ", line);
    concat(chunk, line);
    lset(chunk, out, line);
    chunk.emit_end(line);
}

/// `out = prefix + join(", ", repr(value[0..n])) + suffix`, stored in `out`.
#[allow(clippy::too_many_arguments)]
fn emit_join(
    chunk: &mut Chunk,
    self_idx: usize,
    value: u16,
    out: u16,
    i: u16,
    n: u16,
    prefix: &str,
    suffix: &str,
    line: u32,
) {
    str_const(chunk, prefix, line);
    lset(chunk, out, line);
    emit_i32_zero(chunk, i, line);
    let lp = loop_start(chunk, line);
    loop_break_if_ge(chunk, i, n, line);
    emit_sep_comma(chunk, out, i, line);
    lget(chunk, out, line);
    lget(chunk, value, line);
    lget(chunk, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    recurse(chunk, self_idx, line);
    concat(chunk, line);
    lset(chunk, out, line);
    bump(chunk, i, line);
    loop_end(chunk, lp, line);
    lget(chunk, out, line);
    str_const(chunk, suffix, line);
    concat(chunk, line);
    lset(chunk, out, line);
}
