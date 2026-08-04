//! Kotlin `toString` — the rendering `println`, `print` and string templates use.
//!
//! Kotlin does NOT render values the way a JS console does, and that is the
//! whole reason this exists. `wasi:logging/logging.log` renders its arguments
//! through `Value::Display` (ECMA console/inspect semantics: an array is
//! `1,2,3`, a Map is `[object]`), which is right for JS and wrong for every
//! Kotlin program:
//!
//! | value | console | Kotlin |
//! |---|---|---|
//! | `listOf(1, 2, 3)` | `1,2,3` | `[1, 2, 3]` |
//! | `mapOf("a" to 1)` | `[object]` | `{a=1}` |
//! | `Pair(3, "x")` | `3,x` | `(3, x)` |
//! | a `data class` | `[object]` | `Box(value=42)` |
//!
//! What this REPLACES is the more important part. Rendering used to be done by
//! `wrap_printable_arg` in the walker, which pattern-matched the ARGUMENT
//! EXPRESSION: `println(listOf(1,2,3))` was rewritten at parse time into string
//! concatenation, so it worked, while `val xs = listOf(1,2,3); println(xs)` —
//! the same value, one binding later — printed `1,2,3`. A renderer has to
//! dispatch on the VALUE, which means it belongs in an emitter, not in a
//! syntactic rewrite. Dart reaches the same conclusion in `emit_dart_print`.
//!
//! Composes the shared primitives (`collections`, `dict`, `strings`, `tuples`,
//! `loops`) rather than emitting raw opcodes wherever one exists.

use vybe_compiler::primitives::{dict, expressions, loops, ops, strings, tuples};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// The property that tells a Kotlin `Set` from a `Map`.
///
/// Both are `common:dict.new`'s shape — a struct with a `__keys` array — and a
/// `Set` is exactly a dict whose values are all `true`. Nothing in the VALUES
/// can distinguish them (`mapOf("a" to true)` is a legitimate Map), so the
/// walker stamps this and the renderer reads it. Written with the `__` prefix
/// every reserved key uses, so a Kotlin program cannot collide with it.
pub const SET_MARKER: &str = "__kt_set";

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    func: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, func);
    chunks[current].emit_call(idx, argc, line);
}

/// Push i32 `1` when the value in `slot` is an object (so `STRUCT_GET` on it
/// is safe — it traps on a primitive).
fn emit_is_object(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("object", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
}

/// Push i32 `1` when the object in `slot` is a Kotlin `Map`/`Set`.
///
/// A Kotlin dict is `common:dict.new`'s shape — a plain struct carrying a
/// `__keys` array for insertion order — NOT a runtime `ObjectKind::Map`. So the
/// ECMA `[object Map]` tag test that Python's dicts answer to says `false`
/// here, and the presence of `__keys` is the honest question to ask.
fn emit_has_dict_keys(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let key =
        chunks[current].add_constant(vybe_runtime::Value::String(std::sync::Arc::from("__keys")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
}

/// Kotlin `value.toString()`. Stack: `[value]` → `[string]`.
///
/// The order of the tests is the dispatch order Kotlin's `Any.toString` has:
/// a collection renders structurally, an object renders through its OWN
/// `toString` if it declares one, and everything else coerces.
pub fn emit_to_string(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_array_to_string(chunks, current, v, line);
    chunks[current].emit_else(line);

    emit_is_object(chunks, current, v, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_object_to_string(chunks, current, v, line);
    chunks[current].emit_else(line);
    // `null` → "null", `true` → "true", numbers, strings, chars.
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
}

/// Push i32 `1` when the object in `slot` fills [`ProtocolSlot::ToString`].
///
/// The same key `expressions::emit_rich_to_string` reads, asked as a question
/// so the answer can be used to ORDER the probes rather than only to call it.
fn emit_has_to_string_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let key = chunks[current].add_constant(vybe_runtime::Value::String(std::sync::Arc::from(
        vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::ToString).as_str(),
    )));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
}

/// Push i32 `1` when the object in `slot` carries [`SET_MARKER`].
fn emit_is_set(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let key = chunks[current].add_constant(vybe_runtime::Value::String(std::sync::Arc::from(
        SET_MARKER,
    )));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
}

/// An object: its own `toString`, else a `Set`, else a `Map`, else the ECMA
/// coercion.
///
/// **The slot is probed FIRST.** An object that fills `ProtocolSlot::ToString`
/// has said how it renders, and no structural guess may override that. The
/// Set/Map probes are inferences from SHAPE — `emit_has_dict_keys` asks whether
/// the object carries `__keys`, which an ordinary class instance can also carry
/// — so running them first let a declared `toString` be shadowed by a
/// coincidence. Measured on a Kotlin `enum class`: `println(Color.RED)` printed
/// `{name=RED, ordinal=0}`, the map rendering of the constant's own fields,
/// instead of `RED`.
fn emit_object_to_string(chunks: &mut Vec<Chunk>, current: usize, v: u16, line: u32) {
    emit_has_to_string_slot(chunks, current, v, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    expressions::emit_rich_to_string(&mut chunks[current], v, line);
    chunks[current].emit_else(line);

    emit_is_set(chunks, current, v, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_set_to_string(chunks, current, v, line);
    chunks[current].emit_else(line);

    emit_has_dict_keys(chunks, current, v, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_map_to_string(chunks, current, v, line);
    chunks[current].emit_else(line);

    // A class's own `toString` is reached by its SLOT, never by spelling —
    // `flexclassplan.md`'s bind-don't-name rule. A `data class`'s synthesised
    // `toString`, a user `override fun toString()` and a Java-style
    // `toString()` all fill `ProtocolSlot::ToString`, so one probe reaches
    // every one of them, and a user member that merely happens to be CALLED
    // `toString` cannot be captured by it. That probe, with its
    // `ecma:string.String` fallback for an object that fills no slot, IS
    // `expressions::emit_rich_to_string` — shared, so Kotlin does not carry a
    // second copy of the calling convention that would drift from it.
    expressions::emit_rich_to_string(&mut chunks[current], v, line);

    // One `end` per `if_value` above — slot, set, map. Getting this count
    // wrong does not fail to compile: it silently rebalances the surrounding
    // blocks, and every later value renders as its NEIGHBOUR.
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `[1, 2, 3]` — a Kotlin `Set` renders like a list: its KEYS are its elements.
fn emit_set_to_string(chunks: &mut Vec<Chunk>, current: usize, v: u16, line: u32) {
    let keys = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    dict::emit_keys(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);

    chunks[current].emit_string_const("[", line);
    emit_join_rendered(chunks, current, keys, ", ", Some(SET_MARKER), line);
    chunks[current].emit_string_const("]", line);
    strings::emit_concat(&mut chunks[current], 3, line);
}

/// `[1, 2, 3]` for a List/Array, `(3, x)` for a Pair/Triple.
///
/// The two differ only in their brackets, and only a TAG can tell them apart:
/// `Pair(3, "x")` and `listOf(3, "x")` are the same runtime array otherwise.
/// That is what `tuple_literals_tagged` in the profile buys.
fn emit_array_to_string(chunks: &mut Vec<Chunk>, current: usize, v: u16, line: u32) {
    let is_tuple = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    tuples::emit_is_tuple(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, is_tuple, line);

    let joined = chunks[current].alloc_scratch(1);
    emit_join_rendered(chunks, current, v, ", ", None, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, joined, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, is_tuple, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("(", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, joined, line);
    chunks[current].emit_string_const(")", line);
    strings::emit_concat(&mut chunks[current], 3, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("[", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, joined, line);
    chunks[current].emit_string_const("]", line);
    strings::emit_concat(&mut chunks[current], 3, line);
    chunks[current].emit_end(line);
}

/// Join the array in `arr_slot` with `sep`, rendering each ELEMENT through the
/// class `ToString` slot when present.
///
/// `ecma:array.join` would be shorter, but it stringifies elements the host's
/// way — a nested list inside a list would come back as `1,2` instead of
/// `[1, 2]`, which is the same class of bug this module exists to fix.
fn emit_join_rendered(
    chunks: &mut Vec<Chunk>,
    current: usize,
    arr_slot: u16,
    sep: &str,
    skip: Option<&str>,
    line: u32,
) {
    let sep_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_string_const(sep, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sep_slot, line);
    emit_join_rendered_local(chunks, current, arr_slot, sep_slot, skip, line);
}

fn emit_join_rendered_local(
    chunks: &mut Vec<Chunk>,
    current: usize,
    arr_slot: u16,
    sep_slot: u16,
    skip: Option<&str>,
    line: u32,
) {
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let text = chunks[current].alloc_scratch(1);

    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx, line);
    // Stack: [element]
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem, line);
    expressions::emit_rich_to_string(&mut chunks[current], elem, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);

    // The separator is decided by whether anything has been appended YET, not
    // by the index — an index rule would emit a leading `, ` the moment an
    // element is skipped.
    if let Some(marker) = skip {
        chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
        chunks[current].emit_string_const(marker, line);
        call_import(chunks, current, "wasm:js-string", "equals", 2, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_else(line);
        emit_separated_text(chunks, current, out, text, sep_slot, line);
        chunks[current].emit_end(line);
    } else {
        emit_separated_text(chunks, current, out, text, sep_slot, line);
    }

    strings::emit_concat(&mut chunks[current], 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    loops::emit_for_in_end(chunks, current, idx, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `(out == "" ? "" : sep) + text` — the piece appended for one element.
fn emit_separated_text(
    chunks: &mut Vec<Chunk>,
    current: usize,
    out: u16,
    text: u16,
    sep_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_string_const("", line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
    strings::emit_concat(&mut chunks[current], 2, line);
}

/// Kotlin `joinToString(separator)` — render elements through the same class
/// `ToString` slot path as `println`.
pub fn emit_join_to_string(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 2 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let sep = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sep, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr, line);
    emit_join_rendered_local(chunks, current, arr, sep, None, line);
}

/// `{a=1, b=2}` — Kotlin's map rendering. `=` between key and value, `, `
/// between entries, braces around the whole, insertion order preserved
/// (which is why the keys come from `dict::emit_keys` rather than the host).
fn emit_map_to_string(chunks: &mut Vec<Chunk>, current: usize, v: u16, line: u32) {
    let keys = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let text_key = chunks[current].alloc_scratch(1);
    let text_value = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    dict::emit_keys(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);

    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    let state = loops::emit_for_in_start(chunks, current, keys, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    dict::emit_get_dynamic(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    expressions::emit_rich_to_string(&mut chunks[current], key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_key, line);
    expressions::emit_rich_to_string(&mut chunks[current], value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_value, line);

    // out += (i == 0 ? "" : ", ") + key.toString() + "=" + value.toString()
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const(", ", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_key, line);
    chunks[current].emit_string_const("=", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_value, line);
    strings::emit_concat(&mut chunks[current], 5, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    loops::emit_for_in_end(chunks, current, idx, state, line);

    chunks[current].emit_string_const("{", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_string_const("}", line);
    strings::emit_concat(&mut chunks[current], 3, line);
}

/// Kotlin `println(value)` / `print(value)` — render, then log.
///
/// Stack: `[value]` → `[]`.
pub fn emit_print(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    // Bare `println()` prints an empty line — there is no argument to render,
    // and reaching `emit_to_string` with an empty stack would consume whatever
    // the caller left there.
    if argc == 0 {
        chunks[current].emit_string_const("", line);
    } else {
        emit_to_string(chunks, current, line);
    }
    let log = chunks[current].add_import("wasi:logging/logging", "log");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, log, line);
    chunks[current].emit(1, line);
}
