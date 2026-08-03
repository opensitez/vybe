//! `dataclasses` module functions — `is_dataclass`, `fields`, `asdict`,
//! `astuple`.
//!
//! These are pure, stateless reads over one piece of data the walker already
//! synthesizes: `__dataclass_fields__`, a class-level array of the field names
//! in declaration order (`walker::synthesize_dataclass_members`). Everything
//! here is a loop over that array, so there is no module state and no prelude
//! — the adapter shape (`math_adapter`, `string_adapter`) applies.
//!
//! Declaration order matters: reading the instance property bag instead would
//! make `asdict`/`astuple` non-deterministic, since the object's own
//! properties are not insertion-ordered.
//!
//! `dataclasses.replace` is NOT here: it takes keyword arguments, which do not
//! survive to the emit boundary (`emit_common` receives a value stack and a
//! count, never an argument name), so it has to be lowered in the walker.

use vybe_runtime::{Chunk, Op};

use vybe_compiler::primitives::{reflection, tuples};

/// Marker the walker stamps on every `@dataclass`; `is_dataclass` tests for it
/// and the rest walk it.
const FIELDS_KEY: &str = "__dataclass_fields__";

/// Leave the field-name array for the value in `obj` in `out`, or null.
///
/// Accepts BOTH a class and an instance, as CPython's `is_dataclass`/`fields`
/// do: the array is a class attribute, so an instance reaches it through its
/// `__class__` link while the class carries it directly.
fn emit_field_list(chunk: &mut Chunk, obj: u16, out: u16, line: u32) {
    // out = obj.__dataclass_fields__  (hit when `obj` IS the class)
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_string_const(FIELDS_KEY, line);
    reflection::emit_get_property_in_chunk(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);

    // still null → follow the instance's `__class__` link and retry
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    {
        let cls = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
        chunk.emit_string_const("__class__", line);
        reflection::emit_get_property_in_chunk(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_SET, cls, line);

        chunk.emit_op_u16(Op::LOCAL_GET, cls, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, cls, line);
        chunk.emit_string_const(FIELDS_KEY, line);
        reflection::emit_get_property_in_chunk(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_SET, out, line);
        chunk.emit_end(line);
    }
    chunk.emit_end(line);
}

/// `is_dataclass(x)` — true for a dataclass CLASS and for its instances.
/// Stack: `[x] -> [bool]`.
pub fn emit_is_dataclass(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        vybe_compiler::primitives::instructions::core_wasm::bool_const(chunk, line, false);
        return;
    }
    let obj = chunk.alloc_scratch(1);
    let list = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);
    emit_field_list(chunk, obj, list, line);

    chunk.emit_op_u16(Op::LOCAL_GET, list, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    vybe_compiler::primitives::instructions::core_wasm::bool_const(chunk, line, false);
    chunk.emit_else(line);
    vybe_compiler::primitives::instructions::core_wasm::bool_const(chunk, line, true);
    chunk.emit_end(line);
}

/// How the per-field loop accumulates.
enum Shape {
    /// `asdict` — a Python dict, i.e. a `Map` of name → value.
    Dict,
    /// `astuple` — a tagged tuple of the values.
    Tuple,
    /// `fields` — a list of field objects, each carrying `.name`.
    Descriptors,
}

/// Walk `__dataclass_fields__`, reading each attribute off the instance.
/// Stack: `[obj] -> [result]`. A non-dataclass yields the empty result rather
/// than trapping, which keeps a mis-typed argument diagnosable.
fn emit_walk(chunks: &mut [Chunk], current: usize, argc: u8, shape: Shape, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    let obj = chunk.alloc_scratch(1);
    let list = chunk.alloc_scratch(1);
    let acc = chunk.alloc_scratch(1);
    let i = chunk.alloc_scratch(1);
    let n = chunk.alloc_scratch(1);
    let key = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);
    emit_field_list(chunk, obj, list, line);

    // A missing marker means "not a dataclass" — iterate zero times.
    chunk.emit_op_u16(Op::LOCAL_GET, list, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, list, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, n, line);

    // acc = {} / []
    match shape {
        Shape::Dict => {
            let new_map = chunk.add_import("ecma:map", "new");
            chunk.emit_call(new_map, 0, line);
        }
        Shape::Tuple | Shape::Descriptors => {
            let new_len = chunk.add_import("vybe:js-array", "newWithLength");
            chunk.emit_i32_const(0, line);
            chunk.emit_call(new_len, 1, line);
        }
    }
    chunk.emit_op_u16(Op::LOCAL_SET, acc, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    let block = chunk.emit_block(line);
    let (lp, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, list, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, key, line);

    match shape {
        Shape::Dict => {
            // acc.set(key, obj[key])
            chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
            chunk.emit_op_u16(Op::LOCAL_GET, key, line);
            chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
            chunk.emit_op_u16(Op::LOCAL_GET, key, line);
            reflection::emit_get_property_in_chunk(chunk, line);
            let set = chunk.add_import("ecma:map", "set");
            chunk.emit_call(set, 3, line);
            chunk.emit_op(Op::DROP, line);
        }
        Shape::Tuple => {
            chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
            chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
            chunk.emit_op_u16(Op::LOCAL_GET, key, line);
            reflection::emit_get_property_in_chunk(chunk, line);
            let push = chunk.add_import("ecma:array", "push");
            chunk.emit_call(push, 2, line);
            chunk.emit_op(Op::DROP, line);
        }
        Shape::Descriptors => {
            // `Field` stand-in: the one attribute `fields()` callers read is
            // `.name` (plus `.default`, which the walker does not preserve).
            chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
            chunk.emit_struct_new(0, 0, line);
            chunk.emit_dup(line);
            chunk.emit_op_u16(Op::LOCAL_GET, key, line);
            let name_key = chunk.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(
                "name",
            )));
            chunk.emit_struct_field_op(Op::STRUCT_SET, 0, name_key, line);
            chunk.emit_op(Op::DROP, line);
            let push = chunk.add_import("ecma:array", "push");
            chunk.emit_call(push, 2, line);
            chunk.emit_op(Op::DROP, line);
        }
    }

    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(lp);
    chunk.emit_end(line);
    chunk.patch_block(block);

    chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
    if matches!(shape, Shape::Tuple) {
        tuples::emit_tag(chunks, current, line);
    }
}

/// `asdict(obj)` → `{'x': 1}`. Stack: `[obj] -> [dict]`.
pub fn emit_asdict(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_walk(chunks, current, argc, Shape::Dict, line);
}

/// `astuple(obj)` → `(1, 2)`. Stack: `[obj] -> [tuple]`.
pub fn emit_astuple(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_walk(chunks, current, argc, Shape::Tuple, line);
}

/// `fields(obj)` → a list of field descriptors carrying `.name`.
/// Stack: `[obj] -> [array]`.
pub fn emit_fields(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_walk(chunks, current, argc, Shape::Descriptors, line);
}
