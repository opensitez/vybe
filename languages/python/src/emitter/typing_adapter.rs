//! Python `typing` / `types` runtime surface.
//!
//! Most of `typing` is erased before it reaches the compiler — an annotation
//! is descriptive (PEP 484), and `[compiler] coerces_value_to_type_hint =
//! false` says so. What is left is the handful of names that are ordinary
//! CALLS at run time, and those are what this file emits:
//!
//! * `cast(t, v)` → `v`. CPython does not check anything; it returns the
//!   value unchanged, and the type argument exists for the checker.
//! * `final(x)`, `no_type_check(x)`, `runtime_checkable(x)` → `x`, with the
//!   marker attribute CPython sets on the decorated object.
//! * `TypeVar` / `ParamSpec` / `TypeVarTuple` → a named marker object.
//! * `NewType(name, tp)` → a CALLABLE identity, so `UserId(3)` is `3`.
//! * `get_type_hints(x)` → `x.__annotations__`, or an empty dict.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use vybe_compiler::primitives::{collections, dict};

use super::adapter_util::{lget, lset, new_tagged, set_call_slot, stash_exact, struct_set};
use vybe_compiler::primitives::class_slots::ClassSlot;

/// `typing.cast(typ, val)` → `val`.
pub fn emit_cast(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 2, line);
    lget(&mut chunks[current], base + 1, line);
}

pub fn emit_identity(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    lget(&mut chunks[current], base, line);
}

/// `typing.final(x)` / `no_type_check(x)` / `runtime_checkable(x)` — returns
/// the argument, with `attr` set on it the way CPython's decorator does.
pub fn emit_marker_decorator(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    attr: &str,
    line: u32,
) {
    let base = stash_exact(chunks, current, argc, 1, line);
    let chunk = &mut chunks[current];
    // Only an object can carry the marker; a primitive is returned untouched.
    lget(chunk, base, line);
    let typeof_idx = chunk.add_import("ecma:value", "typeof");
    chunk.emit_call(typeof_idx, 1, line);
    chunk.emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, base, line);
    chunk.emit_bool_const(true, line);
    struct_set(chunk, &ClassSlot::internal(attr), line);
    chunk.emit_end(line);
    lget(chunk, base, line);
}

/// `TypeVar('T')`, `ParamSpec('P')`, `TypeVarTuple('Ts')` — CPython builds an
/// object whose identity is its name; nothing at run time inspects more than
/// that, because annotations are never evaluated.
pub fn emit_type_marker(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    type_name: &str,
    line: u32,
) {
    let base = stash_exact(chunks, current, argc, 1, line);
    let chunk = &mut chunks[current];
    new_tagged(chunk, type_name, &[("__name__", base)], line);
}

/// `typing.NewType(name, tp)` → a callable that returns its argument, and
/// carries `__name__`/`__supertype__` the way CPython's does.
pub fn emit_newtype(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let call_idx = {
        // (this, value) — the Call slot is invoked with the object itself as
        // the leading argument, so the identity is local 1, not local 0.
        let mut helper = Chunk::new("__py_newtype_call");
        helper.arity = 2;
        helper.local_count = helper.local_count.max(2);
        helper.emit_op_u16(Op::LOCAL_GET, 1, line);
        helper.emit_op(Op::RETURN, line);
        chunks.push(helper);
        chunks.len() - 1
    };
    let base = stash_exact(chunks, current, argc, 2, line);
    let chunk = &mut chunks[current];
    new_tagged(
        chunk,
        "NewType",
        &[("__name__", base), ("__supertype__", base + 1)],
        line,
    );
    set_call_slot(chunk, call_idx, line);
}

/// `typing.get_type_hints(obj)` → the object's `__annotations__`, or `{}`.
/// Read through `ecma:object.get` because `struct.get` traps on anything that
/// is not a struct, and a function is a perfectly ordinary argument here.
pub fn emit_get_type_hints(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    let found = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], base, line);
    chunks[current].emit_string_const("__annotations__", line);
    let get_idx = chunks[current].add_import("ecma:object", "get");
    chunks[current].emit_call(get_idx, 2, line);
    lset(&mut chunks[current], found, line);

    lget(&mut chunks[current], found, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], found, line);
    chunks[current].emit_else(line);
    dict::emit_new(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_empty_tuple(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(true, line);
    struct_set(&mut chunks[current], &ClassSlot::internal("__tuple"), line);
}

fn emit_property_or(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    key: &str,
    fallback: impl FnOnce(&mut [Chunk], usize),
    line: u32,
) {
    let base = stash_exact(chunks, current, argc, 1, line);
    let found = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], base, line);
    chunks[current].emit_string_const(key, line);
    let get_idx = chunks[current].add_import("ecma:object", "get");
    chunks[current].emit_call(get_idx, 2, line);
    lset(&mut chunks[current], found, line);

    lget(&mut chunks[current], found, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], found, line);
    chunks[current].emit_else(line);
    fallback(chunks, current);
    chunks[current].emit_end(line);
}

/// `typing.get_origin(GenericAlias)` → alias.__origin__, else None.
pub fn emit_get_origin(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_property_or(
        chunks,
        current,
        argc,
        "__origin__",
        |chunks, current| {
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        },
        line,
    );
}

/// `typing.get_args(GenericAlias)` → alias.__args__, else ().
pub fn emit_get_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_property_or(
        chunks,
        current,
        argc,
        "__args__",
        |chunks, current| emit_empty_tuple(chunks, current, line),
        line,
    );
}

/// Runtime `TypedDict("Name", fields, ...)`: a marker type object. Static
/// annotations are erased by the walker; this gives `typing.is_typeddict` and
/// functional syntax a real value to carry around.
pub fn emit_typed_dict(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    let chunk = &mut chunks[current];
    new_tagged(chunk, "TypedDict", &[("__name__", base)], line);
    chunk.emit_dup(line);
    chunk.emit_bool_const(true, line);
    struct_set(chunk, &ClassSlot::internal("__is_typeddict__"), line);
}

pub fn emit_is_typeddict(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    let chunk = &mut chunks[current];
    lget(chunk, base, line);
    chunk.emit_string_const("__is_typeddict__", line);
    let has_own = chunk.add_import("ecma:object", "hasOwn");
    chunk.emit_call(has_own, 2, line);
}
