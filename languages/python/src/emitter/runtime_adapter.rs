//! Python runtime-surface emitters.
//!
//! These are routed from the Python profile through `common:python.*`.
//! Keep Python-specific call shapes here instead of sending them through
//! the old runtime-helper function table.

use vybe_bytecode::opcode::Op;
use vybe_bytecode::Chunk;
use vybe_emitter::{collections, reflection, target::Target};

/// Python value-equality fallback for `==`/`!=` when no user `__eq__` is found.
/// Plain containers (lists/tuples/dicts — objects with no `__type` class stamp)
/// compare structurally via JSON; class instances and primitives keep
/// identity/primitive equality. Stack: `[a, b]` → `[bool i32]`.
pub fn emit_py_value_eq(chunk: &mut Chunk, line: u32) {
    let typeof_fn = chunk.add_import("ecma:value", "typeof");
    let is_array = chunk.add_import("ecma:array", "isArray");
    let has_own = chunk.add_import("ecma:object", "hasOwn");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let json_str = chunk.add_import("ecma:json", "stringify");
    let str_eq = chunk.add_import("wasm:js-string", "equals");
    let b = chunk.alloc_scratch(1);
    let a = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, b, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a, line);

    // Both operands sets → Python set equality: equal iff the sizes match and
    // one is a subset of the other (element identity, order-independent). A
    // JSON-structural compare would spuriously fail on differing insertion
    // order (`{1,2} == {2,1}`).
    emit_is_set(chunk, a, line);
    emit_is_set(chunk, b, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    {
        let size_key =
            chunk.add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("size")));
        chunk.emit_op_u16(Op::LOCAL_GET, a, line);
        chunk.emit_op_u16(Op::STRUCT_GET, size_key, line);
        chunk.emit_op_u16(Op::LOCAL_GET, b, line);
        chunk.emit_op_u16(Op::STRUCT_GET, size_key, line);
        vybe_emitter::ops::emit_dyn_eq(chunk, line); // sizes equal
        let sub = chunk.add_import("ecma:set", "isSubsetOf");
        chunk.emit_op_u16(Op::LOCAL_GET, a, line);
        chunk.emit_op_u16(Op::LOCAL_GET, b, line);
        chunk.emit_call(sub, 2, line);
        chunk.emit_op(Op::I32_AND, line);
    }
    chunk.emit_else(line);

    // structural = isArray(a) OR (typeof(a)=="object" AND a != null AND
    // !hasOwn(a, "__type")) — i.e. a plain list/tuple/dict, not a class instance.
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line); // i32: 1 if array
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_call(typeof_fn, 1, line);
    chunk.emit_string_const("object", line);
    vybe_emitter::ops::emit_dyn_eq(chunk, line); // i32: 1 if object
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line); // i32: 1 if not null
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_string_const("__type", line);
    chunk.emit_call(has_own, 2, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_op(Op::I32_EQZ, line); // i32: 1 if NOT a class instance
    chunk.emit_op(Op::I32_AND, line);
    // ...AND NOT a class object (has `__mro__`). Class objects carry a
    // self-referential `__mro__`, so JSON.stringify would recurse/throw;
    // `type(D()) is D` must be identity, never structural.
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_string_const("__mro__", line);
    chunk.emit_call(has_own, 2, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_op(Op::I32_EQZ, line); // i32: 1 if NOT a class object
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_OR, line); // array OR plain-object
    chunk.emit_if_value(line);
    // structural: JSON.stringify(a) == JSON.stringify(b)
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_call(json_str, 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    chunk.emit_call(json_str, 1, line);
    chunk.emit_call(str_eq, 2, line);
    chunk.emit_else(line);
    // identity / primitive equality
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    vybe_emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_end(line);
    chunk.emit_end(line); // close the set-equality outer if
}

/// Python `print(...)` — inline emitter that writes to `wasi:cli/stdout`
/// (like PHP `echo`), so `sep`/`end` and the missing trailing newline are
/// all expressible (the line-oriented `wasi:logging/logging.log` sink cannot
/// control separators or suppress the newline).
///
/// Argument convention set by the Python walker: when args are present,
/// `arg0 = sep`, `arg1 = end`, `args[2..] = items`. A bare `print()` compiles
/// with `argc == 0` and emits just the default end (`"\n"`). Each item is
/// converted to its Python display form via `emit_py_repr`.
pub fn emit_print(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let repr_idx = crate::emitter::repr_adapter::ensure_py_repr_chunk(chunks, line);
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
            emit_py_repr(chunk, repr_idx, line);
            part_count += 1;
        }
        chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
        part_count += 1;

        vybe_emitter::strings::emit_concat(chunk, part_count, line);
        chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    }

    vybe_emitter::io::emit_write_stdout_with_imports(
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
    emit_datetime_binop_or(
        chunk,
        a_slot,
        b_slot,
        DtOp::Add,
        "__add__",
        vybe_emitter::ops::emit_dyn_add,
        line,
    );
    chunk.emit_end(line);
}

/// `date`/`datetime`/`timedelta` arithmetic, else the normal dunder/dyn
/// path. The datetime values are plain objects, so they would otherwise
/// reach `emit_dyn_add`, which coerces through `wasm:js-number.toF64` and
/// throws on an object. Their arithmetic is defined in
/// `datetime_adapter::emit_dt_binop`; this only chooses when to use it.
fn emit_datetime_binop_or(
    chunk: &mut Chunk,
    a_slot: u16,
    b_slot: u16,
    op: DtOp,
    dunder: &str,
    fallback: fn(&mut Chunk, u32),
    line: u32,
) {
    crate::emitter::datetime_adapter::emit_is_datetime(chunk, a_slot, line);
    chunk.emit_if_value(line);
    crate::emitter::datetime_adapter::emit_dt_binop(chunk, a_slot, b_slot, op, line);
    chunk.emit_else(line);
    emit_object_binop_or(chunk, a_slot, b_slot, dunder, fallback, line);
    chunk.emit_end(line);
}

use crate::emitter::datetime_adapter::DtOp;

/// `-a` — a duration negates, an object may define `__neg__`, anything else
/// negates numerically. Stack: `[a]` → `[result]`.
pub fn emit_pyneg(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let a_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, a_slot, line);

    crate::emitter::datetime_adapter::emit_is_datetime(chunk, a_slot, line);
    chunk.emit_if_value(line);
    crate::emitter::datetime_adapter::emit_dt_neg(chunk, a_slot, line);
    chunk.emit_else(line);
    emit_unary_dunder_or(
        chunk,
        a_slot,
        "__neg__",
        vybe_emitter::ops::emit_dyn_neg,
        line,
    );
    chunk.emit_end(line);
}

/// `<`, `>`, `<=`, `>=`. Datetime values order by the instant or duration
/// they denote; an object may define the matching dunder; everything else
/// gets exactly the comparison the shared emitter already performed.
pub fn emit_pylt(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_relational(
        chunks,
        current,
        "__lt__",
        vybe_emitter::ops::emit_dyn_lt,
        line,
    );
}
pub fn emit_pygt(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_relational(
        chunks,
        current,
        "__gt__",
        vybe_emitter::ops::emit_dyn_gt,
        line,
    );
}
pub fn emit_pyle(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_relational(
        chunks,
        current,
        "__le__",
        vybe_emitter::ops::emit_dyn_le,
        line,
    );
}
pub fn emit_pyge(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_relational(
        chunks,
        current,
        "__ge__",
        vybe_emitter::ops::emit_dyn_ge,
        line,
    );
}

fn emit_relational(
    chunks: &mut [Chunk],
    current: usize,
    dunder: &str,
    cmp: fn(&mut Chunk, u32),
    line: u32,
) {
    let chunk = &mut chunks[current];
    let b_slot = chunk.alloc_scratch(1);
    let a_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a_slot, line);

    crate::emitter::datetime_adapter::emit_is_datetime(chunk, a_slot, line);
    chunk.emit_if_value(line);
    crate::emitter::datetime_adapter::emit_dt_cmp(chunk, a_slot, b_slot, cmp, line);
    chunk.emit_else(line);
    // `set </<=/>/>=` set → subset/superset per Python semantics. Both
    // operands must be sets; a lone set (`{1} < 2`) is a TypeError in
    // CPython, so mixed operands fall through to the numeric/object path.
    emit_is_set(chunk, a_slot, line);
    emit_is_set(chunk, b_slot, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    emit_set_relational(chunk, a_slot, b_slot, dunder, line);
    chunk.emit_else(line);
    emit_object_binop_or(chunk, a_slot, b_slot, dunder, cmp, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    // The comparison ops yield an i32; Python's `bool` is a real value.
    vybe_emitter::ops::emit_i32_to_bool(chunk, line);
}

/// `set` ordering: `<=`→subset, `<`→proper subset, `>=`→superset,
/// `>`→proper superset. Composes `ecma:set.isSubsetOf`/`isSupersetOf`; the
/// proper (`<`/`>`) forms additionally require unequal sizes. Both slots hold
/// Sets. Leaves an i32 bool on the stack.
fn emit_set_relational(chunk: &mut Chunk, a_slot: u16, b_slot: u16, dunder: &str, line: u32) {
    let (host_fn, strict) = match dunder {
        "__le__" => ("isSubsetOf", false),
        "__lt__" => ("isSubsetOf", true),
        "__ge__" => ("isSupersetOf", false),
        "__gt__" => ("isSupersetOf", true),
        _ => ("isSubsetOf", false),
    };
    let idx = chunk.add_import("ecma:set", host_fn);
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_call(idx, 2, line); // i32 bool
    if strict {
        // AND size(a) != size(b)
        let size_key =
            chunk.add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("size")));
        chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
        chunk.emit_op_u16(Op::STRUCT_GET, size_key, line);
        chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
        chunk.emit_op_u16(Op::STRUCT_GET, size_key, line);
        vybe_emitter::ops::emit_dyn_eq(chunk, line); // 1 if sizes equal
        chunk.emit_op(Op::I32_EQZ, line); // 1 if sizes differ
        chunk.emit_op(Op::I32_AND, line);
    }
}

/// A user `__neg__` when the operand is an object carrying one, else
/// `fallback`. Mirrors `emit_object_binop_or` for the one-operand case.
fn emit_unary_dunder_or(
    chunk: &mut Chunk,
    a_slot: u16,
    dunder: &str,
    fallback: fn(&mut Chunk, u32),
    line: u32,
) {
    let typeof_fn = chunk.add_import("ecma:value", "typeof");
    let key = chunk.add_constant(vybe_bytecode::Value::String(std::sync::Arc::from(dunder)));
    let method = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(typeof_fn, 1, line);
    chunk.emit_string_const("object", line);
    vybe_emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method, line);
    chunk.emit_op_u16(Op::LOCAL_GET, method, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, method, line);
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    fallback(chunk, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    fallback(chunk, line);
    chunk.emit_end(line);
}

/// Dispatch a binary operator to a user dunder when the left operand is a real
/// object (`typeof == "object"`, excluding arrays/strings/numbers) carrying
/// that method; otherwise emit `fallback`. Stack effect: consumes nothing
/// extra — reads `a_slot`/`b_slot`, pushes the result.
fn emit_object_binop_or(
    chunk: &mut Chunk,
    a_slot: u16,
    b_slot: u16,
    dunder: &str,
    fallback: fn(&mut Chunk, u32),
    line: u32,
) {
    let typeof_fn = chunk.add_import("ecma:value", "typeof");
    let key = chunk.add_constant(vybe_bytecode::Value::String(std::sync::Arc::from(dunder)));
    let method = chunk.alloc_scratch(1);
    // Only real objects (typeof == "object") can carry the dunder; STRUCT_GET on
    // a primitive traps, so gate the lookup behind the type check.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(typeof_fn, 1, line);
    chunk.emit_string_const("object", line);
    vybe_emitter::ops::emit_dyn_eq(chunk, line); // i32: 1 if object
    chunk.emit_if_value(line);
    // object: dispatch to the dunder if present, else fallback
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method, line);
    chunk.emit_op_u16(Op::LOCAL_GET, method, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line); // i32: 1 if present
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, method, line);
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    fallback(chunk, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    // primitive: fallback directly
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    fallback(chunk, line);
    chunk.emit_end(line);
}

/// Python `str(x)` — the same display form `print` uses (dict → `{'k': v}`,
/// list → `[..]`, True/False/None), so `str({'a':1})` isn't `[object Object]`.
/// Stack: `[x]` → `[string]`.
pub fn emit_str(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
        return;
    }
    let repr_idx = crate::emitter::repr_adapter::ensure_py_repr_chunk(chunks, line);
    emit_py_repr(&mut chunks[current], repr_idx, line);
}

/// Python `repr(x)` — the repr *form* (strings quoted, `__repr__` dispatch,
/// containers formatted recursively). Routes straight to the recursive
/// `__py_repr` chunk, unlike `str()` which uses the str-form top level.
pub fn emit_repr(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
        return;
    }
    let repr_idx = crate::emitter::repr_adapter::ensure_py_repr_chunk(chunks, line);
    let scratch = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, scratch, line);
    chunks[current].emit_op_u16(Op::REF_FUNC, repr_idx as u16, line);
    chunks[current].emit(0u8, line); // upvalue count
    chunks[current].emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunks[current].emit_op(Op::CALL_REF, line);
    chunks[current].emit(1u8, line);
}

/// Inline Python repr: Bool→True/False, None→None, Array→[elem, ...], else passthrough.
fn emit_py_repr(chunk: &mut Chunk, repr_idx: usize, line: u32) {
    let test_bool = chunk.add_import("wasm:js-boolean", "test");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let is_array = chunk.add_import("ecma:array", "isArray");
    let is_view = chunk.add_import("ecma:arraybuffer", "isView");
    let test_undef = chunk.add_import("wasm:js-undefined", "test");
    let scratch = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, scratch, line);

    // User-defined `__str__` (falling back to `__repr__`) dispatch: if the value
    // is an object carrying either method, call it and use its result. The
    // method lookup returns undefined for primitives, so `print(5)` etc. fall
    // straight through to the default formatting below.
    let get_method = std::sync::Arc::from("__vybe_js_get_method");
    let get_method_c = chunk.add_constant(vybe_bytecode::Value::String(get_method));
    let str_method = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::GLOBAL_GET, get_method_c, line);
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_string_const("__str__", line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, str_method, line);
    // fall back to __repr__ when __str__ is absent (statement-if: side effect
    // only, produces no stack value)
    chunk.emit_op_u16(Op::LOCAL_GET, str_method, line);
    chunk.emit_call(test_undef, 1, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::GLOBAL_GET, get_method_c, line);
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_string_const("__repr__", line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, str_method, line);
    chunk.emit_end(line);
    // Default formatting when there's no usable dunder OR the value is an array
    // (arrays return a non-function from the method lookup and must use the
    // list formatter below). Otherwise call the dunder with the receiver.
    chunk.emit_op_u16(Op::LOCAL_GET, str_method, line);
    chunk.emit_call(test_undef, 1, line); // i32: 1 if undefined
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line); // i32: 1 if array
    chunk.emit_op(Op::I32_OR, line); // default when undefined OR array
    chunk.emit_if_value(line);

    // bytes (Uint8Array / ArrayBuffer view) → Python `b'…'` via the source
    // prelude helper, which iterates/indexes the typed array natively.
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_call(is_view, 1, line);
    chunk.emit_if_value(line);
    let repr_fn = chunk.add_constant(vybe_bytecode::Value::String(std::sync::Arc::from(
        "__vybe_bytes_repr",
    )));
    chunk.emit_op_u16(Op::GLOBAL_GET, repr_fn, line);
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_else(line);

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
    // Recurse through the shared `__py_repr` chunk so NESTED tuples/lists/dicts
    // render correctly. `JSON.stringify` flattened every nested array to `[...]`
    // and erased the `__tuple`/`__typename` tags; the chunk walks each element
    // and renders it per Python repr rules at any depth.
    chunk.emit_op_u16(Op::REF_FUNC, repr_idx as u16, line);
    chunk.emit(0u8, line);
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_op(Op::CALL_REF, line);
    chunk.emit(1u8, line);
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
    vybe_emitter::strings::emit_to_string(chunk, line);
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
    // dict → recurse through `__py_repr` (keys and values repr'd; nested tuples
    // preserved), instead of `JSON.stringify` which flattened them.
    chunk.emit_op_u16(Op::REF_FUNC, repr_idx as u16, line);
    chunk.emit(0u8, line);
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_op(Op::CALL_REF, line);
    chunk.emit(1u8, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line); // close the isView (bytes) branch

    chunk.emit_else(line);
    // has __str__/__repr__ → call it with the receiver
    chunk.emit_op_u16(Op::LOCAL_GET, str_method, line);
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_end(line); // close the __str__-dispatch branch
}

/// Rewrite an already-formatted list string `"[1, 9]"` for a named tuple into
/// its Python repr `"P(x=1, y=9)"`, interleaving `scratch.__fields` with the
/// (already element-repr'd) values and prefixing `scratch.__typename`. The
/// values keep the formatting the list path produced; only the structure is
/// rebuilt here. Stack: `[list_string] -> [named_string]`.
fn emit_py_named_tuple_repr(chunk: &mut Chunk, scratch: u16, line: u32) {
    use vybe_bytecode::Value;
    use vybe_emitter::strings::emit_str_concat;
    let js_len = chunk.add_import("wasm:js-string", "length");
    let slice = chunk.add_import("ecma:array", "slice");
    let split = chunk.add_import("ecma:string", "split");

    let b = chunk.alloc_scratch(7);
    let (s, inner, parts, fields, result, i, n) = (b, b + 1, b + 2, b + 3, b + 4, b + 5, b + 6);
    chunk.emit_op_u16(Op::LOCAL_SET, s, line); // consume the list string

    // inner = s.slice(1, s.length - 1) — strip the `[` `]`.
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_call(js_len, 1, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_call(slice, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, inner, line);

    // parts = inner.split(", ") — the element strings (flat-value assumption).
    chunk.emit_op_u16(Op::LOCAL_GET, inner, line);
    chunk.emit_string_const(", ", line);
    chunk.emit_call(split, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, parts, line);

    // fields = scratch.__fields ; n = fields.length
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    let fk = chunk.add_constant(Value::String(std::sync::Arc::from(
        vybe_emitter::tuples::FIELDS_TAG,
    )));
    chunk.emit_op_u16(Op::STRUCT_GET, fk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, fields, line);
    chunk.emit_op_u16(Op::LOCAL_GET, fields, line);
    let len_k = chunk.add_constant(Value::String(std::sync::Arc::from("length")));
    chunk.emit_op_u16(Op::STRUCT_GET, len_k, line);
    chunk.emit_op_u16(Op::LOCAL_SET, n, line);

    // result = scratch.__typename + "("
    chunk.emit_op_u16(Op::LOCAL_GET, scratch, line);
    let tk = chunk.add_constant(Value::String(std::sync::Arc::from(
        vybe_emitter::tuples::TYPENAME_TAG,
    )));
    chunk.emit_op_u16(Op::STRUCT_GET, tk, line);
    chunk.emit_string_const("(", line);
    emit_str_concat(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);

    // i = 0
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);

    let state = vybe_emitter::loops::emit_loop_start(std::slice::from_mut(chunk), 0, line);
    // while i < n
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    vybe_emitter::ops::emit_dyn_lt(chunk, line);
    vybe_emitter::loops::emit_loop_cond(std::slice::from_mut(chunk), 0, line);

    // if i > 0 { result += ", " }
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_i32_const(0, line);
    vybe_emitter::ops::emit_dyn_gt(chunk, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    chunk.emit_string_const(", ", line);
    emit_str_concat(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_end(line);

    // result += fields[i] + "=" + parts[i]
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    chunk.emit_op_u16(Op::LOCAL_GET, fields, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    emit_str_concat(chunk, line);
    chunk.emit_string_const("=", line);
    emit_str_concat(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, parts, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    emit_str_concat(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);

    // i += 1
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_emitter::loops::emit_loop_end(std::slice::from_mut(chunk), 0, state, line);

    // result + ")"
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    chunk.emit_string_const(")", line);
    emit_str_concat(chunk, line);
}

/// Python `bytes.decode([encoding])` — UTF-8 decode a Uint8Array to a `str`.
/// The receiver is arg0; any encoding argument is ignored (UTF-8 only).
pub fn emit_bytes_decode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let n = argc.max(1);
    let base = chunk.alloc_scratch(n as u16);
    // Args on stack in call order (top = arg n-1). Pop into slots.
    for i in (0..n as u16).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + i, line);
    }
    let dec_new = chunk.add_import("web:encoding", "decoderNew");
    let dec = chunk.add_import("web:encoding", "decode");
    chunk.emit_call(dec_new, 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, base, line); // receiver = arg0 (bytes)
    chunk.emit_call(dec, 2, line);
}

/// Python `*` operator: array repeat, string repeat, or numeric multiply.
/// Stack: [a, b] → [result]
/// Emit `seq * count` sequence repetition: `newWithLength(count).fill(seq).flat(1)`.
/// Stack effect: pushes the repeated array (reads slots, leaves result on stack).
fn emit_array_repeat_slots(chunk: &mut Chunk, seq_slot: u16, count_slot: u16, line: u32) {
    let new_arr = chunk.add_import("ecma:array", "newWithLength");
    let fill = chunk.add_import("ecma:array", "fill");
    let flat = chunk.add_import("ecma:array", "flat");
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_call(new_arr, 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, seq_slot, line);
    chunk.emit_call(fill, 2, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_call(flat, 2, line);
}

pub fn emit_pymul(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let is_array = chunk.add_import("ecma:array", "isArray");
    let str_repeat = chunk.add_import("ecma:string", "repeat");
    let test_str = chunk.add_import("wasm:js-string", "test");
    let b_slot = chunk.alloc_scratch(1);
    let a_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a_slot, line);

    // Sequence repetition is commutative in Python: `seq * n` == `n * seq`. The
    // sequence may be either operand, so probe both sides before the numeric
    // fallback (`3 * 'ab'` and `'ab' * 3` both → 'ababab').

    // isArray(a): array repeat via newWithLength(n).fill(arr).flat(1)
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_if_value(line);
    emit_array_repeat_slots(chunk, a_slot, b_slot, line);
    chunk.emit_else(line);

    // string(a): a.repeat(b)
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(test_str, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_call(str_repeat, 2, line);
    chunk.emit_else(line);

    // reversed — string(b): b.repeat(a)  (`3 * 'ab'`)
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_call(test_str, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(str_repeat, 2, line);
    chunk.emit_else(line);

    // reversed — isArray(b): array repeat (`3 * [1, 2]`)
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_if_value(line);
    emit_array_repeat_slots(chunk, b_slot, a_slot, line);
    chunk.emit_else(line);

    // `timedelta * n`, else user `__mul__` on an object, else numeric multiply
    emit_datetime_binop_or(
        chunk,
        a_slot,
        b_slot,
        DtOp::Mul,
        "__mul__",
        emit_f64_mul,
        line,
    );
    chunk.emit_end(line); // isArray(b)
    chunk.emit_end(line); // string(b)
    chunk.emit_end(line); // string(a)
    chunk.emit_end(line); // isArray(a)
}

/// Numeric `*` fallback (`[a, b] → [a*b]`), for `emit_object_binop_or`.
fn emit_f64_mul(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_MUL, line);
}

// ── Remaining binary-operator dunders ───────────────────────────────────────
//
// `-`, `/`, `//`, `%`, `**` route through these (via `__pysub__` … builtins the
// walker emits). Each dispatches to the user dunder when the left operand is a
// real object carrying it, otherwise runs the same numeric fallback the shared
// compiler emits for that operator on the Python profile (so plain-number code
// is byte-for-byte identical — it takes the primitive branch).

/// Numeric `-` fallback, matching the Python-profile `F64_SUB`.
fn emit_f64_sub(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_SUB, line);
}

/// Numeric `/` fallback, matching the Python-profile `F64_DIV`.
fn emit_f64_div(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_DIV, line);
}

/// Numeric `//` fallback: `F64_DIV` then floor (Python-profile `BinOp::FloorDiv`).
fn emit_py_floordiv(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_DIV, line);
    vybe_emitter::math::emit_floor(chunk, line);
}

/// Numeric `%` fallback: Python floored modulo (Python-profile `BinOp::Mod`).
fn emit_py_mod(chunk: &mut Chunk, line: u32) {
    vybe_emitter::math::emit_python_floor_mod(chunk, line);
}

/// Numeric `**` fallback (Python-profile `BinOp::Pow`).
fn emit_py_pow(chunk: &mut Chunk, line: u32) {
    vybe_emitter::math::emit_pow(chunk, line);
}

/// `issubclass(sub, base)` — true when `base` is in `sub.__mro__` (the ancestor
/// class objects stamped by the shared class machinery). A class is a subclass
/// of itself, so `sub` is its own MRO head. Stack: `[sub, base]` → `[bool]`.
pub fn emit_issubclass(chunks: &mut [Chunk], current: usize, line: u32) {
    let base_slot = chunks[current].alloc_scratch(1);
    let sub_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, base_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sub_slot, line);
    // mro = sub.__mro__
    chunks[current].emit_op_u16(Op::LOCAL_GET, sub_slot, line);
    chunks[current].emit_string_const("__mro__", line);
    reflection::emit_get_property_in_chunk(&mut chunks[current], line);
    // base in mro  (collections::emit_contains: [array, value] → bool)
    chunks[current].emit_op_u16(Op::LOCAL_GET, base_slot, line);
    vybe_emitter::collections::emit_contains(chunks, current, line);
}

/// Python `vars(obj)` → the object's namespace dict. Builds a fresh dict from
/// the object's own enumerable entries (`ecma:object.entries` already skips the
/// `__`-prefixed internal stamps), so `vars(obj)['a']` reads a data attribute
/// rather than indexing a keys array. A bare `vars()` (argc 0) yields `{}`.
pub fn emit_vars(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        chunk.emit_op(Op::NULL, line);
        let new_obj = chunk.add_import("ecma:object", "create");
        chunk.emit_call(new_obj, 1, line);
        return;
    }
    let from_entries = chunk.add_import("ecma:object", "fromEntries");
    reflection::emit_object_view_in_chunk(chunk, reflection::ObjectKeysMode::Entries, line);
    chunk.emit_call(from_entries, 1, line); // {k: v, …}
}

/// Python `dir(obj)` → the names on the object AND on its class. `object.keys`
/// alone sees only instance attributes; class variables and methods live on the
/// class object, reached via the `__class__` link. Concatenates the two key sets
/// (`__`-internals already filtered by `keys`). Ordering/dedup is not yet
/// Python-exact, which is enough for membership (`'x' in dir(obj)`).
pub fn emit_dir(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let new_len = chunk.add_import("vybe:js-array", "newWithLength");
    if argc == 0 {
        // `dir()` with no argument → module/local names; not modelled. Empty list.
        chunk.emit_i32_const(0, line);
        chunk.emit_call(new_len, 1, line);
        return;
    }
    let concat = chunk.add_import("ecma:array", "concat");
    let obj = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);
    // instance keys
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    reflection::emit_object_view_in_chunk(chunk, reflection::ObjectKeysMode::Own, line);
    // class keys via the `__class__` link (empty if unlinked, so keys() is valid)
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_string_const("__class__", line);
    reflection::emit_get_property_in_chunk(chunk, line);
    let cls = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, cls, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cls, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_call(new_len, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, cls, line);
    reflection::emit_object_view_in_chunk(chunk, reflection::ObjectKeysMode::Own, line);
    chunk.emit_end(line);
    // instance_keys.concat(class_keys)
    chunk.emit_call(concat, 2, line);
}

/// Python `type(x)`. A user instance carries a `__class__` link to its class
/// object (stamped at construction, gated on `class_introspection_metadata`),
/// so `type(obj) is Cls` and `type(obj).__name__` resolve to the real class.
/// Non-instances fall back to `ecma:value.typeof` (unchanged builtin behavior).
pub fn emit_py_type(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let v = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, v, line);
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");

    // if v != null (guard so hasOwn never runs on null → TypeError)
    chunk.emit_op_u16(Op::LOCAL_GET, v, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    {
        // if hasOwn(v, "__class__") → v.__class__ else typeof(v)
        chunk.emit_op_u16(Op::LOCAL_GET, v, line);
        chunk.emit_string_const("__class__", line);
        reflection::emit_has_own_in_chunk(chunk, line);
        chunk.emit_call(cast_bool, 1, line);
        chunk.emit_if_value(line);
        chunk.emit_op_u16(Op::LOCAL_GET, v, line);
        chunk.emit_string_const("__class__", line);
        reflection::emit_get_property_in_chunk(chunk, line);
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, v, line);
        reflection::emit_typeof_in_chunk(chunk, line);
        chunk.emit_end(line);
    }
    chunk.emit_else(line);
    {
        chunk.emit_op_u16(Op::LOCAL_GET, v, line);
        reflection::emit_typeof_in_chunk(chunk, line);
    }
    chunk.emit_end(line);
}

/// Python `hasattr(obj, name)` routed through shared reflection property tests.
pub fn emit_hasattr(chunks: &mut [Chunk], current: usize, line: u32) {
    reflection::emit_has_in(chunks, current, line);
}

/// Python `getattr(obj, name[, default])`. The actual property access goes
/// through shared reflection; only the optional default is Python-specific.
pub fn emit_getattr(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc <= 2 {
        reflection::emit_get_property(chunks, current, line);
        return;
    }

    let default_slot = chunks[current].alloc_scratch(1);
    let name_slot = chunks[current].alloc_scratch(1);
    let obj_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, default_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    reflection::emit_has_in(chunks, current, line);
    let cast_bool = chunks[current].add_import("wasm:js-boolean", "cast");
    chunks[current].emit_call(cast_bool, 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    reflection::emit_get_property(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, default_slot, line);
    chunks[current].emit_end(line);
}

/// Python `setattr(obj, name, value)` returns `None`; the write itself goes
/// through shared reflection.
pub fn emit_setattr(chunks: &mut [Chunk], current: usize, line: u32) {
    reflection::emit_set_property(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// Python `delattr(obj, name)` returns `None`; deletion is the shared object
/// reflection operation.
pub fn emit_delattr(chunks: &mut [Chunk], current: usize, line: u32) {
    reflection::emit_object_op(chunks, current, reflection::ObjectOp::Delete, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// Push `1` when the value in `slot` is a Set (`typeof == "object"` and its
/// `__type` stamp is `"Set"`, which covers both `set` and `frozenset`), else
/// `0`. Guarded by the typeof check because `STRUCT_GET` traps on primitives.
fn emit_is_set(chunk: &mut Chunk, slot: u16, line: u32) {
    let typeof_fn = chunk.add_import("ecma:value", "typeof");
    let type_key = chunk.add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__type")));
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_call(typeof_fn, 1, line);
    chunk.emit_string_const("object", line);
    vybe_emitter::ops::emit_dyn_eq(chunk, line); // i32: 1 if object
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, type_key, line);
    chunk.emit_string_const("Set", line);
    vybe_emitter::ops::emit_dyn_eq(chunk, line); // i32: 1 if __type == "Set"
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
}

/// Python `hash(x)`. Mutable containers (`set`, `list`) are unhashable and
/// raise `TypeError`; everything else routes to the runtime hash helper. A
/// `dict` is a plain object here so it is not distinguished, but the common
/// cases (`hash({1})`, `hash([1])`) are covered.
fn emit_hash_guarded(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);

    emit_is_set(chunk, slot, line); // i32: 1 if set
    let is_array = chunk.add_import("ecma:array", "isArray");
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_call(is_array, 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line); // i32: 1 if list
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    {
        chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
        vybe_emitter::instructions::core_wasm::dup(chunk, line);
        chunk.emit_string_const("unhashable type", line);
        vybe_emitter::errors::emit_exception_new_finalize(chunk, "TypeError", line);
        vybe_emitter::errors::emit_throw(chunk, line);
    }
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    vybe_emitter::collections::emit_runtime_helper_call(chunks, current, "__vybe_hash", 1, line);
}

/// `a - b`. Python overloads `-` on sets to mean set difference (`{1,2,3} -
/// {2} == {1,3}`); the shared compiler already routes `|`/`&`/`^` to
/// `ecma:set` because those stay bitwise BinOps, but the walker desugars `-`
/// to `__pysub__`, so the same set check lives here. When both operands are
/// sets, call `ecma:set.difference`; otherwise dispatch `__sub__`/numeric.
pub fn emit_pysub(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let b_slot = chunk.alloc_scratch(1);
    let a_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a_slot, line);

    emit_is_set(chunk, a_slot, line);
    emit_is_set(chunk, b_slot, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    {
        chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
        let diff = chunk.add_import("ecma:set", "difference");
        chunk.emit_call(diff, 2, line);
    }
    chunk.emit_else(line);
    {
        chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
        emit_arith_dunder(chunk, "__sub__", emit_f64_sub, line);
    }
    chunk.emit_end(line);
}

/// `a / b` with `__truediv__` dispatch on object operands.
pub fn emit_pytruediv(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_arith_dunder(&mut chunks[current], "__truediv__", emit_f64_div, line);
}

/// `a // b` with `__floordiv__` dispatch on object operands.
pub fn emit_pyfloordiv(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_arith_dunder(&mut chunks[current], "__floordiv__", emit_py_floordiv, line);
}

/// `a % b` with `__mod__` dispatch on object operands.
/// Python `%` is overloaded: a string left operand means printf-style
/// formatting (`'%d' % 42`), a number means modulo, and an object may define
/// `__mod__`. Detect a string left at runtime and route to the shared sprintf
/// helper; otherwise fall through to the arithmetic/dunder path unchanged.
/// A tuple right operand spreads into positional args (Python semantics); a
/// scalar or list is a single argument.
pub fn emit_pymod(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let b_slot = chunks[current].alloc_scratch(1);
    let a_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a_slot, line);

    // if isString(a) → printf-style formatting
    chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
    let str_test = chunks[current].add_import("wasm:js-string", "test");
    chunks[current].emit_call(str_test, 1, line);
    chunks[current].emit_if_value(line);
    {
        // fmt = a
        chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
        // args = isTuple(b) ? b : [b]
        // Only a tagged tuple spreads into positional args; a scalar or list is
        // a single argument. Keyed on the tuple tag alone — an `isArray` guard
        // here misfires on an unboxed numeric operand.
        chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
        vybe_emitter::tuples::emit_is_tuple(chunks, current, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line); // tuple → spread
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line); // scalar/list → single arg
        chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 1, line);
        chunks[current].emit_end(line);
        // stack: [fmt, args] → formatted string
        vybe_emitter::sprintf::emit_sprintf_from_array(chunks, current, line);
    }
    chunks[current].emit_else(line);
    {
        chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
        emit_arith_dunder(&mut chunks[current], "__mod__", emit_py_mod, line);
    }
    chunks[current].emit_end(line);
}

/// `a ** b` with `__pow__` dispatch on object operands.
pub fn emit_pypow(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_arith_dunder(&mut chunks[current], "__pow__", emit_py_pow, line);
}

// ── Python format-spec numeric primitives ───────────────────────────────────
//
// Python's format mini-language is its own sprintf dialect, NOT C printf, so
// its numeric conversions compose the ECMA number primitives directly (their
// details match CPython where it matters — `toFixed`/`toExponential` use Rust
// formatting, i.e. round-half-to-even like Python, and render Infinity/NaN per
// spec). The walker parses the static spec and routes each numeric type here.

/// `__py_fmt_fixed(value, precision)` → fixed-point string via `toFixed`.
/// Backs `:.Nf` where the printf path is unreliable and `.N%` (as
/// `fixed(value*100, N) + "%"`). Stack: `[value, precision]` → `[string]`.
pub fn emit_py_fmt_fixed(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[current].add_import("ecma:number", "toFixed");
    chunks[current].emit_call(idx, 2, line);
}

/// `__py_fmt_sci(value, precision)` → lowercase scientific string. ECMA
/// `toExponential` emits a 1-digit exponent (`1.5e+3`); Python pads it to at
/// least two digits (`1.5e+03`), so split at `e` and `padStart` the exponent
/// digits. `:E` is this uppercased by the caller. Stack: `[value, precision]`
/// → `[string]`.
pub fn emit_py_fmt_sci(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let prec = c.alloc_scratch(1);
    let val = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, prec, line);
    c.emit_op_u16(Op::LOCAL_SET, val, line);

    // s = toExponential(val, prec)
    let s = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, val, line);
    c.emit_op_u16(Op::LOCAL_GET, prec, line);
    let toexp = c.add_import("ecma:number", "toExponential");
    c.emit_call(toexp, 2, line);
    c.emit_op_u16(Op::LOCAL_SET, s, line);

    // parts = s.split("e") → [mantissa, "+3"]
    let parts = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, s, line);
    c.emit_string_const("e", line);
    let split = c.add_import("ecma:string", "split");
    c.emit_call(split, 2, line);
    c.emit_op_u16(Op::LOCAL_SET, parts, line);

    let aget = c.add_import("ecma:array", "get");
    let exppart = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, parts, line);
    c.emit_f64_const(1.0, line);
    c.emit_call(aget, 2, line);
    c.emit_op_u16(Op::LOCAL_SET, exppart, line);

    // expSign = exppart.charAt(0)
    let expsign = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, exppart, line);
    c.emit_f64_const(0.0, line);
    let charat = c.add_import("ecma:string", "charAt");
    c.emit_call(charat, 2, line);
    c.emit_op_u16(Op::LOCAL_SET, expsign, line);

    // padded = exppart.slice(1).padStart(2, "0")
    let padded = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, exppart, line);
    c.emit_f64_const(1.0, line);
    let slice = c.add_import("ecma:string", "slice");
    c.emit_call(slice, 2, line);
    c.emit_f64_const(2.0, line);
    c.emit_string_const("0", line);
    let padstart = c.add_import("ecma:string", "padStart");
    c.emit_call(padstart, 3, line);
    c.emit_op_u16(Op::LOCAL_SET, padded, line);

    // result = mantissa + "e" + expSign + padded
    c.emit_op_u16(Op::LOCAL_GET, parts, line);
    c.emit_f64_const(0.0, line);
    c.emit_call(aget, 2, line);
    c.emit_string_const("e", line);
    c.emit_op_u16(Op::LOCAL_GET, expsign, line);
    c.emit_op_u16(Op::LOCAL_GET, padded, line);
    vybe_emitter::strings::emit_concat(c, 4, line);
}

/// `__py_fmt_group(str_value)` → insert `,` thousands separators into the
/// integer part of an already-formatted numeric string, preserving any sign and
/// fractional part (`"-1234.5"` → `"-1,234.5"`). Pure string surgery — printf
/// has no grouping. Stack: `[string]` → `[string]`.
pub fn emit_py_fmt_group(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let s = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, s, line);

    let length = c.add_import("wasm:js-string", "length");
    let charat = c.add_import("ecma:string", "charAt");
    let indexof = c.add_import("ecma:string", "indexOf");
    let slice = c.add_import("ecma:string", "slice");

    // dot = s.indexOf(".")  (−1 if none)
    let dot = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, s, line);
    c.emit_string_const(".", line);
    c.emit_call(indexof, 2, line);
    c.emit_op_u16(Op::LOCAL_SET, dot, line);

    // intEnd = dot < 0 ? s.length : dot   (dot < 0 → no fractional part)
    let int_end = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, dot, line);
    let tof64 = c.add_import("wasm:js-number", "toF64");
    c.emit_call(tof64, 1, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_if_value(line);
    c.emit_op_u16(Op::LOCAL_GET, s, line);
    c.emit_call(length, 1, line);
    c.emit_else(line);
    c.emit_op_u16(Op::LOCAL_GET, dot, line);
    c.emit_call(tof64, 1, line);
    c.emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_SET, int_end, line);

    // frac = s.slice(intEnd)  (includes the ".", or "" )
    let frac = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, s, line);
    c.emit_op_u16(Op::LOCAL_GET, int_end, line);
    c.emit_call(slice, 2, line);
    c.emit_op_u16(Op::LOCAL_SET, frac, line);

    // start = (s[0] is '+' or '-') ? 1 : 0  ; sign = s.slice(0, start)
    let start = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, s, line);
    c.emit_f64_const(0.0, line);
    c.emit_call(charat, 2, line);
    let sign_is = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, sign_is, line);
    // A leading '+'/'-' is a sign: "+-".includes(firstChar) AND firstChar != ""
    // (includes("") is vacuously true, so guard on a non-empty char).
    c.emit_string_const("+-", line);
    c.emit_op_u16(Op::LOCAL_GET, sign_is, line);
    let includes = c.add_import("ecma:string", "includes");
    c.emit_call(includes, 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(c, line);
    c.emit_op_u16(Op::LOCAL_GET, sign_is, line);
    c.emit_call(length, 1, line); // 1 if a real char, 0 if empty ("" not a sign)
    c.emit_op(Op::I32_AND, line);
    c.emit_if_value(line);
    c.emit_i32_const(1, line);
    c.emit_else(line);
    c.emit_i32_const(0, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_SET, start, line);

    // digits = s.slice(start, intEnd)  ; sign = s.slice(0, start)
    let sign = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, s, line);
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_GET, start, line);
    c.emit_call(slice, 3, line);
    c.emit_op_u16(Op::LOCAL_SET, sign, line);

    let digits = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, s, line);
    c.emit_op_u16(Op::LOCAL_GET, start, line);
    c.emit_op_u16(Op::LOCAL_GET, int_end, line);
    c.emit_call(slice, 3, line);
    c.emit_op_u16(Op::LOCAL_SET, digits, line);

    // n = digits.length ; out = "" ; i = 0
    let n = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_GET, digits, line);
    c.emit_call(length, 1, line);
    c.emit_op_u16(Op::LOCAL_SET, n, line);
    let out = c.alloc_scratch(1);
    c.emit_string_const("", line);
    c.emit_op_u16(Op::LOCAL_SET, out, line);
    let i = c.alloc_scratch(1);
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, i, line);

    // for i in 0..n: if i>0 and (n-i)%3==0: out += ","  ; out += digits.charAt(i)
    let block = c.emit_block(line);
    let (loop_p, _) = c.emit_loop_s(line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_op_u16(Op::LOCAL_GET, n, line);
    c.emit_op(Op::I32_GE_S, line);
    c.emit_br_if(1, line);
    // comma?  i>0 && (n-i)%3==0
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::I32_GT_S, line);
    c.emit_op_u16(Op::LOCAL_GET, n, line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_op(Op::I32_SUB, line);
    c.emit_i32_const(3, line);
    c.emit_op(Op::I32_REM_S, line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::I32_EQ, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, out, line);
    c.emit_string_const(",", line);
    vybe_emitter::strings::emit_str_concat(c, line);
    c.emit_op_u16(Op::LOCAL_SET, out, line);
    c.emit_end(line);
    // out += digits.charAt(i)
    c.emit_op_u16(Op::LOCAL_GET, out, line);
    c.emit_op_u16(Op::LOCAL_GET, digits, line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_call(charat, 2, line);
    vybe_emitter::strings::emit_str_concat(c, line);
    c.emit_op_u16(Op::LOCAL_SET, out, line);
    // i++
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, i, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(loop_p);
    c.emit_end(line);
    c.patch_block(block);

    // result = sign + out + frac
    c.emit_op_u16(Op::LOCAL_GET, sign, line);
    c.emit_op_u16(Op::LOCAL_GET, out, line);
    c.emit_op_u16(Op::LOCAL_GET, frac, line);
    vybe_emitter::strings::emit_concat(c, 3, line);
}

/// Shared body for the pure-arithmetic dunders: stash `[a, b]`, then dispatch to
/// `dunder` on an object left operand or run `fallback`.
fn emit_arith_dunder(chunk: &mut Chunk, dunder: &str, fallback: fn(&mut Chunk, u32), line: u32) {
    let b_slot = chunk.alloc_scratch(1);
    let a_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a_slot, line);
    // `-` is the only one of these datetime defines (`a - b`); the rest
    // fall straight through to the dunder/numeric path for every operand.
    if dunder == "__sub__" {
        emit_datetime_binop_or(chunk, a_slot, b_slot, DtOp::Sub, dunder, fallback, line);
    } else {
        emit_object_binop_or(chunk, a_slot, b_slot, dunder, fallback, line);
    }
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
    // `zip(a, b, …)` → array of tuples, stopping at the SHORTEST input (Python
    // semantics). Shared `vybe_emitter` op; `argc` = number of iterables.
    // `zip(a, b, …)` → array of tuples, stopping at the SHORTEST input (Python
    // semantics). Shared `vybe_emitter` op — now builds real (tagged) tuples
    // per row, so `list(zip(...))` reprs `[(a, b), …]`.
    if name == "python.zip" {
        collections::emit_zip(chunks, current, argc, collections::ZipLen::Shortest, line);
        return true;
    }
    // `random.choice(seq)` → one uniformly-random element. Shared op.
    if name == "python.random_choice" {
        vybe_emitter::random::emit_sample(chunks, current, argc, line);
        return true;
    }
    // `random.shuffle(seq)` → in-place Fisher-Yates. Shared op.
    if name == "python.random_shuffle" {
        vybe_emitter::random::emit_shuffle(chunks, current, argc, line);
        return true;
    }
    // `random.random()` → seedable uniform float in [0, 1).
    if name == "python.random" {
        vybe_emitter::random::emit_next_unit(chunks, current, line);
        return true;
    }
    // `random.randint(a, b)` → seedable int in [a, b] (inclusive).
    if name == "python.randint" {
        vybe_emitter::random::emit_rand_int_inclusive(chunks, current, line);
        return true;
    }
    // `random.sample(population, k)` → k unique elements (seedable partial
    // Fisher-Yates on a copy).
    if name == "python.random_sample" {
        vybe_emitter::random::emit_sample_k(chunks, current, line);
        return true;
    }
    // `random.seed(n)` — seed the global PRNG; `seed()` seeds from entropy.
    // Returns None.
    if name == "python.seed" {
        if argc == 0 {
            let r = chunks[current].add_import("ecma:math", "random");
            chunks[current].emit_call(r, 0, line);
            vybe_emitter::instructions::core_wasm::f64_const(
                &mut chunks[current],
                line,
                1073741824.0,
            );
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::I32_FROM_F64, line);
        }
        vybe_emitter::random::emit_seed(chunks, current, line);
        chunks[current].emit_op(Op::NULL, line);
        return true;
    }
    // `isinstance(obj, "Class")` → shared `classes::emit_instanceof`.
    if name == "python.instanceof" {
        vybe_emitter::classes::emit_instanceof(chunks, current, line);
        return true;
    }
    // `a is b` → object identity: reference equality for objects, value
    // identity for interned primitives (`emit_js_strict_eq` does exactly this —
    // REF_EQ on the object/cross-type branch, typed value compare otherwise).
    if name == "python.is_identity" {
        vybe_emitter::ops::emit_js_strict_eq(&mut chunks[current], line);
        vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
        return true;
    }
    if name == "python.is_not_identity" {
        vybe_emitter::ops::emit_js_strict_eq(&mut chunks[current], line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
        return true;
    }
    // `callable(x)` → shared reflection callable probe as a real Bool.
    if name == "python.callable" {
        reflection::emit_is_callable(chunks, current, line);
        vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
        return true;
    }
    // Regex adapters. Source calls are pattern-first — `re.findall(pat, s)`,
    // `re.split(pat, s)`, `re.sub(pat, repl, s)` — while `ecma:regexp` is
    // subject-first. Reorder args through scratch locals and call the host fn
    // directly instead of routing through the `__ecma_regexp_*_pat_first`
    // bundle chunks (which were themselves just this reorder + call).
    match name {
        "python.regex_findall" => {
            emit_python_findall(chunks, current, line);
            return true;
        }
        "python.regex_split" => {
            emit_regexp_pat_first(chunks, current, "split", line);
            return true;
        }
        "python.regex_sub" => {
            emit_regexp_replace_pat_first(chunks, current, line);
            return true;
        }
        _ => {}
    }
    if name == "python.hash" {
        emit_hash_guarded(chunks, current, line);
        return true;
    }
    let global = match name {
        "python.hex" => "__vybe_pyhex",
        "python.oct" => "__vybe_pyoct",
        "python.bin" => "__vybe_pybin",
        "python.bytes" | "python.encode" => "__vybe_to_bytes",
        "python.map" => "__vybe_pymap",
        "python.filter" => "__vybe_pyfilter",
        "python.any" => "__vybe_pyany",
        "python.all" => "__vybe_pyall",
        "python.iter" => "__vybe_pyiter",
        "python.next" => "__vybe_pynext",
        "python.isinf" => "__vybe_isinf",
        "python.id" => "__vybe_id",
        "python.hash" => "__vybe_hash",
        "python.format_map" => "__vybe_format_map",
        "python.setdefault" => "__vybe_setdefault",
        "python.tostring" => "__vybe_tostring",
        _ => return false,
    };
    collections::emit_runtime_helper_call(chunks, current, global, argc, line);
    true
}

/// Python `re.findall(pat, subject)` → `ecma:regexp.matchAll` flattened to
/// Python's result shape: no capture groups → the full match string; exactly
/// one group → that group; two+ groups → a tuple (array) of the groups.
/// Leverages the shared `loops::emit_for_in_*` scaffold. Stack `[pat, subject]`
/// → `[list]`.
fn emit_python_findall(chunks: &mut [Chunk], current: usize, line: u32) {
    use vybe_emitter::instructions::core_wasm::i32_const;
    use vybe_emitter::{collections, loops, ops};

    let (pat, subj, arr, result, idx, m, len, flat) = {
        let c = &mut chunks[current];
        (
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
        )
    };

    // arr = matchAll(subject, pat)  (pattern-first source → subject-first host)
    chunks[current].emit_op_u16(Op::LOCAL_SET, subj, line); // subject (stack top)
    chunks[current].emit_op_u16(Op::LOCAL_SET, pat, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, subj, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pat, line);
    let ma = chunks[current].add_import("ecma:regexp", "matchAll");
    chunks[current].emit_call(ma, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr, line);

    // result = []
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);

    // for m in arr:
    let state = loops::emit_for_in_start(chunks, current, arr, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, m, line); // element (a match array)

    // len = m.length
    chunks[current].emit_op_u16(Op::LOCAL_GET, m, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);

    // flat = (len <= 1) ? m[0]
    //      : (len <= 2) ? m[1]
    //      : m.slice(1, len)          # tuple of capture groups
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    i32_const(&mut chunks[current], line, 1);
    ops::emit_dyn_le(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, m, line);
    i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    i32_const(&mut chunks[current], line, 2);
    ops::emit_dyn_le(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, m, line);
    i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, m, line);
    i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    collections::emit_slice(chunks, current, line); // ecma:array.slice(m, 1, len)
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, flat, line);

    // result.push(flat)
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, flat, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    loops::emit_for_in_end(chunks, current, idx, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}

/// Pattern-first 2-arg regex adapter (`split`→`split`).
/// Stack `[pat, subject]` → `ecma:regexp.<method>(subject, pat)`.
fn emit_regexp_pat_first(chunks: &mut [Chunk], current: usize, method: &str, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    chunks[current].emit_op_u16(Op::LOCAL_SET, base + 1, line); // subject (top)
    chunks[current].emit_op_u16(Op::LOCAL_SET, base, line); // pat
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line); // subject
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line); // pat
    let idx = chunks[current].add_import("ecma:regexp", method);
    chunks[current].emit_call(idx, 2, line);
}

/// `re.sub(pat, repl, subject)` → `ecma:regexp.replaceAll(subject, pat, repl)`
/// (always-global, matching Python/PHP semantics).
fn emit_regexp_replace_pat_first(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    chunks[current].emit_op_u16(Op::LOCAL_SET, base + 2, line); // subject (top)
    chunks[current].emit_op_u16(Op::LOCAL_SET, base + 1, line); // repl
    chunks[current].emit_op_u16(Op::LOCAL_SET, base, line); // pat
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 2, line); // subject
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line); // pat
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line); // repl
    let idx = chunks[current].add_import("ecma:regexp", "replaceAll");
    chunks[current].emit_call(idx, 3, line);
}
