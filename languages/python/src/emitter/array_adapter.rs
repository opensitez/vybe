//! Python `array` adapter — bytecode-only.
//!
//! An `array.array` is a list of numbers that also remembers its typecode.
//! This VM's arrays are Objects carrying a property map (the same mechanism
//! the `__tuple` tag uses), so the value here IS an array with `typecode` /
//! `itemsize` stamped on it. Indexing, `len`, `append`, `extend`, `count` and
//! `reverse` then come free from the list surface already in the profile, and
//! only the typecode-aware bits need emitting.
//!
//! No new host fns.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_compiler::primitives::instructions::core_wasm;

/// Width in bytes of each `array` typecode, per CPython's table.
const ITEMSIZES: &[(&str, i32)] = &[
    ("b", 1),
    ("B", 1),
    ("u", 4),
    ("h", 2),
    ("H", 2),
    ("i", 4),
    ("I", 4),
    ("l", 8),
    ("L", 8),
    ("q", 8),
    ("Q", 8),
    ("f", 4),
    ("d", 8),
];

fn struct_set(chunk: &mut Chunk, key: &str, line: u32) {
    let k = chunk.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_SET, k, line);
    chunk.emit_op(Op::DROP, line);
}

/// The typecode's item size, resolved from the code held in `tc`. The typecode
/// is a runtime string, so this is a comparison chain over the table above —
/// an unknown code falls through to 1, the smallest CPython defines.
/// Stack: `[]` → `[num]`.
fn emit_itemsize_of(chunk: &mut Chunk, tc: u16, line: u32) {
    let equals = chunk.add_import("wasm:js-string", "equals");
    for (code, size) in ITEMSIZES {
        chunk.emit_op_u16(Op::LOCAL_GET, tc, line);
        chunk.emit_string_const(code, line);
        chunk.emit_call(equals, 2, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if_value(line);
        core_wasm::i32_const(chunk, line, *size);
        chunk.emit_else(line);
    }
    core_wasm::i32_const(chunk, line, 1);
    for _ in ITEMSIZES {
        chunk.emit_end(line);
    }
}

/// `array.array(typecode[, initializer])`. Stack: `[tc, init?]` → `[array]`.
pub fn emit_array_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let init = chunks[current].alloc_scratch(1);
    let tc = chunks[current].alloc_scratch(1);

    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, init, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, tc, line);
    if argc < 2 {
        vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, init, line);
    }

    // Copy the initializer — `array.array('i', xs)` must not alias `xs`.
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, init, line);
    let from = chunk.add_import("ecma:array", "from");
    chunk.emit_call(from, 1, line);

    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::LOCAL_GET, tc, line);
    struct_set(chunk, "typecode", line);
    chunk.emit_dup(line);
    emit_itemsize_of(chunk, tc, line);
    struct_set(chunk, "itemsize", line);
    // `byteswap` reverses each item's bytes in place. Every typecode this
    // runtime stores is a plain number, so there are no stored bytes to
    // reverse and the operation is a no-op — but it has to EXIST, since
    // `hasattr(a, 'byteswap')` asks the object.
    chunk.emit_dup(line);
    chunk.emit_op(Op::NULL, line);
    struct_set(chunk, "byteswap", line);
}

/// `a.tolist()` — a plain list copy, without the array's stamps.
/// Stack: `[a]` → `[array]`.
pub fn emit_tolist(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    core_wasm::i32_const(chunk, line, 0);
    let slice = chunk.add_import("ecma:array", "slice");
    chunk.emit_call(slice, 2, line);
}

/// `a.buffer_info()` → `(address, length)`. There is no addressable buffer
/// behind a boxed list, so the address is 0; the length is what callers read.
/// Stack: `[a]` → `[tuple]`.
pub fn emit_buffer_info(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let a = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a, line);
    core_wasm::f64_const(&mut chunks[current], line, 0.0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    // Python returns a real tuple — built the one canonical way.
    vybe_compiler::primitives::tuples::emit_tuple(chunks, current, 2, line);
}

/// `a.frombytes(b)` — append each byte. Stack: `[a, b]` → `[null]`.
pub fn emit_frombytes(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let b = chunks[current].alloc_scratch(1);
    let a = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a, line);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(std::slice::from_mut(chunk), 0, line);

    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let push = chunk.add_import("ecma:array", "push");
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op(Op::NULL, line);
}
