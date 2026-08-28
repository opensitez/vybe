//! `System.IO.MemoryStream` — a seekable byte stream.
//!
//! ## The value
//!
//! ```text
//! { __type: "MemoryStream",
//!   __ms_buf:  <array>,   the backing store; its LENGTH is `Capacity`
//!   __ms_pos:  <number>,  the cursor
//!   __ms_len:  <number>,  the logical length, always <= Capacity
//!   __ms_expandable, __ms_writable, __ms_visible, __ms_closed: <bool> }
//! ```
//!
//! ⛔ Keys are `__`-prefixed and LOWERCASE. A dotnet type with no property
//! accessor resolves its properties as an ordinary lowercased struct-field
//! read, so a PascalCase key is unreadable from a case-insensitive frontend —
//! the convention `thread_adapter`'s `CANCELLED_KEY` documents.
//!
//! `Capacity`, `Length`, `Position`, `CanRead`, `CanWrite` and `CanSeek` are
//! COMPUTED properties declared in `tree_register::shared_emit_accessors`, not
//! stored fields: a stored `Length` goes stale on every write, and
//! `Capacity`'s setter has to resize and can throw.

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::errors;
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::ops;

use super::thread_adapter::DELAY_TOKEN_KEY;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::object_fields::field_slot;

const TYPE_KEY: &str = "__type";
const BUF: &str = "__ms_buf";
const POS: &str = "__ms_pos";
const LEN: &str = "__ms_len";
const EXPANDABLE: &str = "__ms_expandable";
const WRITABLE: &str = "__ms_writable";
const VISIBLE: &str = "__ms_visible";
const CLOSED: &str = "__ms_closed";

/// .NET's growth rule (`MemoryStream.EnsureCapacity`): at least 256, and at
/// least double the current capacity.
const MIN_GROWTH: f64 = 256.0;

fn push_str(chunk: &mut Chunk, text: &str, line: u32) {
    chunk.emit_string_const(text, line);
}

fn num(chunk: &mut Chunk, v: f64, line: u32) {
    chunk.emit_f64_const(v, line);
}

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn call(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, name);
    chunk.emit_call(idx, argc, line);
}

/// `[obj, value] → []`
fn field_set(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
}

/// `[] → [value]` — a field of the object in `slot`.
fn field(chunk: &mut Chunk, slot: u16, key: &str, line: u32) {
    get(chunk, slot, line);
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        Dest::Stack,
        line,
    );
}

/// `[value] → []` — store into a field of the object in `slot`.
fn store(chunk: &mut Chunk, slot: u16, key: &str, line: u32) {
    let tmp = chunk.alloc_scratch(1);
    set(chunk, tmp, line);
    get(chunk, slot, line);
    get(chunk, tmp, line);
    field_set(chunk, key, line);
}

/// Throw a .NET exception of `class` with `message`.
fn throw(chunk: &mut Chunk, class: &str, message: &str, line: u32) {
    class_slots::emit_class_alloc(chunk, line);
    core_wasm::dup(chunk, line);
    push_str(chunk, message, line);
    errors::emit_exception_new_finalize(chunk, class, line);
    errors::emit_throw(chunk, line);
}

/// `if <cond on stack> { throw }`
fn throw_if(chunk: &mut Chunk, class: &str, message: &str, line: u32) {
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    throw(chunk, class, message, line);
    chunk.emit_end(line);
}

/// Every member starts here: a closed stream answers `ObjectDisposedException`
/// to everything, which is what `Close()` then `WriteByte(1)` must do.
fn guard_open(chunk: &mut Chunk, obj: u16, line: u32) {
    field(chunk, obj, CLOSED, line);
    throw_if(
        chunk,
        "ObjectDisposedException",
        "Cannot access a closed Stream.",
        line,
    );
}

/// `[] → [capacity]`
fn capacity_of(chunk: &mut Chunk, obj: u16, line: u32) {
    field(chunk, obj, BUF, line);
    call(chunk, "ecma:array", "length", 1, line);
}

/// Grow the backing store so `needed` bytes fit, or throw on a fixed stream.
/// Stack: `[needed] → []`
fn ensure_capacity(chunk: &mut Chunk, obj: u16, line: u32) {
    let needed = chunk.alloc_scratch(3);
    let cap = needed + 1;
    let target = needed + 2;
    set(chunk, needed, line);
    capacity_of(chunk, obj, line);
    set(chunk, cap, line);

    get(chunk, needed, line);
    get(chunk, cap, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    // A stream over a caller-supplied buffer cannot grow — .NET answers
    // `NotSupportedException`.
    field(chunk, obj, EXPANDABLE, line);
    ops::emit_dyn_not(chunk, line);
    throw_if(
        chunk,
        "NotSupportedException",
        "Memory stream is not expandable.",
        line,
    );

    // .NET's rule: at least 256, and at least double what we have.
    get(chunk, needed, line);
    set(chunk, target, line);
    get(chunk, target, line);
    num(chunk, MIN_GROWTH, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    num(chunk, MIN_GROWTH, line);
    set(chunk, target, line);
    chunk.emit_end(line);
    get(chunk, target, line);
    get(chunk, cap, line);
    num(chunk, 2.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    get(chunk, cap, line);
    num(chunk, 2.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    set(chunk, target, line);
    chunk.emit_end(line);

    get(chunk, target, line);
    emit_resize_buffer(chunk, obj, line);
    chunk.emit_end(line);
}

/// Resize the backing store to the length on the stack, zero-filling any new
/// tail. Stack: `[new_capacity] → []`
fn emit_resize_buffer(chunk: &mut Chunk, obj: u16, line: u32) {
    let want = chunk.alloc_scratch(3);
    let fresh = want + 1;
    let i = want + 2;
    set(chunk, want, line);

    get(chunk, want, line);
    call(chunk, "ecma:array", "newWithLength", 1, line);
    set(chunk, fresh, line);
    // `newWithLength` leaves holes; a MemoryStream's unwritten bytes are ZERO.
    get(chunk, fresh, line);
    num(chunk, 0.0, line);
    call(chunk, "ecma:array", "fill", 2, line);
    chunk.emit_op(Op::DROP, line);

    num(chunk, 0.0, line);
    set(chunk, i, line);
    let block = chunk.emit_block(line);
    let (lp, _) = chunk.emit_loop_s(line);
    get(chunk, i, line);
    get(chunk, want, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, i, line);
    capacity_of(chunk, obj, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, fresh, line);
    get(chunk, i, line);
    field(chunk, obj, BUF, line);
    get(chunk, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    get(chunk, i, line);
    num(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(lp);
    chunk.emit_end(line);
    chunk.patch_block(block);

    get(chunk, fresh, line);
    store(chunk, obj, BUF, line);
}

/// The shared constructor tail: stamp an object built from `buf`/`len`/flags.
/// Stack: `[] → [stream]`
#[allow(clippy::too_many_arguments)]
fn build_stream(
    chunks: &mut Vec<Chunk>,
    current: usize,
    buf: u16,
    len: u16,
    expandable: bool,
    writable: bool,
    visible: bool,
    line: u32,
) {
    let idx = chunks[current].add_import("ecma:object", "new");
    chunks[current].emit_call(idx, 0, line);
    let obj = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    set(chunk, obj, line);

    get(chunk, obj, line);
    push_str(chunk, "MemoryStream", line);
    field_set(chunk, TYPE_KEY, line);
    get(chunk, obj, line);
    get(chunk, buf, line);
    field_set(chunk, BUF, line);
    get(chunk, obj, line);
    num(chunk, 0.0, line);
    field_set(chunk, POS, line);
    get(chunk, obj, line);
    get(chunk, len, line);
    field_set(chunk, LEN, line);
    for (key, value) in [
        (EXPANDABLE, expandable),
        (WRITABLE, writable),
        (VISIBLE, visible),
        (CLOSED, false),
    ] {
        get(chunk, obj, line);
        chunk.emit_bool_const(value, line);
        field_set(chunk, key, line);
    }
    get(chunk, obj, line);
}

/// `New MemoryStream()` / `(capacity)` / `(buffer)` / `(buffer, offset, count)`
/// / `(buffer, offset, count, writable, publiclyVisible)`.
pub fn emit_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    match argc {
        0 => {
            let (buf, len) = {
                let chunk = &mut chunks[current];
                let buf = chunk.alloc_scratch(2);
                let len = buf + 1;
                num(chunk, 0.0, line);
                call(chunk, "ecma:array", "newWithLength", 1, line);
                set(chunk, buf, line);
                num(chunk, 0.0, line);
                set(chunk, len, line);
                (buf, len)
            };
            build_stream(chunks, current, buf, len, true, true, true, line);
        }
        1 => {
            // `MemoryStream(capacity)` or `MemoryStream(buffer)` — the
            // argument's KIND decides, exactly as the overloads do.
            let (buf, len, arg) = {
                let chunk = &mut chunks[current];
                let arg = chunk.alloc_scratch(3);
                let buf = arg + 1;
                let len = arg + 2;
                set(chunk, arg, line);
                (buf, len, arg)
            };
            let chunk = &mut chunks[current];
            get(chunk, arg, line);
            call(chunk, "ecma:array", "isArray", 1, line);
            ops::emit_dyn_to_bool(chunk, line);
            chunk.emit_if(line);
            get(chunk, arg, line);
            set(chunk, buf, line);
            get(chunk, arg, line);
            call(chunk, "ecma:array", "length", 1, line);
            set(chunk, len, line);
            chunk.emit_else(line);
            get(chunk, arg, line);
            call(chunk, "ecma:array", "newWithLength", 1, line);
            set(chunk, buf, line);
            get(chunk, buf, line);
            num(chunk, 0.0, line);
            call(chunk, "ecma:array", "fill", 2, line);
            chunk.emit_op(Op::DROP, line);
            num(chunk, 0.0, line);
            set(chunk, len, line);
            chunk.emit_end(line);

            // A buffer-backed stream is FIXED; a capacity-backed one expands.
            let expandable = chunks[current].alloc_scratch(1);
            let chunk = &mut chunks[current];
            get(chunk, arg, line);
            call(chunk, "ecma:array", "isArray", 1, line);
            ops::emit_dyn_not(chunk, line);
            set(chunk, expandable, line);
            build_stream_dynamic(chunks, current, buf, len, expandable, line);
        }
        _ => {
            // `(buffer, offset, count[, writable, publiclyVisible])`.
            let (buf, len, visible, writable) = {
                let chunk = &mut chunks[current];
                let scratch = chunk.alloc_scratch(6);
                let visible = scratch;
                let writable = scratch + 1;
                let count = scratch + 2;
                let offset = scratch + 3;
                let src = scratch + 4;
                let buf = scratch + 5;
                chunk.emit_bool_const(true, line);
                set(chunk, visible, line);
                chunk.emit_bool_const(true, line);
                set(chunk, writable, line);
                if argc >= 5 {
                    set(chunk, visible, line);
                }
                if argc >= 4 {
                    set(chunk, writable, line);
                }
                set(chunk, count, line);
                set(chunk, offset, line);
                set(chunk, src, line);
                get(chunk, src, line);
                get(chunk, offset, line);
                get(chunk, offset, line);
                get(chunk, count, line);
                chunk.emit_op(Op::F64_ADD, line);
                call(chunk, "ecma:array", "slice", 3, line);
                set(chunk, buf, line);
                (buf, count, visible, writable)
            };
            build_stream_visible(chunks, current, buf, len, visible, writable, line);
        }
    }
}

/// [`build_stream`] with `expandable` taken from a slot rather than a literal.
fn build_stream_dynamic(
    chunks: &mut Vec<Chunk>,
    current: usize,
    buf: u16,
    len: u16,
    expandable: u16,
    line: u32,
) {
    build_stream(chunks, current, buf, len, true, true, true, line);
    let obj = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    set(chunk, obj, line);
    get(chunk, obj, line);
    get(chunk, expandable, line);
    field_set(chunk, EXPANDABLE, line);
    get(chunk, obj, line);
}

/// [`build_stream`] for the slice overloads: fixed, with `publiclyVisible`
/// from a slot.
fn build_stream_visible(
    chunks: &mut Vec<Chunk>,
    current: usize,
    buf: u16,
    len: u16,
    visible: u16,
    writable: u16,
    line: u32,
) {
    build_stream(chunks, current, buf, len, false, true, true, line);
    let obj = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    set(chunk, obj, line);
    for (key, slot) in [(VISIBLE, visible), (WRITABLE, writable)] {
        get(chunk, obj, line);
        get(chunk, slot, line);
        field_set(chunk, key, line);
    }
    get(chunk, obj, line);
}

// ── Properties ───────────────────────────────────────────────────────────
//
// Computed, not stored: `Length` changes on every write and `Capacity`'s
// setter has to resize and can throw. They are declared in
// `tree_register::shared_emit_accessors`.

/// `[stream] → [n]`
pub fn emit_capacity(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj = chunk.alloc_scratch(1);
    set(chunk, obj, line);
    capacity_of(chunk, obj, line);
}

/// `[stream] → [n]`
pub fn emit_length(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj = chunk.alloc_scratch(1);
    set(chunk, obj, line);
    field(chunk, obj, LEN, line);
}

/// `[stream] → [n]`
pub fn emit_position(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj = chunk.alloc_scratch(1);
    set(chunk, obj, line);
    field(chunk, obj, POS, line);
}

/// `[stream, n] → []`
pub fn emit_set_position(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(2);
    let obj = value + 1;
    set(chunk, value, line);
    set(chunk, obj, line);
    guard_open(chunk, obj, line);
    get(chunk, value, line);
    store(chunk, obj, POS, line);
}

/// `[stream, n] → []` — resize, with .NET's two refusals.
pub fn emit_set_capacity(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(2);
    let obj = value + 1;
    set(chunk, value, line);
    set(chunk, obj, line);
    guard_open(chunk, obj, line);

    // Below the logical length the data would be lost — .NET refuses.
    get(chunk, value, line);
    field(chunk, obj, LEN, line);
    ops::emit_dyn_lt(chunk, line);
    throw_if(
        chunk,
        "ArgumentOutOfRangeException",
        "capacity was less than the current size.",
        line,
    );
    field(chunk, obj, EXPANDABLE, line);
    ops::emit_dyn_not(chunk, line);
    throw_if(
        chunk,
        "NotSupportedException",
        "Memory stream is not expandable.",
        line,
    );
    get(chunk, value, line);
    emit_resize_buffer(chunk, obj, line);
}

fn emit_const_bool_property(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    let chunk = &mut chunks[current];
    let obj = chunk.alloc_scratch(1);
    set(chunk, obj, line);
    field(chunk, obj, key, line);
}

/// `CanRead` / `CanSeek` — true for any stream that is not closed.
pub fn emit_can_read(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj = chunk.alloc_scratch(1);
    set(chunk, obj, line);
    field(chunk, obj, CLOSED, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_i32_to_bool(chunk, line);
}

pub fn emit_can_seek(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_can_read(chunks, current, line);
}

pub fn emit_can_write(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_const_bool_property(chunks, current, WRITABLE, line);
}

// ── Reading and writing ──────────────────────────────────────────────────

/// `stream.WriteByte(b)`. `[stream, b] → []`
pub fn emit_write_byte(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let byte = chunk.alloc_scratch(2);
    let obj = byte + 1;
    set(chunk, byte, line);
    set(chunk, obj, line);
    guard_open(chunk, obj, line);
    emit_store_byte_at_position(chunk, obj, byte, line);
}

/// Write one byte at the cursor, growing and extending as .NET does.
fn emit_store_byte_at_position(chunk: &mut Chunk, obj: u16, byte: u16, line: u32) {
    field(chunk, obj, POS, line);
    num(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    ensure_capacity(chunk, obj, line);

    field(chunk, obj, BUF, line);
    field(chunk, obj, POS, line);
    get(chunk, byte, line);
    chunk.emit_op(Op::ARRAY_SET, line);

    field(chunk, obj, POS, line);
    num(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    store(chunk, obj, POS, line);

    // The logical length only ever grows to the furthest byte written.
    field(chunk, obj, POS, line);
    field(chunk, obj, LEN, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    field(chunk, obj, POS, line);
    store(chunk, obj, LEN, line);
    chunk.emit_end(line);
}

/// `stream.Write(buffer[, offset, count])` — the 1-arg form is the
/// `Span`/`ReadOnlySpan` overload, which writes the whole thing.
pub fn emit_write(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(5);
    let count = scratch;
    let offset = scratch + 1;
    let src = scratch + 2;
    let obj = scratch + 3;
    let i = scratch + 4;
    if argc >= 3 {
        set(chunk, count, line);
        set(chunk, offset, line);
    }
    set(chunk, src, line);
    set(chunk, obj, line);
    if argc < 3 {
        num(chunk, 0.0, line);
        set(chunk, offset, line);
        get(chunk, src, line);
        call(chunk, "ecma:array", "length", 1, line);
        set(chunk, count, line);
    }
    guard_open(chunk, obj, line);

    num(chunk, 0.0, line);
    set(chunk, i, line);
    let block = chunk.emit_block(line);
    let (lp, _) = chunk.emit_loop_s(line);
    get(chunk, i, line);
    get(chunk, count, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);
    let byte = chunk.alloc_scratch(1);
    get(chunk, src, line);
    get(chunk, offset, line);
    get(chunk, i, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    set(chunk, byte, line);
    emit_store_byte_at_position(chunk, obj, byte, line);
    get(chunk, i, line);
    num(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(lp);
    chunk.emit_end(line);
    chunk.patch_block(block);
}

/// `stream.ReadByte()` — the byte at the cursor, or `-1` at the end.
pub fn emit_read_byte(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj = chunk.alloc_scratch(2);
    let out = obj + 1;
    set(chunk, obj, line);
    guard_open(chunk, obj, line);
    num(chunk, -1.0, line);
    set(chunk, out, line);
    field(chunk, obj, POS, line);
    field(chunk, obj, LEN, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    field(chunk, obj, BUF, line);
    field(chunk, obj, POS, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    set(chunk, out, line);
    field(chunk, obj, POS, line);
    num(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    store(chunk, obj, POS, line);
    chunk.emit_end(line);
    get(chunk, out, line);
}

/// `stream.Read(buffer[, offset, count])` → the number of bytes read.
pub fn emit_read(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(6);
    let count = scratch;
    let offset = scratch + 1;
    let dst = scratch + 2;
    let obj = scratch + 3;
    let i = scratch + 4;
    let avail = scratch + 5;
    if argc >= 3 {
        set(chunk, count, line);
        set(chunk, offset, line);
    }
    set(chunk, dst, line);
    set(chunk, obj, line);
    if argc < 3 {
        num(chunk, 0.0, line);
        set(chunk, offset, line);
        get(chunk, dst, line);
        call(chunk, "ecma:array", "length", 1, line);
        set(chunk, count, line);
    }
    guard_open(chunk, obj, line);

    // Never read past the logical end.
    field(chunk, obj, LEN, line);
    field(chunk, obj, POS, line);
    chunk.emit_op(Op::F64_SUB, line);
    set(chunk, avail, line);
    get(chunk, count, line);
    get(chunk, avail, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    get(chunk, avail, line);
    set(chunk, count, line);
    chunk.emit_end(line);
    get(chunk, count, line);
    num(chunk, 0.0, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    num(chunk, 0.0, line);
    set(chunk, count, line);
    chunk.emit_end(line);

    num(chunk, 0.0, line);
    set(chunk, i, line);
    let block = chunk.emit_block(line);
    let (lp, _) = chunk.emit_loop_s(line);
    get(chunk, i, line);
    get(chunk, count, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, dst, line);
    get(chunk, offset, line);
    get(chunk, i, line);
    chunk.emit_op(Op::F64_ADD, line);
    field(chunk, obj, BUF, line);
    field(chunk, obj, POS, line);
    get(chunk, i, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    get(chunk, i, line);
    num(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(lp);
    chunk.emit_end(line);
    chunk.patch_block(block);

    field(chunk, obj, POS, line);
    get(chunk, count, line);
    chunk.emit_op(Op::F64_ADD, line);
    store(chunk, obj, POS, line);
    get(chunk, count, line);
}

/// `stream.Seek(offset, origin)` → the new position.
///
/// `SeekOrigin` is `Begin`/`Current`/`End` = 0/1/2.
pub fn emit_seek(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(4);
    let origin = scratch;
    let offset = scratch + 1;
    let obj = scratch + 2;
    let base = scratch + 3;
    set(chunk, origin, line);
    set(chunk, offset, line);
    set(chunk, obj, line);
    guard_open(chunk, obj, line);

    num(chunk, 0.0, line);
    set(chunk, base, line);
    get(chunk, origin, line);
    num(chunk, 1.0, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    field(chunk, obj, POS, line);
    set(chunk, base, line);
    chunk.emit_end(line);
    get(chunk, origin, line);
    num(chunk, 2.0, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    field(chunk, obj, LEN, line);
    set(chunk, base, line);
    chunk.emit_end(line);

    get(chunk, base, line);
    get(chunk, offset, line);
    chunk.emit_op(Op::F64_ADD, line);
    store(chunk, obj, POS, line);
    field(chunk, obj, POS, line);
}

/// `stream.SetLength(n)` — truncate, or extend with zeros.
pub fn emit_set_length(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(2);
    let obj = value + 1;
    set(chunk, value, line);
    set(chunk, obj, line);
    guard_open(chunk, obj, line);

    // Growing past the fixed buffer is what `NotSupportedException` is for.
    get(chunk, value, line);
    ensure_capacity(chunk, obj, line);

    // Bytes between the old and new length must read as ZERO, whatever the
    // buffer happened to be carrying there.
    let i = chunk.alloc_scratch(1);
    field(chunk, obj, LEN, line);
    set(chunk, i, line);
    let block = chunk.emit_block(line);
    let (lp, _) = chunk.emit_loop_s(line);
    get(chunk, i, line);
    get(chunk, value, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);
    field(chunk, obj, BUF, line);
    get(chunk, i, line);
    num(chunk, 0.0, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    get(chunk, i, line);
    num(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(lp);
    chunk.emit_end(line);
    chunk.patch_block(block);

    get(chunk, value, line);
    store(chunk, obj, LEN, line);
    // A cursor past the new end is clamped, as .NET does.
    field(chunk, obj, POS, line);
    get(chunk, value, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    get(chunk, value, line);
    store(chunk, obj, POS, line);
    chunk.emit_end(line);
}

/// `stream.ToArray()` — a COPY of the logical content, capacity excluded.
pub fn emit_to_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj = chunk.alloc_scratch(1);
    set(chunk, obj, line);
    field(chunk, obj, BUF, line);
    num(chunk, 0.0, line);
    field(chunk, obj, LEN, line);
    call(chunk, "ecma:array", "slice", 3, line);
}

/// `stream.GetBuffer()` — the BACKING STORE itself.
pub fn emit_get_buffer(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj = chunk.alloc_scratch(1);
    set(chunk, obj, line);
    field(chunk, obj, VISIBLE, line);
    ops::emit_dyn_not(chunk, line);
    throw_if(
        chunk,
        "UnauthorizedAccessException",
        "MemoryStream's internal buffer cannot be accessed.",
        line,
    );
    field(chunk, obj, BUF, line);
}

/// `stream.TryGetBuffer(segment)` — answers the SEGMENT; the out-param desugar
/// turns a non-null into `True`.
pub fn emit_try_get_buffer(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let obj = {
        let chunk = &mut chunks[current];
        let obj = chunk.alloc_scratch(1);
        set(chunk, obj, line);
        obj
    };
    let idx = chunks[current].add_import("ecma:object", "new");
    chunks[current].emit_call(idx, 0, line);
    let seg = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    set(chunk, seg, line);
    get(chunk, seg, line);
    push_str(chunk, "ArraySegment", line);
    field_set(chunk, TYPE_KEY, line);
    for key in ["Array", "array"] {
        get(chunk, seg, line);
        field(chunk, obj, BUF, line);
        field_set(chunk, key, line);
    }
    for key in ["Offset", "offset"] {
        get(chunk, seg, line);
        num(chunk, 0.0, line);
        field_set(chunk, key, line);
    }
    for key in ["Count", "count"] {
        get(chunk, seg, line);
        field(chunk, obj, LEN, line);
        field_set(chunk, key, line);
    }
    get(chunk, seg, line);
}

/// `stream.WriteTo(other)` — the logical content, appended to `other`.
pub fn emit_write_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let dest = chunk.alloc_scratch(2);
    let obj = dest + 1;
    set(chunk, dest, line);
    set(chunk, obj, line);
    guard_open(chunk, obj, line);
    get(chunk, dest, line);
    field(chunk, obj, BUF, line);
    num(chunk, 0.0, line);
    field(chunk, obj, LEN, line);
    call(chunk, "ecma:array", "slice", 3, line);
    num(chunk, 0.0, line);
    field(chunk, obj, LEN, line);
    emit_write(chunks, current, 4, line);
    let chunk = &mut chunks[current];
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `stream.Close()` / `Dispose()`.
pub fn emit_close(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj = chunk.alloc_scratch(1);
    set(chunk, obj, line);
    get(chunk, obj, line);
    chunk.emit_bool_const(true, line);
    field_set(chunk, CLOSED, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `stream.Flush()` — nothing is buffered, so this only checks the stream is
/// open.
pub fn emit_flush(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj = chunk.alloc_scratch(1);
    set(chunk, obj, line);
    guard_open(chunk, obj, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `src.CopyTo(dest[, bufferSize])` — copy from the SOURCE'S CURRENT POSITION
/// to its end, into `dest` at ITS current position, advancing both.
///
/// ⛔ Not `WriteTo`. `WriteTo` writes the whole stream from 0 regardless of the
/// cursor; `CopyTo` starts where the reader left off and leaves the source
/// positioned at the end. A test that seeks and then copies tells them apart.
///
/// `bufferSize` is accepted and only VALIDATED — .NET throws
/// `ArgumentOutOfRangeException` for a non-positive size, and the chunking it
/// otherwise controls is unobservable here because the copy is in-memory.
pub fn emit_copy_to(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(4);
    let size = scratch;
    let dest = scratch + 1;
    let obj = scratch + 2;
    let count = scratch + 3;

    // ⛔ `argc >= 3` is the RECEIVER plus two arguments. Do not "fix" it to 2:
    // measured, that breaks the one-argument `CopyTo(dest)` (8 tests). The
    // two-argument form does not reach here with three either — the dotnet
    // lookup resolves both `CopyTo` registrations to the FIRST one, so
    // `CopyTo(dest, 0)` arrives with the size still on the stack and pops it
    // as the destination. That is the overload-resolution defect, not this
    // test's shape.
    if argc >= 3 {
        set(chunk, size, line);
    } else {
        num(chunk, 1.0, line);
        set(chunk, size, line);
    }
    set(chunk, dest, line);
    set(chunk, obj, line);
    guard_open(chunk, obj, line);

    // .NET validates the destination before anything else.
    get(chunk, dest, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    throw_if(
        chunk,
        "ArgumentNullException",
        "Value cannot be null. (Parameter 'destination')",
        line,
    );
    get(chunk, size, line);
    num(chunk, 1.0, line);
    ops::emit_dyn_lt(chunk, line);
    throw_if(
        chunk,
        "ArgumentOutOfRangeException",
        "Positive number required. (Parameter 'bufferSize')",
        line,
    );

    // count = len - pos, clamped at zero: a cursor past the end copies nothing.
    field(chunk, obj, LEN, line);
    field(chunk, obj, POS, line);
    chunk.emit_op(Op::F64_SUB, line);
    set(chunk, count, line);
    get(chunk, count, line);
    num(chunk, 0.0, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    num(chunk, 0.0, line);
    set(chunk, count, line);
    chunk.emit_end(line);

    // dest.Write(src.buf.slice(pos, len), 0, count)
    get(chunk, dest, line);
    field(chunk, obj, BUF, line);
    field(chunk, obj, POS, line);
    field(chunk, obj, LEN, line);
    call(chunk, "ecma:array", "slice", 3, line);
    num(chunk, 0.0, line);
    get(chunk, count, line);
    emit_write(chunks, current, 4, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);

    // The source is left AT THE END — that is what makes a second CopyTo a
    // no-op, which .NET programs rely on.
    field(chunk, obj, LEN, line);
    store(chunk, obj, POS, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `src.CopyToAsync(dest[, bufferSize | cancellationToken])` — the same copy,
/// handed back as an already-settled task.
///
/// The copy is in-memory, so there is nothing to defer; returning a resolved
/// promise is what makes `Await src.CopyToAsync(dest)` and `.Result` both
/// answer, and it is the same object JS reaches (ECMA-262 §27.2.4.7).
pub fn emit_copy_to_async(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // A third argument is a CancellationToken, not a buffer size. It is not
    // discarded: the task carries it, and `task.IsCanceled` asks the token —
    // `CopyToAsync(dest, size, cts.Token)` on a cancelled source answers True
    // in .NET, and a resolved task with nothing on it answered False.
    let token = if argc >= 3 {
        let slot = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
        Some(slot)
    } else {
        None
    };
    emit_copy_to(chunks, current, 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    core_wasm::undefined(chunk, line);
    call(chunk, "ecma:promise", "resolve", 1, line);
    if let Some(slot) = token {
        core_wasm::dup(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        class_slots::emit_class_set(
            chunk,
            ObjSource::Stack,
            &field_slot(DELAY_TOKEN_KEY),
            ValueSource::Stack,
            line,
        );
    }
}

/// `[value] → [Task]` — settle what the sync operation already produced.
///
/// Every `*Async` member on a MemoryStream is the synchronous operation plus a
/// completed task: the bytes are in memory, so there is nothing to defer. The
/// promise is the same object JS reaches (ECMA-262 §27.2.4.7), which is what
/// lets `Await ms.FlushAsync()` and `.Result` both answer.
fn settle(chunks: &mut [Chunk], current: usize, line: u32) {
    call(&mut chunks[current], "ecma:promise", "resolve", 1, line);
}

/// `FlushAsync([cancellationToken])`.
pub fn emit_flush_async(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 2 {
        chunks[current].emit_op(Op::DROP, line);
    }
    emit_flush(chunks, current, line);
    settle(chunks, current, line);
}

/// `DisposeAsync()` — closes, then settles.
pub fn emit_dispose_async(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_close(chunks, current, line);
    settle(chunks, current, line);
}

/// `WriteAsync(buffer[, offset, count][, cancellationToken])`.
///
/// ⛔ The trailing `CancellationToken` has to be dropped BEFORE the write, or
/// it is taken for `count` and the write copies a token's worth of bytes.
pub fn emit_write_async(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let effective = if argc == 3 || argc == 5 {
        chunks[current].emit_op(Op::DROP, line);
        argc - 1
    } else {
        argc
    };
    emit_write(chunks, current, effective, line);
    settle(chunks, current, line);
}

/// `ReadAsync(buffer[, offset, count][, cancellationToken])` — settles with the
/// BYTE COUNT, which is what the caller awaits.
pub fn emit_read_async(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let effective = if argc == 3 || argc == 5 {
        chunks[current].emit_op(Op::DROP, line);
        argc - 1
    } else {
        argc
    };
    emit_read(chunks, current, effective, line);
    settle(chunks, current, line);
}

/// `CanTimeout` — always false. A MemoryStream has no device behind it, so
/// .NET reports false and `ReadTimeout`/`WriteTimeout` throw
/// `InvalidOperationException` rather than answering.
pub fn emit_can_timeout(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    core_wasm::bool_const(chunk, line, false);
}
