//! C strings backed by guest linear memory; managed values retain their adapter.
use vybe_compiler::primitives::{canon_marshal, instructions::host, pointers};
use vybe_runtime::{Chunk, Op};

fn get(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn scan(chunks: &mut [Chunk], current: usize, ptr: u16, limit: Option<u16>, line: u32) -> u16 {
    let c = &mut chunks[current];
    let len = c.alloc_scratch(1);
    c.emit_i32_const(0, line);
    set(c, len, line);
    let done = c.emit_block(line);
    let (repeat, _) = c.emit_loop_s(line);
    if let Some(limit) = limit {
        get(c, len, line);
        get(c, limit, line);
        c.emit_op(Op::I32_GE_U, line);
        c.emit_br_if(1, line);
    }
    get(c, ptr, line);
    get(c, len, line);
    c.emit_op(Op::I32_ADD, line);
    pointers::emit_linear_i32_load8_u(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op(Op::I32_EQZ, line);
    c.emit_br_if(1, line);
    get(c, len, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    set(c, len, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);
    len
}

pub fn emit_read(chunks: &mut [Chunk], current: usize, text: bool, line: u32) {
    let c = &mut chunks[current];
    let ptr = c.alloc_scratch(1);
    set(c, ptr, line);
    get(c, ptr, line);
    host::emit(c, "wasm:js-number", "test", 1, line);
    let branch = c.emit_if_value(line);
    let len = scan(chunks, current, ptr, None, line);
    let c = &mut chunks[current];
    if text {
        canon_marshal::emit_load_utf8(c, line, ptr, len);
    } else {
        get(c, len, line);
    }
    c.emit_else(line);
    get(c, ptr, line);
    host::emit(
        c,
        "libc:string",
        if text { "charToStr" } else { "strlenCString" },
        1,
        line,
    );
    c.emit_end(line);
    c.patch_block(branch);
}

/// Adapt immutable string backing to C byte storage before the shared copy.
/// This encodes the whole string, including embedded NULs, plus its terminator.
pub fn emit_byte_copy(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_encoded_byte_copy(chunks, current, false, line);
}

/// FILE's legacy string storage maps one code point to one byte, not UTF-8.
pub fn emit_binary_byte_copy(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_encoded_byte_copy(chunks, current, true, line);
}

fn emit_encoded_byte_copy(chunks: &mut [Chunk], current: usize, binary: bool, line: u32) {
    let c = &mut chunks[current];
    let first = c.alloc_scratch(3);
    let (dst, src, count) = (first, first + 1, first + 2);
    set(c, count, line);
    set(c, src, line);
    set(c, dst, line);
    get(c, src, line);
    host::emit(c, "wasm:js-string", "test", 1, line);
    let string_source = c.emit_if(line);
    if binary {
        get(c, src, line);
        c.emit_string_const("latin1", line);
        host::emit(c, "node:buffer", "from", 2, line);
    } else {
        host::emit(c, "web:encoding", "encoderNew", 0, line);
        get(c, src, line);
        c.emit_string_const("\0", line);
        host::emit(c, "wasm:js-string", "concat", 2, line);
        host::emit(c, "web:encoding", "encode", 2, line);
    }
    set(c, src, line);
    c.emit_end(line);
    c.patch_block(string_source);
    get(c, dst, line);
    get(c, src, line);
    get(c, count, line);
    pointers::emit_byte_copy(chunks, current, line);
}

/// Search byte-addressed C storage, returning an address rather than a suffix.
pub fn emit_strrchr_linear(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let first = c.alloc_scratch(4);
    let (ptr, needle, found, byte) = (first, first + 1, first + 2, first + 3);
    c.emit_i32_const(255, line);
    c.emit_op(Op::I32_AND, line);
    set(c, needle, line);
    set(c, ptr, line);
    c.emit_i32_const(0, line);
    set(c, found, line);
    let done = c.emit_block(line);
    let (repeat, _) = c.emit_loop_s(line);
    get(c, ptr, line);
    pointers::emit_linear_i32_load8_u(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op_u16(Op::LOCAL_TEE, byte, line);
    get(c, needle, line);
    c.emit_op(Op::I32_EQ, line);
    let matched = c.emit_if(line);
    get(c, ptr, line);
    set(c, found, line);
    c.emit_end(line);
    c.patch_block(matched);
    get(c, byte, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_br_if(1, line);
    get(c, ptr, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    set(c, ptr, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);
    get(c, found, line);
}

pub fn emit_strncpy(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let first = c.alloc_scratch(5);
    let (dst, src, n, bytes, count) = (first, first + 1, first + 2, first + 3, first + 4);
    set(c, n, line);
    set(c, src, line);
    set(c, dst, line);
    get(c, dst, line);
    host::emit(c, "wasm:js-number", "test", 1, line);
    let destination = c.emit_if_value(line);
    get(c, src, line);
    set(c, bytes, line);
    get(c, src, line);
    host::emit(c, "wasm:js-number", "test", 1, line);
    let source = c.emit_if(line);
    let len = scan(chunks, current, src, Some(n), line);
    let c = &mut chunks[current];
    get(c, len, line);
    set(c, count, line);
    c.emit_else(line);
    host::emit(c, "web:encoding", "encoderNew", 0, line);
    get(c, src, line);
    host::emit(c, "libc:string", "charToStr", 1, line);
    host::emit(c, "web:encoding", "encode", 2, line);
    set(c, bytes, line);
    get(c, bytes, line);
    host::emit(c, "ecma:array", "length", 1, line);
    get(c, n, line);
    host::emit(c, "ecma:math", "min", 2, line);
    set(c, count, line);
    c.emit_end(line);
    c.patch_block(source);
    get(c, dst, line);
    get(c, bytes, line);
    get(c, count, line);
    pointers::emit_byte_copy(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op(Op::DROP, line);
    get(c, dst, line);
    get(c, count, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_i32_const(0, line);
    get(c, n, line);
    get(c, count, line);
    c.emit_op(Op::I32_SUB, line);
    pointers::emit_linear_memory_fill(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op(Op::DROP, line);
    get(c, dst, line);
    c.emit_else(line);
    get(c, dst, line);
    get(c, src, line);
    get(c, n, line);
    host::emit(c, "libc:string", "strncpyCarray", 3, line);
    c.emit_end(line);
    c.patch_block(destination);
}

pub fn emit_pointer(chunks: &mut [Chunk], current: usize, write: bool, line: u32) {
    let c = &mut chunks[current];
    let argc = if write { 4 } else { 2 };
    let first = c.alloc_scratch(argc);
    for i in (0..argc).rev() {
        set(c, first + i, line);
    }
    get(c, first, line);
    host::emit(c, "wasm:js-number", "test", 1, line);
    let branch = c.emit_if_value(line);
    get(c, first, line);
    get(c, first + 1, line);
    c.emit_op(Op::I32_ADD, line);
    if write {
        get(c, first + 2, line);
        pointers::emit_linear_i32_store8(chunks, current, line);
        let c = &mut chunks[current];
        c.emit_op(Op::DROP, line);
        get(c, first, line);
    }
    let c = &mut chunks[current];
    c.emit_else(line);
    for i in 0..argc {
        get(c, first + i, line);
    }
    host::emit(
        c,
        "libc:string",
        if write { "charPtrWrite" } else { "charPtrAdd" },
        argc as u8,
        line,
    );
    c.emit_end(line);
    c.patch_block(branch);
}

pub fn emit_strdup(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let first = c.alloc_scratch(3);
    let (src, count, dst) = (first, first + 1, first + 2);
    set(c, src, line);
    get(c, src, line);
    host::emit(c, "wasm:js-number", "test", 1, line);
    let source = c.emit_if(line);
    let len = scan(chunks, current, src, None, line);
    let c = &mut chunks[current];
    get(c, len, line);
    set(c, count, line);
    c.emit_else(line);
    host::emit(c, "web:encoding", "encoderNew", 0, line);
    get(c, src, line);
    host::emit(c, "libc:string", "charToStr", 1, line);
    host::emit(c, "web:encoding", "encode", 2, line);
    set(c, src, line);
    get(c, src, line);
    host::emit(c, "ecma:array", "length", 1, line);
    set(c, count, line);
    c.emit_end(line);
    c.patch_block(source);
    get(c, count, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    pointers::emit_linear_alloc(chunks, current, line);
    let c = &mut chunks[current];
    set(c, dst, line);
    get(c, dst, line);
    get(c, src, line);
    get(c, count, line);
    pointers::emit_byte_copy(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op(Op::DROP, line);
    get(c, dst, line);
    get(c, count, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_i32_const(0, line);
    pointers::emit_linear_i32_store8(chunks, current, line);
    let c = &mut chunks[current];
    c.emit_op(Op::DROP, line);
    get(c, dst, line);
}
