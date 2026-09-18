//! Python `functools` adapters.
//!
//! Metadata-copying is Python surface behavior; the storage itself is ordinary
//! ECMA object properties so callables stay compatible with the shared runtime.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn copy_attr(chunk: &mut Chunk, wrapper: u16, wrapped: u16, attr: &str, line: u32) {
    let get = chunk.add_import("ecma:object", "get");
    let set = chunk.add_import("ecma:object", "set");

    lget(chunk, wrapper, line);
    chunk.emit_string_const(attr, line);
    lget(chunk, wrapped, line);
    chunk.emit_string_const(attr, line);
    chunk.emit_call(get, 2, line);
    chunk.emit_call(set, 3, line);
    chunk.emit_op(Op::DROP, line);
}

pub fn emit_update_wrapper(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let n = argc.max(2);
    let base = chunks[current].alloc_scratch(n as u16);
    for offset in (0..argc as u16).rev() {
        lset(&mut chunks[current], base + offset, line);
    }
    for offset in argc as u16..n as u16 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        lset(&mut chunks[current], base + offset, line);
    }

    copy_attr(&mut chunks[current], base, base + 1, "__name__", line);
    copy_attr(&mut chunks[current], base, base + 1, "__qualname__", line);
    copy_attr(&mut chunks[current], base, base + 1, "__doc__", line);

    let set = chunks[current].add_import("ecma:object", "set");
    lget(&mut chunks[current], base, line);
    chunks[current].emit_string_const("__wrapped__", line);
    lget(&mut chunks[current], base + 1, line);
    chunks[current].emit_call(set, 3, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], base, line);
}
