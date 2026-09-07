//! Adapter helpers for Python `difflib`.

use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn emit_python_string(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    lget(&mut chunks[current], slot, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
}

/// `difflib.SequenceMatcher.ratio` style score.
///
/// Python accepts generic sequences; this adapter applies Python's own string
/// conversion at the edge, then reuses the shared similarity primitive for the
/// matching count. Stack: `[a, b] -> [float ratio]`.
pub fn emit_ratio(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 2..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    let b = chunks[current].alloc_scratch(1);
    let a = chunks[current].alloc_scratch(1);
    let as_text = chunks[current].alloc_scratch(1);
    let bs_text = chunks[current].alloc_scratch(1);
    let a_len = chunks[current].alloc_scratch(1);
    let b_len = chunks[current].alloc_scratch(1);
    let total = chunks[current].alloc_scratch(1);

    if argc >= 2 {
        lset(&mut chunks[current], b, line);
    } else {
        chunks[current].emit_string_const("", line);
        lset(&mut chunks[current], b, line);
    }
    if argc >= 1 {
        lset(&mut chunks[current], a, line);
    } else {
        chunks[current].emit_string_const("", line);
        lset(&mut chunks[current], a, line);
    }

    emit_python_string(chunks, current, a, line);
    lset(&mut chunks[current], as_text, line);
    emit_python_string(chunks, current, b, line);
    lset(&mut chunks[current], bs_text, line);

    lget(&mut chunks[current], as_text, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    lset(&mut chunks[current], a_len, line);
    lget(&mut chunks[current], bs_text, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    lset(&mut chunks[current], b_len, line);

    lget(&mut chunks[current], a_len, line);
    lget(&mut chunks[current], b_len, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    lset(&mut chunks[current], total, line);

    lget(&mut chunks[current], total, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], as_text, line);
    lget(&mut chunks[current], bs_text, line);
    vybe_compiler::primitives::string_similarity::emit_similar_text(chunks, current, 2, line);
    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    lget(&mut chunks[current], total, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_end(line);
}
