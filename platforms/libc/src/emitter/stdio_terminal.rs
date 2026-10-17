//! Numeric FILE-handle transport and existing WASI standard-terminal queries.
use vybe_runtime::{Chunk, opcode::Op};

pub fn emit_fileno(chunk: &mut Chunk, line: u32) {
    // The libc FILE registry uses the descriptor itself as its numeric handle.
    chunk.emit_op(Op::I64_TRUNC_F64_S, line);
    chunk.emit_op(Op::I32_WRAP_I64, line);
}

pub fn emit_isatty(chunk: &mut Chunk, line: u32) {
    let fd = chunk.alloc_scratch(1);
    emit_fileno(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, fd, line);
    for (number, module, getter) in [
        (0, "wasi:cli/terminal-stdin", "get-terminal-stdin"),
        (1, "wasi:cli/terminal-stdout", "get-terminal-stdout"),
        (2, "wasi:cli/terminal-stderr", "get-terminal-stderr"),
    ] {
        chunk.emit_op_u16(Op::LOCAL_GET, fd, line);
        chunk.emit_i32_const(number, line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_if_value(line);
        let import = chunk.add_import(module, getter);
        chunk.emit_call(import, 0, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_else(line);
    }
    // Nonstandard handles in this virtual FILE registry are not terminals.
    chunk.emit_i32_const(0, line);
    for _ in 0..3 {
        chunk.emit_end(line);
    }
}
