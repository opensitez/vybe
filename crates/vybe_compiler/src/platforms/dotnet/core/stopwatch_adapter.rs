use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

pub fn emit_stopwatch_start_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let new_idx = chunks[0].add_import("wasi:clocks", "stopwatchNew");
    let start_idx = chunks[0].add_import("wasi:clocks", "stopwatchStart");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, new_idx, line);
    chunk.emit(0, line);
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, start_idx, line);
    chunk.emit(1, line);
    chunk.emit_op(Op::DROP, line);
}

pub fn emit_stopwatch_restart(chunks: &mut [Chunk], current: usize, line: u32) {
    let reset_idx = chunks[0].add_import("wasi:clocks", "stopwatchReset");
    let start_idx = chunks[0].add_import("wasi:clocks", "stopwatchStart");
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, reset_idx, line);
    chunk.emit(1, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, start_idx, line);
    chunk.emit(1, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::NULL, line);
}

pub fn emit_stopwatch_elapsed_ms(chunks: &mut [Chunk], current: usize, line: u32) {
    let elapsed_idx = chunks[0].add_import("wasi:clocks", "stopwatchElapsed");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, elapsed_idx, line);
    chunk.emit(1, line);
}

pub fn emit_stopwatch_is_running(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let key = chunk.add_constant(Value::String(Arc::from("isrunning")));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
}
