//! PHP time/sleep adapters.
//!
//! PHP exposes `sleep(seconds)` and `usleep(microseconds)`. WASI does not have
//! a `wasi:clocks.sleep` function; current WASI clocks provide
//! `wasi:clocks/monotonic-clock.wait-for(duration)`. Route through the shared
//! threading primitive so PHP uses the same sleep machinery as other
//! languages while preserving PHP's units and return shape.

use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

pub fn emit_php_sleep(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_sleep_scaled_to_millis(chunks, current, argc, 1000.0, line);
}

pub fn emit_php_usleep(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_sleep_scaled_to_millis(chunks, current, argc, 0.001, line);
}

fn emit_sleep_scaled_to_millis(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    millis_per_unit: f64,
    line: u32,
) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        chunk.emit_i32_const(0, line);
        return;
    }

    crate::emitter::numeric_adapter::emit_php_floatval(chunks, current, 1, line);

    let wait_for_idx = chunks[current].add_import("wasi:clocks/monotonic-clock", "wait-for");
    let chunk = &mut chunks[current];
    chunk.emit_f64_const(millis_per_unit, line);
    chunk.emit_op(Op::F64_MUL, line);
    vybe_compiler::primitives::threading::emit_thread_sleep(chunk, wait_for_idx, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_i32_const(0, line);
}
