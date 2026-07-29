//! Smoke tests for the `emitter::loops` helpers. Each helper takes
//! `(chunks: &mut [Chunk], current: usize, ...slot args, line: u32)` so
//! the imports it registers land on `chunks[0]` (module-level WASM
//! convention). These tests drive them through a one-element slice.

use vybe_bytecode::Chunk;
use vybe_compiler::primitives::loops;

fn one_chunk(local_count: u16) -> Vec<Chunk> {
    let mut c = Chunk::new("test");
    c.local_count = local_count;
    vec![c]
}

#[test]
fn emit_map_produces_bytecode() {
    let mut chunks = one_chunk(10);
    loops::emit_map(&mut chunks, 0, 1, 2, 3, 4, 0);
    assert!(
        chunks[0].code.len() > 20,
        "map should emit substantial bytecode"
    );
}

#[test]
fn emit_filter_produces_bytecode() {
    let mut chunks = one_chunk(10);
    loops::emit_filter(&mut chunks, 0, 1, 2, 3, 4, 5, 0);
    assert!(
        chunks[0].code.len() > 20,
        "filter should emit substantial bytecode"
    );
}

#[test]
fn emit_foreach_produces_bytecode() {
    let mut chunks = one_chunk(10);
    loops::emit_foreach(&mut chunks, 0, 1, 2, 3, 0);
    assert!(chunks[0].code.len() > 15, "foreach should emit bytecode");
}

#[test]
fn emit_reduce_produces_bytecode() {
    let mut chunks = one_chunk(10);
    loops::emit_reduce(&mut chunks, 0, 1, 2, 3, 4, 0);
    assert!(
        chunks[0].code.len() > 20,
        "reduce should emit substantial bytecode"
    );
}

#[test]
fn emit_any_produces_bytecode() {
    let mut chunks = one_chunk(10);
    loops::emit_any_every(&mut chunks, 0, 1, 2, 3, true, 0);
    assert!(chunks[0].code.len() > 15, "any should emit bytecode");
}

#[test]
fn emit_every_produces_bytecode() {
    let mut chunks = one_chunk(10);
    loops::emit_any_every(&mut chunks, 0, 1, 2, 3, false, 0);
    assert!(chunks[0].code.len() > 15, "every should emit bytecode");
}

#[test]
fn emit_for_in_produces_loop() {
    let mut chunks = one_chunk(10);
    let state = loops::emit_for_in_start(&mut chunks, 0, 1, 2, 0);
    // element is on stack here — drop it to simulate body
    chunks[0].emit_op(vybe_bytecode::opcode::Op::DROP, 0);
    loops::emit_for_in_end(&mut chunks, 0, 2, state, 0);
    assert!(
        chunks[0].code.len() > 10,
        "for-in should emit loop bytecode"
    );
}
