use vybe_bytecode::Chunk;
use vybe_compiler_common::loops;

#[test]
fn emit_map_produces_bytecode() {
    let mut chunk = Chunk::new("test");
    // Need enough locals allocated
    chunk.local_count = 10;
    loops::emit_map(&mut chunk, 1, 2, 3, 4, 0);
    assert!(chunk.code.len() > 20, "map should emit substantial bytecode");
}

#[test]
fn emit_filter_produces_bytecode() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 10;
    loops::emit_filter(&mut chunk, 1, 2, 3, 4, 5, 0);
    assert!(chunk.code.len() > 20, "filter should emit substantial bytecode");
}

#[test]
fn emit_foreach_produces_bytecode() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 10;
    loops::emit_foreach(&mut chunk, 1, 2, 3, 0);
    assert!(chunk.code.len() > 15, "foreach should emit bytecode");
}

#[test]
fn emit_reduce_produces_bytecode() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 10;
    loops::emit_reduce(&mut chunk, 1, 2, 3, 4, 0);
    assert!(chunk.code.len() > 20, "reduce should emit substantial bytecode");
}

#[test]
fn emit_any_produces_bytecode() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 10;
    loops::emit_any_every(&mut chunk, 1, 2, 3, true, 0);
    assert!(chunk.code.len() > 15, "any should emit bytecode");
}

#[test]
fn emit_every_produces_bytecode() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 10;
    loops::emit_any_every(&mut chunk, 1, 2, 3, false, 0);
    assert!(chunk.code.len() > 15, "every should emit bytecode");
}

#[test]
fn emit_for_in_produces_loop() {
    let mut chunk = Chunk::new("test");
    chunk.local_count = 10;
    let (loop_start, exit_jump) = loops::emit_for_in_start(&mut chunk, 1, 2, 0);
    // element is on stack here — drop it to simulate body
    chunk.emit_op(vybe_bytecode::opcode::Op::DROP, 0);
    loops::emit_for_in_end(&mut chunk, 2, loop_start, exit_jump, 0);
    assert!(chunk.code.len() > 10, "for-in should emit loop bytecode");
}
