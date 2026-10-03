//! Smoke tests for the `emitter::loops` helpers. Each helper takes
//! `(chunks: &mut [Chunk], current: usize, ...slot args, line: u32)` so
//! the imports it registers land on `chunks[0]` (module-level WASM
//! convention). These tests drive them through a one-element slice.

use vybe_compiler::primitives::loops;
use vybe_runtime::Chunk;

#[test]
fn for_in_numeric_condition_handles_empty_arrays_and_live_length() {
    use vybe_runtime::{Value, VM};
    use vybe_runtime::opcode::Op;
    for (values, grow, expected) in [
        (vec![], false, 0),
        (vec![1, 2, 3], false, 6),
        (vec![1], true, 3),
    ] {
        let mut chunks = one_chunk(3);
        vybe_compiler::primitives::globals::emit_read(&mut chunks[0], "items", 0);
        chunks[0].emit_op_u16(Op::LOCAL_SET, 1, 0);
        chunks[0].emit_i32_const(0, 0);
        chunks[0].emit_op_u16(Op::LOCAL_SET, 0, 0);
        let state = loops::emit_for_in_start(&mut chunks, 0, 1, 2, 0);
        chunks[0].emit_op_u16(Op::LOCAL_GET, 0, 0);
        chunks[0].emit_op(Op::I32_ADD, 0);
        chunks[0].emit_op_u16(Op::LOCAL_SET, 0, 0);
        if grow {
            chunks[0].emit_op_u16(Op::LOCAL_GET, 2, 0);
            chunks[0].emit_op(Op::I32_EQZ, 0);
            chunks[0].emit_if(0);
            chunks[0].emit_op_u16(Op::LOCAL_GET, 1, 0);
            chunks[0].emit_i32_const(2, 0);
            vybe_compiler::primitives::collections::emit_push(&mut chunks, 0, 0);
            chunks[0].emit_op(Op::DROP, 0);
            chunks[0].emit_end(0);
        }
        loops::emit_for_in_end(&mut chunks, 0, 2, state, 0);
        chunks[0].emit_op_u16(Op::LOCAL_GET, 0, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        let mut vm = VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        vm.set_global_owned("items".to_owned(), Value::Object(vybe_runtime::heap::alloc(
            vybe_runtime::object::Object::new_array(values.into_iter().map(Value::I32).collect()),
        )));
        assert_eq!(vm.run(chunks).expect("for-in execution").as_i32(), expected);
    }
}

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
    loops::emit_filter(
        &mut chunks,
        0,
        vybe_runtime::chunk::ReceiverAbi::Ambient,
        1,
        2,
        3,
        4,
        5,
        0,
    );
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
    loops::emit_any_every(&mut chunks, 0, 1, 2, 3, 4, true, 0);
    assert!(chunks[0].code.len() > 15, "any should emit bytecode");
}

#[test]
fn emit_every_produces_bytecode() {
    let mut chunks = one_chunk(10);
    loops::emit_any_every(&mut chunks, 0, 1, 2, 3, 4, false, 0);
    assert!(chunks[0].code.len() > 15, "every should emit bytecode");
}

#[test]
fn emit_for_in_produces_loop() {
    let mut chunks = one_chunk(10);
    let state = loops::emit_for_in_start(&mut chunks, 0, 1, 2, 0);
    // element is on stack here — drop it to simulate body
    chunks[0].emit_op(vybe_runtime::opcode::Op::DROP, 0);
    loops::emit_for_in_end(&mut chunks, 0, 2, state, 0);
    assert!(
        chunks[0].code.len() > 10,
        "for-in should emit loop bytecode"
    );
}
