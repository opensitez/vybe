//! Core-Wasm loop-control kernels, without frontend/compiler dependencies.
//! Full compiler consumer gates are separate; these isolate VM branch behavior.
use vybe_runtime::{Chunk, Op, VM, Value};

fn init() -> Chunk {
    let mut chunk = Chunk::new("structured_loop_control");
    chunk.local_count = 1;
    chunk.emit_i32_const(0, 1);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 1);
    chunk
}

fn increment(chunk: &mut Chunk) {
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 1);
    chunk.emit_i32_const(1, 1);
    chunk.emit_op(Op::I32_ADD, 1);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 1);
}

fn finish(mut chunk: Chunk) -> (Value, u64, usize) {
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 1);
    chunk.emit_op(Op::RETURN, 1);
    let bytes = chunk.code.len();
    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    let result = vm.run(vec![chunk]).unwrap();
    (result, vm.take_runtime_perf().unwrap().dispatches, bytes)
}

fn header_loop(known: bool, optimized: bool, limit: i32) -> (Value, u64, usize) {
    let mut chunk = init();
    let exit = chunk.emit_block(1);
    let (again, _) = chunk.emit_loop_s(1);
    if optimized {
        if !known { chunk.emit_br(1, 1); }
    } else {
        chunk.emit_i32_const(i32::from(known), 1);
        chunk.emit_op(Op::I32_EQZ, 1);
        chunk.emit_br_if(1, 1);
    }
    increment(&mut chunk);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 1);
    chunk.emit_i32_const(limit, 1);
    chunk.emit_op(Op::I32_GE_S, 1);
    chunk.emit_br_if(1, 1);
    chunk.emit_br(0, 1);
    chunk.emit_end(1);
    chunk.patch_loop(again);
    chunk.emit_end(1);
    chunk.patch_block(exit);
    finish(chunk)
}

fn tail_loop(repeat: bool, optimized: bool, continue_body: bool, limit: i32) -> (Value, u64, usize) {
    let mut chunk = init();
    let exit = chunk.emit_block(1);
    let (again, _) = chunk.emit_loop_s(1);
    let body = chunk.emit_block(1);
    increment(&mut chunk);
    if repeat {
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 1);
        chunk.emit_i32_const(limit, 1);
        chunk.emit_op(Op::I32_GE_S, 1);
        chunk.emit_br_if(2, 1);
    }
    if continue_body {
        chunk.emit_br(0, 1);
        // A continue must skip this assignment and reach the tail condition.
        chunk.emit_i32_const(999, 1);
        chunk.emit_op_u16(Op::LOCAL_SET, 0, 1);
    }
    chunk.emit_end(1);
    chunk.patch_block(body);
    if optimized {
        if repeat { chunk.emit_br(0, 1); }
    } else {
        chunk.emit_i32_const(i32::from(repeat), 1);
        chunk.emit_br_if(0, 1);
    }
    chunk.emit_end(1);
    chunk.patch_loop(again);
    chunk.emit_end(1);
    chunk.patch_block(exit);
    finish(chunk)
}

#[test]
fn constant_header_conditions_avoid_per_iteration_tests() {
    for known in [false, true] {
        let old = header_loop(known, false, 10000);
        let new = header_loop(known, true, 10000);
        assert_eq!(old.0, new.0);
        assert_eq!(new.0, Value::I32(if known { 10000 } else { 0 }));
        assert_eq!(old.1 - new.1, if known { 30000 } else { 2 });
        assert!(new.2 < old.2);
        eprintln!("header known={known}: old={old:?} new={new:?}");
    }
}

#[test]
fn constant_tail_conditions_preserve_continue_and_break_targets() {
    for repeat in [false, true] {
        for continue_body in [false, true] {
            let old = tail_loop(repeat, false, continue_body, 10000);
            let new = tail_loop(repeat, true, continue_body, 10000);
            assert_eq!(old.0, new.0);
            assert_eq!(new.0, Value::I32(if repeat { 10000 } else { 1 }));
            assert_eq!(old.1 - new.1, if repeat { 9999 } else { 2 });
            assert!(new.2 < old.2);
            eprintln!("tail repeat={repeat} continue={continue_body}: old={old:?} new={new:?}");
        }
    }
}
