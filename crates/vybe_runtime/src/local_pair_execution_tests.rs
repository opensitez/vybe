use super::VM;
use crate::{chunk::Chunk, opcode::Op};

fn op(chunk: &mut Chunk, operation: Op) {
    chunk.code.extend_from_slice(&operation.encode());
}

fn constant(chunk: &mut Chunk, mut value: i32) {
    op(chunk, Op::I32_CONST);
    loop {
        let byte = (value & 127) as u8;
        value >>= 7;
        let done = (value == 0 && byte & 64 == 0) || (value == -1 && byte & 64 != 0);
        chunk.code.push(if done { byte } else { byte | 128 });
        if done { break; }
    }
}

fn local(chunk: &mut Chunk, operation: Op, slot: u16) {
    op(chunk, operation);
    chunk.code.extend_from_slice(&slot.to_be_bytes());
}

fn program(alias: bool, write: bool) -> Chunk {
    let mut chunk = Chunk::new("script");
    chunk.local_count = 2;
    constant(&mut chunk, 17);
    local(&mut chunk, Op::LOCAL_SET, 0);
    constant(&mut chunk, 20);
    local(&mut chunk, Op::LOCAL_SET, 1);
    // Retain the numerator below the read pair. A wrong result becomes a
    // zero denominator, turning the guest assertion into a standard trap.
    constant(&mut chunk, 1);
    local(&mut chunk, Op::LOCAL_GET, 0);
    local(&mut chunk, Op::LOCAL_GET, if alias { 0 } else { 1 });
    if write {
        local(&mut chunk, Op::LOCAL_SET, 0);
        local(&mut chunk, Op::LOCAL_GET, 0);
        op(&mut chunk, Op::I32_SUB);
        constant(&mut chunk, -3);
    } else {
        op(&mut chunk, Op::I32_ADD);
        constant(&mut chunk, if alias { 34 } else { 37 });
    }
    op(&mut chunk, Op::I32_EQ);
    op(&mut chunk, Op::I32_DIV_S);
    op(&mut chunk, Op::RETURN);
    chunk
}

#[test]
fn consecutive_local_reads_preserve_values_and_retained_operand() {
    for alias in [false, true] {
        for history in [false, true] {
            let mut vm = VM::new();
            vm.record_error_context(history);
            vm.run(vec![program(alias, false)]).expect("local read-pair guest assertion");
        }
    }
}

#[test]
fn existing_get_get_set_priority_preserves_read_before_write() {
    for history in [false, true] {
        let mut vm = VM::new();
        vm.record_error_context(history);
        vm.run(vec![program(false, true)]).expect("read-before-write guest assertion");
    }
}

#[test]
fn counters_include_both_local_reads() {
    let mut vm = VM::new();
    vm.record_runtime_perf(true);
    vm.run(vec![program(false, false)]).expect("counter guest assertion");
    assert_eq!(vm.take_runtime_perf().unwrap().local_get, 2);
}
