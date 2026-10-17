use super::VM;
use crate::{chunk::Chunk, opcode::Op};

fn op(chunk: &mut Chunk, operation: Op) {
    chunk.code.extend_from_slice(&operation.encode());
}

fn integer(chunk: &mut Chunk, mut value: i32) {
    op(chunk, Op::I32_CONST);
    loop {
        let byte = (value & 127) as u8;
        value >>= 7;
        let done = (value == 0 && byte & 64 == 0)
            || (value == -1 && byte & 64 != 0);
        chunk.code.push(if done { byte } else { byte | 128 });
        if done { break; }
    }
}

fn local(chunk: &mut Chunk, operation: Op) {
    op(chunk, operation);
    chunk.code.extend_from_slice(&0u16.to_be_bytes());
}

fn program(count: u16, duplicates: bool) -> Chunk {
    let mut chunk = Chunk::new("script");
    chunk.local_count = 1;
    for index in 0..count {
        integer(&mut chunk, if duplicates { 0 } else { i32::from(index) });
        integer(&mut chunk, i32::from(index) + 17);
    }
    op(&mut chunk, Op::STRUCT_NEW);
    chunk.code.extend_from_slice(&0u16.to_be_bytes());
    chunk.code.extend_from_slice(&count.to_be_bytes());
    local(&mut chunk, Op::LOCAL_SET);
    for index in 0..if duplicates { count.min(1) } else { count } {
        integer(&mut chunk, 1);
        local(&mut chunk, Op::LOCAL_GET);
        integer(&mut chunk, i32::from(index));
        op(&mut chunk, Op::ARRAY_GET);
        integer(&mut chunk, i32::from(if duplicates { count - 1 } else { index }) + 17);
        op(&mut chunk, Op::I32_EQ);
        op(&mut chunk, Op::I32_DIV_S);
        op(&mut chunk, Op::DROP);
    }
    op(&mut chunk, Op::RETURN);
    chunk
}

#[test]
fn object_literals_preserve_properties_across_reservation_boundaries() {
    for count in [0, 1, 3, 4, 8, 16, 63, 64, 65, 256, 1024] {
        let mut vm = VM::new();
        vm.run(vec![program(count, false)]).unwrap();
    }
}

#[test]
fn object_literals_preserve_last_write_for_duplicate_keys() {
    for count in [1, 4, 8, 64, 65, 256, 1024] {
        let mut vm = VM::new();
        vm.run(vec![program(count, true)]).unwrap();
    }
}
