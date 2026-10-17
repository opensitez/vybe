use vybe_runtime::{Chunk, Op, VM};

#[path = "../src/emitter/byte_memory.rs"]
mod byte_memory;

fn kernel(destination: f64, source: f64, stride: f64, count: f64) -> Vec<Chunk> {
    let mut chunks = vec![Chunk::new("<column16>")];
    let c = &mut chunks[0];
    c.emit_i32_const(1,0);
    c.emit_op_idx(Op::MEMORY_GROW,0u32,0);
    c.emit_op(Op::DROP,0);
    c.emit_i32_const(0,0);
    c.emit_i32_const(129,0);
    c.emit_i32_const(65536,0);
    c.emit_op_idx(Op::MEMORY_FILL,0u32,0);
    for i in 0..128 {
        c.emit_i32_const(i,0);
        c.emit_i32_const((i*37+129)&255,0);
        c.emit_op(Op::I32_STORE8,0);
    }
    for value in [destination,source,stride,count] {c.emit_f64_const(value,0)}
    byte_memory::emit_copy_column_16(&mut chunks,0,0);
    chunks[0].emit_op(Op::RETURN,0);
    chunks
}

fn initial() -> Vec<u8> {
    let mut bytes=vec![129;65536];
    for (i,byte) in bytes[..128].iter_mut().enumerate() {*byte=(i*37+129) as u8}
    bytes
}

fn reference(bytes: &mut [u8], destination: i32, source: i32, stride: i32, count: u32) -> bool {
    let mut destination=destination as u32;
    let mut source=source as u32;
    for _ in 0..count {
        let from=source as usize;
        let to=destination as usize;
        if from as u64+2>bytes.len() as u64 {return true}
        let word=[bytes[from],bytes[from+1]];
        if to as u64+2>bytes.len() as u64 {return true}
        bytes[to..to+2].copy_from_slice(&word);
        source=source.wrapping_add(2);
        destination=destination.wrapping_add(stride as u32);
    }
    false
}

#[test]
fn emitter_keeps_scalar_word_order_aliases_and_trap_prefix() {
    for source in [0,1,32,65530,65534,65535,-1] {
        for destination in [0,1,31,32,33,65530,65534,65535,-1] {
            for stride in [0,1,2,3,8,-2] {
                for count in [0,1,2,8] {
                    let mut expected=initial();
                    let trapped=reference(&mut expected,destination,source,stride,count);
                    let mut vm=VM::new();
                    let result=vm.run(kernel(destination as f64,source as f64,stride as f64,count as f64));
                    assert_eq!(result.is_err(),trapped,"{destination}/{source}/{stride}/{count}");
                    match result {
                        Ok(value) => assert_eq!(value.as_f64(),destination as f64),
                        Err(error) => assert!(error.to_string().contains("out of bounds"),"{error}"),
                    }
                    vm.memory.with_buffer(|bytes| assert_eq!(bytes,&expected));
                }
            }
        }
    }
}

#[test]
fn zero_count_does_not_convert_pointers_and_large_addresses_wrap() {
    for destination in [f64::NAN,f64::INFINITY,f64::NEG_INFINITY] {
        let mut vm=VM::new();
        let result=vm.run(kernel(destination,f64::NAN,f64::INFINITY,0.)).unwrap().as_f64();
        assert!(if destination.is_nan() {result.is_nan()} else {result==destination});
        vm.memory.with_buffer(|bytes|assert_eq!(bytes,&initial()));
    }
    let mut expected=initial();
    assert!(!reference(&mut expected,2,3,3,8));
    let mut vm=VM::new();
    assert_eq!(vm.run(kernel(4294967298.,4294967299.,3.,8.)).unwrap().as_f64(),4294967298.);
    vm.memory.with_buffer(|bytes|assert_eq!(bytes,&expected));
}
