#[path = "../src/emitter/heap_memory.rs"]
mod heap_memory;

use vybe_runtime::{Chunk, Op, VM, Value};

#[test]
fn shrink_keeps_address_prefix_and_updates_requested_size() {
    for size in [1, 7, 16, 64, 127, 128] {
        let mut seed = vec![0xa5; 65_536];
        seed[4088..4092].copy_from_slice(&128u32.to_le_bytes());
        let mut expected = seed.clone();
        expected[4088..4092].copy_from_slice(&(size as u32).to_le_bytes());
        let mut vm = VM::new();
        vm.memory.with_buffer_mut(|memory| *memory = seed);
        let mut chunks = vec![Chunk::new("realloc-in-place")];
        chunks[0].emit_f64_const(4096.0, 0);
        chunks[0].emit_f64_const(size as f64, 0);
        heap_memory::emit_realloc(&mut chunks, 0, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        assert_eq!(vm.run(chunks).unwrap(), Value::F64(4096.0));
        assert_eq!(vm.memory.with_buffer(|memory| memory.to_vec()), expected);
    }
}

#[test]
fn zero_and_invalid_sizes_do_not_read_or_mutate_the_old_block() {
    for size in [0.0, -1.0, f64::NAN, f64::INFINITY, 4294967295.0] {
        let seed = vec![0xa5; 65_536];
        let mut vm = VM::new();
        vm.memory.with_buffer_mut(|memory| *memory = seed.clone());
        let mut chunks = vec![Chunk::new("realloc-invalid-size")];
        chunks[0].emit_f64_const(4294967295.0, 0);
        chunks[0].emit_f64_const(size, 0);
        heap_memory::emit_realloc(&mut chunks, 0, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        assert_eq!(vm.run(chunks).unwrap(), Value::I32(0));
        assert_eq!(vm.memory.with_buffer(|memory| memory.to_vec()), seed);
    }
}

#[test]
fn zero_size_bypasses_nonfinite_old_pointer_conversion() {
    for old in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        let seed = vec![0xa5; 65_536];
        let mut vm = VM::new();
        vm.memory.with_buffer_mut(|memory| *memory = seed.clone());
        let mut chunks = vec![Chunk::new("realloc-zero-nonfinite-pointer")];
        chunks[0].emit_f64_const(old, 0);
        chunks[0].emit_f64_const(0.0, 0);
        heap_memory::emit_realloc(&mut chunks, 0, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        assert_eq!(vm.run(chunks).unwrap(), Value::I32(0));
        assert_eq!(vm.memory.with_buffer(|memory| memory.to_vec()), seed);
    }
}

#[test]
fn growth_and_null_allocation_preserve_headers_and_distinct_prefix_bytes() {
    for old in [0, 4096] {
        let mut seed = vec![0xa5; 65_536];
        seed[4088..4092].copy_from_slice(&128u32.to_le_bytes());
        for i in 0..128 {
            seed[4096+i] = (i*37+129) as u8;
        }
        let mut expected = seed.clone();
        expected[16384..16388].copy_from_slice(&256u32.to_le_bytes());
        if old != 0 {
            expected[16392..16520].copy_from_slice(&seed[4096..4224]);
        }
        let mut vm = VM::new();
        vm.memory.with_buffer_mut(|memory| *memory = seed);
        let mut chunks = vec![Chunk::new("realloc-growth-and-null")];
        chunks[0].emit_i32_const(16384, 0);
        vybe_compiler::primitives::globals::emit_write(&mut chunks[0], "__c_heap_ptr", 0);
        chunks[0].emit_f64_const(old as f64, 0);
        chunks[0].emit_f64_const(256.0, 0);
        heap_memory::emit_realloc(&mut chunks, 0, 0);
        chunks[0].emit_op(Op::RETURN, 0);
        vybe_compiler::primitives::globals::finalize_global_table(&mut chunks);
        assert_eq!(vm.run(chunks).unwrap(), Value::I32(16392));
        assert_eq!(vm.memory.with_buffer(|memory| memory.to_vec()), expected);
    }
}
