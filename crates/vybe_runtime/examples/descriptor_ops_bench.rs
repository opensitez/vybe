//! Compiler-free descriptor instruction microbenchmark.
//!
//! Run with `cargo run -p vybe_runtime --release --example descriptor_ops_bench
//! -- 100000`. This measures interpreter execution plus VM/link initialization,
//! not PHP compilation or a full application. The loop-only case exposes the
//! shared overhead; do not treat microbenchmark ratios as WordPress speedups.
use std::time::Instant;
use vybe_runtime::chunk::TypeEntry;
use vybe_runtime::{Chunk, Op, VM, Value};

#[derive(Clone, Copy, Debug)]
enum Work {
    LoopOnly,
    GetDescriptor,
    CastDescriptor,
    AllocateDescriptorInstance,
}

fn program(iterations: i32, work: Work) -> Chunk {
    assert!(iterations > 0);
    let mut chunk = Chunk::new("script");
    chunk.local_count = 3;
    chunk.types = vec![
        TypeEntry { name: "BenchInstance".into(), ..Default::default() },
        TypeEntry {
            name: "#desc BenchInstance".into(),
            describes_index: 1,
            ..Default::default()
        },
    ];
    // One descriptor reference shared by every instance. No guest property
    // map, class-name comparison, host import, or PHP opcode is involved.
    chunk.emit_struct_new(2, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_struct_new_desc(1, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 1, 0);
    chunk.emit_i32_const(iterations, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 2, 0);
    chunk.emit_loop_s(0);
    match work {
        Work::LoopOnly => {}
        Work::GetDescriptor => {
            chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
            chunk.emit_op_u16(Op::REF_GET_DESC, 1, 0);
            chunk.emit_op(Op::DROP, 0);
        }
        Work::CastDescriptor => {
            chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
            chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
            chunk.emit_op_u16(Op::REF_CAST_DESC_EQ, 1, 0);
            chunk.emit_op(Op::DROP, 0);
        }
        Work::AllocateDescriptorInstance => {
            chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
            chunk.emit_op_u16(Op::STRUCT_NEW_DEFAULT_DESC, 1, 0);
            chunk.emit_op(Op::DROP, 0);
        }
    }
    chunk.emit_op_u16(Op::LOCAL_GET, 2, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_SUB, 0);
    chunk.emit_op_u16(Op::LOCAL_TEE, 2, 0);
    chunk.emit_br_if(0, 0);
    chunk.emit_end(0);
    chunk.emit_op_u16(Op::LOCAL_GET, 2, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk
}

fn main() {
    let iterations = std::env::args().nth(1)
        .map(|arg| arg.parse::<i32>().expect("iterations must be an i32"))
        .unwrap_or(100_000);
    assert!(iterations > 0, "iterations must be positive");
    println!("round,operation,iterations,ir_bytes,elapsed_ns,ns_per_iteration");
    let cases = [Work::LoopOnly, Work::GetDescriptor, Work::CastDescriptor,
        Work::AllocateDescriptorInstance];
    for work in cases {
        assert!(matches!(VM::new().run(vec![program(100, work)]).unwrap(), Value::I32(0)));
    }
    for round in 0..5 {
        // Alternate order to make consistent order bias visible in raw data.
        for index in 0..cases.len() {
            let work = cases[if round % 2 == 0 { index } else { cases.len() - 1 - index }];
            let chunk = program(iterations, work);
            let bytes = chunk.code.len();
            let mut vm = VM::new();
            let start = Instant::now();
            let result = vm.run(vec![chunk]).expect("descriptor workload must execute");
            let elapsed = start.elapsed().as_nanos();
            assert!(matches!(result, Value::I32(0)));
            println!("{round},{work:?},{iterations},{bytes},{elapsed},{:.3}",
                elapsed as f64 / f64::from(iterations));
        }
    }
}
