//! Small-code, steady-state local assignment controls.
//! VM::run setup is timed, but amortized over repeated execution of one loop.
use std::hint::black_box;
use std::sync::Arc;
use std::time::Instant;
use vybe_runtime::{Chunk, Op, Value, VM};
use vybe_runtime::value::Object;

#[derive(Clone, Copy, Debug)]
enum Work { Copy, SelfCopy, TeeDrop }

fn local(chunk: &mut Chunk, op: Op, slot: u16) {
    chunk.emit_op(op, 1);
    chunk.code.extend_from_slice(&slot.to_be_bytes());
    chunk.lines.resize(chunk.code.len(), 1);
}

fn workload(work: Work, rounds: usize) -> Chunk {
    assert!(rounds > 0 && rounds <= i32::MAX as usize);
    let mut chunk = Chunk::new("steady-local-assignment");
    chunk.local_count = 8;
    chunk.scratch_high_water = 8;
    chunk.result_arity = 1;
    let key = chunk.add_constant(Value::String(Arc::from("__steady_seed")));
    chunk.emit_op_u32(Op::GLOBAL_GET, key, 1);
    local(&mut chunk, Op::LOCAL_SET, 0);
    chunk.emit_i32_const(rounds as i32, 1);
    local(&mut chunk, Op::LOCAL_SET, 2);
    chunk.emit_op_u8_u8(Op::LOOP, 0, 0, 1);
    let target = if matches!(work, Work::SelfCopy) { 0 } else { 1 };
    for _ in 0..32 {
        local(&mut chunk, Op::LOCAL_GET, 0);
        if matches!(work, Work::TeeDrop) {
            local(&mut chunk, Op::LOCAL_TEE, target);
            chunk.emit_op(Op::DROP, 1);
        } else {
            local(&mut chunk, Op::LOCAL_SET, target);
        }
    }
    local(&mut chunk, Op::LOCAL_GET, 2);
    chunk.emit_i32_const(1, 1);
    chunk.emit_op(Op::I32_SUB, 1);
    local(&mut chunk, Op::LOCAL_TEE, 2);
    chunk.emit_op(Op::BR_IF, 1);
    chunk.code.push(0); // depth zero: repeat the innermost loop
    chunk.lines.resize(chunk.code.len(), 1);
    chunk.emit_op(Op::END, 1);
    // A prematurely exited loop must fail, rather than produce a plausible seed.
    local(&mut chunk, Op::LOCAL_GET, 2);
    chunk.emit_op_u8_u8(Op::IF, 0, 0, 1);
    chunk.emit_op(Op::UNREACHABLE, 1);
    chunk.emit_op(Op::END, 1);
    local(&mut chunk, Op::LOCAL_GET, target);
    chunk.emit_op(Op::RETURN, 1);
    chunk
}

fn execute(seed: &Value, chunk: Chunk) -> f64 {
    let before_count = match seed {
        Value::Object(value) => Some(Arc::strong_count(value)),
        Value::String(value) => Some(Arc::strong_count(value)),
        _ => None,
    };
    let mut vm = VM::new();
    vm.set_global("__steady_seed", seed.clone());
    let chunks = vec![chunk];
    let start = Instant::now();
    let result = vm.run(black_box(chunks)).expect("steady loop must complete");
    let elapsed = start.elapsed().as_secs_f64();
    match (&result, seed) {
        (Value::Object(actual), Value::Object(expected)) => assert!(Arc::ptr_eq(actual, expected)),
        (Value::String(actual), Value::String(expected)) => assert!(Arc::ptr_eq(actual, expected)),
        _ => assert_eq!(&result, seed),
    }
    drop(result);
    drop(vm);
    match seed {
        Value::Object(value) => assert_eq!(Some(Arc::strong_count(value)), before_count),
        Value::String(value) => assert_eq!(Some(Arc::strong_count(value)), before_count),
        _ => {}
    }
    elapsed
}

fn seeds() -> Vec<(&'static str, Value)> {
    vec![
        ("i32", Value::I32(71)),
        ("v128", Value::V128([0x31; 16])),
        ("string", Value::String(Arc::from("steady-assignment"))),
        ("object", Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![Value::I32(71)])))),
    ]
}

#[test]
fn steady_local_assignments_preserve_values_identity_and_reference_lifetimes() {
    for (_, seed) in seeds() {
        for work in [Work::Copy, Work::SelfCopy, Work::TeeDrop] {
            for rounds in [1, 2, 17] { execute(&seed, workload(work, rounds)); }
        }
    }
}

#[test]
#[ignore = "steady-state assignment baseline; not a before/after speedup claim"]
fn loaded_steady_assignment_benchmark() {
    let seeds = seeds();
    let works = [Work::Copy, Work::SelfCopy, Work::TeeDrop];
    let chunks: Vec<_> = works.iter().map(|&work| workload(work, 8192)).collect();
    let count = seeds.len() * works.len();
    let mut samples: Vec<Vec<f64>> = (0..count).map(|_| Vec::new()).collect();
    for round in 0..23 {
        for step in 0..count {
            let index = (round + step) % count;
            let elapsed = execute(&seeds[index / 3].1, chunks[index % 3].clone());
            if round >= 2 { samples[index].push(elapsed); }
        }
    }
    for (index, samples) in samples.iter_mut().enumerate() {
        samples.sort_by(f64::total_cmp);
        eprintln!("steady-assignment type={} work={:?} assignments=262144 code_bytes={} p50_ms={:.3} p95_ms={:.3}",
            seeds[index / 3].0, works[index % 3], chunks[index % 3].code.len(),
            samples[10] * 1000.0, samples[19] * 1000.0);
    }
}
