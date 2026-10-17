//! Small receiver-call throughput and reference-lifetime gates.
//! Timings include VM::run setup/decode; different arities are not an A/B gain.
use crate::{Chunk, Op, Value, VM};
use crate::chunk::ReceiverAbi;
use crate::value::{Object, ObjectKind};
use std::sync::{Arc, atomic::{AtomicUsize, Ordering}};
use std::time::Instant;

fn exercise(user_count: usize, repetitions: usize) -> f64 {
    let payload = crate::heap::alloc(Object::new_array(vec![Value::I32(71)]));
    let weak = Arc::downgrade(&payload);
    let calls = Arc::new(AtomicUsize::new(0));
    let observed = calls.clone();
    let expected = payload.clone();
    let mut vm = VM::new();
    vm.register_host_fn("ecma:array", "join", Box::new(move |_, args| {
        assert_eq!(args.len(), user_count + 1);
        for arg in args {
            let Value::Object(arg) = arg else { panic!("reference changed shape") };
            assert!(Arc::ptr_eq(arg, &expected));
        }
        observed.fetch_add(1, Ordering::Relaxed);
        Value::I32(123)
    }));
    let index = vm.host_registry[&("ecma:array".to_owned(), "join".to_owned())];
    assert!(vm.host_fn_declares_receiver(index));
    let mut function = Object::new();
    function.kind = ObjectKind::HostFunction(index);
    vm.set_global("__allocation_call", Value::Object(crate::heap::alloc(function)));
    vm.set_global("__allocation_payload", Value::Object(payload.clone()));
    let mut chunk = Chunk::new("receiver-allocation-control");
    chunk.module_receiver_abi = ReceiverAbi::Parameter;
    let callee = chunk.add_constant(Value::String(Arc::from("__allocation_call")));
    let argument = chunk.add_constant(Value::String(Arc::from("__allocation_payload")));
    for _ in 0..repetitions {
        chunk.emit_op_u32(Op::GLOBAL_GET, callee, 1);
        for _ in 0..=user_count {
            chunk.emit_op_u32(Op::GLOBAL_GET, argument, 1);
        }
        chunk.emit_op_u8_u8(Op::CALL_REF, user_count as u8 + 1, 1, 1);
        chunk.emit_op(Op::DROP, 1);
    }
    chunk.emit_i32_const(42, 1);
    chunk.emit_op(Op::RETURN, 1);
    let chunks = vec![chunk];
    let start = Instant::now();
    let result = vm.run(chunks).unwrap();
    let elapsed = start.elapsed().as_secs_f64();
    assert_eq!(result, Value::I32(42));
    assert_eq!(calls.load(Ordering::Relaxed), repetitions);
    drop(vm);
    assert_eq!(Arc::strong_count(&payload), 1, "call argument references leaked");
    drop(payload);
    assert!(weak.upgrade().is_none());
    elapsed
}

#[test]
fn inline_and_scratch_receiver_calls_preserve_identity_without_retaining_args() {
    for count in 0..=5 { exercise(count, 64); }
}

#[test]
#[ignore = "receiver-call allocation control; no isolated before/after claim"]
fn loaded_receiver_argument_allocation_benchmark() {
    let mut samples: Vec<Vec<f64>> = (0..6).map(|_| Vec::new()).collect();
    for round in 0..23 {
        for step in 0..6 {
            let count = (round + step) % 6;
            let elapsed = exercise(count, 4096);
            if round >= 2 { samples[count].push(elapsed); }
        }
    }
    for (count, samples) in samples.iter_mut().enumerate() {
        samples.sort_by(f64::total_cmp);
        eprintln!("receiver-allocation users={count} calls=4096 p50_ms={:.3} p95_ms={:.3}; includes VM::run setup/decode and identity checks",
            samples[10] * 1000.0, samples[19] * 1000.0);
    }
}
