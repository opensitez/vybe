use std::sync::{Arc, Mutex};
use vybe_runtime::chunk::Chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::value::{Object, Value};
use vybe_runtime::vm::VM;

fn global(chunk: &mut Chunk, name: &str) {
    let constant = chunk.add_constant(Value::String(Arc::from(name)));
    chunk.emit_op_u32(Op::GLOBAL_GET, constant, 1);
}

fn local(chunk: &mut Chunk, op: Op, slot: u16) {
    chunk.emit_op_u16(op, slot, 1);
}

fn program(reverse: bool, separate: bool, repeat: usize) -> Chunk {
    let mut chunk = Chunk::new("local-copy-assignment-dispatch");
    chunk.local_count = 2;
    chunk.scratch_high_water = 2;
    chunk.result_arity = 1;
    let source = u16::from(reverse);
    let target = 1 - source;
    global(&mut chunk, "source");
    local(&mut chunk, Op::LOCAL_SET, source);
    global(&mut chunk, "initial");
    local(&mut chunk, Op::LOCAL_SET, target);
    for _ in 0..repeat {
        local(&mut chunk, Op::LOCAL_GET, source);
        if separate { chunk.emit_op(Op::NOP, 1); }
        local(&mut chunk, Op::LOCAL_SET, target);
    }
    local(&mut chunk, Op::LOCAL_GET, target);
    chunk.emit_op(Op::RETURN, 1);
    chunk
}

fn same_value(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Null, Value::Null) | (Value::Undefined, Value::Undefined) => true,
        (Value::TypedNull(a), Value::TypedNull(b)) => a == b,
        (Value::Bool(a), Value::Bool(b)) => a == b,
        (Value::I32(a), Value::I32(b)) => a == b,
        (Value::I64(a), Value::I64(b)) => a == b,
        (Value::F32(a), Value::F32(b)) => a.to_bits() == b.to_bits(),
        (Value::F64(a), Value::F64(b)) => a.to_bits() == b.to_bits(),
        (Value::V128(a), Value::V128(b)) => a == b,
        (Value::String(a), Value::String(b))
        | (Value::Symbol(a), Value::Symbol(b)) => Arc::ptr_eq(a, b),
        (Value::Object(a), Value::Object(b)) => Arc::ptr_eq(a, b),
        (Value::WeakRef(a), Value::WeakRef(b)) => std::sync::Weak::ptr_eq(a, b),
        (Value::BigInt(a), Value::BigInt(b)) => Arc::ptr_eq(a, b),
        _ => false,
    }
}

#[test]
fn fused_local_copies_match_separate_instructions_for_all_variants() {
    let object = Arc::new(Mutex::new(Object::new()));
    let values = vec![
        Value::Null, Value::TypedNull(17), Value::Undefined, Value::Bool(true),
        Value::I32(i32::MIN), Value::I64(i64::MAX),
        Value::F32(f32::from_bits(0xff80_0042)), Value::F32(-0.0),
        Value::F64(f64::from_bits(0xfff0_0000_0000_0042)), Value::F64(-0.0),
        Value::V128(std::array::from_fn(|i| i as u8 * 17)),
        Value::String(Arc::from("local string")), Value::Symbol(Arc::from("symbol")),
        Value::Object(object.clone()), Value::WeakRef(Arc::downgrade(&object)),
        Value::bigint_i64(i64::MIN),
    ];
    for source in &values {
        for initial in &values {
            for reverse in [false, true] {
                for separate in [false, true] {
                    let mut vm = VM::new();
                    vm.set_global("source", source.clone());
                    vm.set_global("initial", initial.clone());
                    let result = vm.run(vec![program(reverse, separate, 3)]).unwrap();
                    assert!(same_value(&result, source),
                        "source={:?} initial={:?} reverse={reverse} separate={separate}",
                        source.tag(), initial.tag());
                }
            }
        }
    }
}

#[test]
fn a_fused_invalid_destination_still_traps() {
    let mut chunk = program(false, false, 0);
    // Replace the final read/return with a copy to an undeclared, absent slot.
    chunk.code.truncate(chunk.code.len() - 10);
    chunk.lines.truncate(chunk.code.len());
    local(&mut chunk, Op::LOCAL_GET, 0);
    local(&mut chunk, Op::LOCAL_SET, u16::MAX);
    chunk.emit_op(Op::RETURN, 1);
    let mut vm = VM::new();
    vm.set_global("source", Value::I32(7));
    vm.set_global("initial", Value::I32(9));
    let error = vm.run(vec![chunk]).unwrap_err();
    assert!(error.to_string().contains("local index out of bounds"), "{error}");
}

#[test]
#[ignore = "VM local-copy scenario; NOP control has an extra dispatch, not an isolated speedup"]
fn local_copy_execution_measurement() {
    use std::time::Instant;
    for (name, source) in [
        ("numeric", Value::I32(7)),
        ("object", Value::Object(Arc::new(Mutex::new(Object::new())))),
    ] {
        for separate in [false, true] {
            let chunk = program(false, separate, 8192);
            let mut samples = Vec::new();
            for sample in 0..13 {
                let mut vm = VM::new();
                vm.set_global("source", source.clone());
                vm.set_global("initial", source.clone());
                let chunks = vec![chunk.clone()];
                let start = Instant::now();
                let result = vm.run(chunks).unwrap();
                let elapsed = start.elapsed().as_secs_f64();
                assert!(same_value(&result, &source));
                if sample != 0 { samples.push(elapsed); }
            }
            samples.sort_by(f64::total_cmp);
            eprintln!("VM local-copy name={name} separate={separate} copies=8192 p50_ms={:.3} p95_ms={:.3}",
                (samples[5] + samples[6]) * 500.0, samples[11] * 1000.0);
        }
    }
}
