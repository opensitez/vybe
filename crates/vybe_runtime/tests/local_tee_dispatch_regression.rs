use std::sync::{Arc, Mutex, Weak};
use vybe_runtime::chunk::Chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::vm::VM;

fn global(chunk: &mut Chunk, name: &str) {
    let index = chunk.add_constant(Value::String(Arc::from(name)));
    chunk.emit_op_u32(Op::GLOBAL_GET, index, 1);
}

fn program(slot: u16, repetitions: usize) -> Chunk {
    let mut chunk = Chunk::new("local-tee-dispatch");
    chunk.local_count = 1;
    chunk.scratch_high_water = 1;
    chunk.result_arity = 1;
    global(&mut chunk, "initial");
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 1);
    global(&mut chunk, "source");
    for _ in 0..repetitions { chunk.emit_op_u16(Op::LOCAL_TEE, slot, 1); }
    // Return both the retained operand and a fresh read of the assigned local.
    chunk.emit_op_u16(Op::LOCAL_GET, slot, 1);
    chunk.emit_array_new_fixed(0, 2, 1);
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
        (Value::WeakRef(a), Value::WeakRef(b)) => Weak::ptr_eq(a, b),
        (Value::BigInt(a), Value::BigInt(b)) => Arc::ptr_eq(a, b),
        _ => false,
    }
}

fn check_result(result: &Value, source: &Value) {
    let Value::Object(object) = result else { panic!("expected pair array"); };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else { panic!("expected pair array"); };
    assert_eq!(values.len(), 2);
    assert!(same_value(&values[0], source), "tee changed its operand");
    assert!(same_value(&values[1], source), "tee did not update its local");
}

#[test]
fn tee_retains_operand_and_updates_local_for_every_value_variant() {
    let object = Arc::new(Mutex::new(Object::new()));
    let values = vec![
        Value::Null, Value::TypedNull(17), Value::Undefined, Value::Bool(true),
        Value::I32(i32::MIN), Value::I64(i64::MAX),
        Value::F32(f32::from_bits(0xff80_0042)), Value::F32(-0.0),
        Value::F64(f64::from_bits(0xfff0_0000_0000_0042)), Value::F64(-0.0),
        Value::V128(std::array::from_fn(|i| i as u8 * 17)),
        Value::String(Arc::from("tee string")), Value::Symbol(Arc::from("symbol")),
        Value::Object(object.clone()), Value::WeakRef(Arc::downgrade(&object)),
        Value::bigint_i64(i64::MIN),
    ];
    for source in &values {
        for initial in &values {
            for repetitions in [1, 4] {
                let mut vm = VM::new();
                vm.set_global("source", source.clone());
                vm.set_global("initial", initial.clone());
                check_result(&vm.run(vec![program(0, repetitions)]).unwrap(), source);
            }
        }
    }
}

#[test]
fn tee_to_an_absent_local_still_traps() {
    let mut vm = VM::new();
    vm.set_global("source", Value::I32(7));
    vm.set_global("initial", Value::I32(9));
    let error = vm.run(vec![program(u16::MAX, 1)]).unwrap_err();
    assert!(error.to_string().contains("local index out of bounds"), "{error}");
}

#[test]
#[ignore = "VM tee execution measurement; includes decode/setup, not isolated kernel time"]
fn local_tee_execution_measurement() {
    use std::time::Instant;
    for (name, source) in [
        ("numeric", Value::I32(7)),
        ("object", Value::Object(Arc::new(Mutex::new(Object::new())))),
    ] {
        let chunk = program(0, 8192);
        let mut samples = Vec::new();
        for sample in 0..21 {
            let mut vm = VM::new();
            vm.set_global("source", source.clone());
            vm.set_global("initial", source.clone());
            let chunks = vec![chunk.clone()];
            let start = Instant::now();
            let result = vm.run(chunks).unwrap();
            let elapsed = start.elapsed().as_secs_f64();
            check_result(&result, &source);
            if sample != 0 { samples.push(elapsed); }
        }
        samples.sort_by(f64::total_cmp);
        eprintln!("VM local-tee name={name} assignments=8192 p50_ms={:.6} p95_ms={:.6}",
            (samples[9] + samples[10]) * 500.0, samples[18] * 1000.0);
    }
}
