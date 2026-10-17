// Runtime-only keyed struct access benchmark.
// Timing excludes VM setup, linking, object allocation, validation, and warmup.
// This exercises the keyed object representation, not the typed GC struct path.
use std::time::Instant;
use vybe_runtime::{Chunk, Op, Value, VM};
use vybe_runtime::value::Object;

fn kernel(mode: &str, iterations: i32) -> Chunk {
    let mut c = Chunk::new("struct-benchmark");
    c.arity = 1;
    c.param_count = 1;
    c.local_count = 2;
    let key = c.add_constant(Value::String("value".into()));
    c.emit_i32_const(0, 1);
    c.emit_op_u16(Op::LOCAL_SET, 1, 1);
    c.emit_block(1);
    c.emit_loop_s(1);
    c.emit_op_u16(Op::LOCAL_GET, 1, 1);
    c.emit_i32_const(iterations, 1);
    c.emit_op(Op::I32_GE_S, 1);
    c.emit_br_if(1, 1);
    if mode != "get" {
        c.emit_op_u16(Op::LOCAL_GET, 0, 1);
        c.emit_op_u16(Op::LOCAL_GET, 1, 1);
        c.emit_struct_field_op(Op::STRUCT_SET, 0, key, 1);
    }
    if mode != "set" {
        c.emit_op_u16(Op::LOCAL_GET, 0, 1);
        c.emit_struct_field_op(Op::STRUCT_GET, 0, key, 1);
        c.emit_op(Op::DROP, 1);
    }
    c.emit_op_u16(Op::LOCAL_GET, 1, 1);
    c.emit_i32_const(1, 1);
    c.emit_op(Op::I32_ADD, 1);
    c.emit_op_u16(Op::LOCAL_SET, 1, 1);
    c.emit_br(0, 1);
    c.emit_end(1);
    c.emit_end(1);
    c.emit_op_u16(Op::LOCAL_GET, 0, 1);
    c.emit_struct_field_op(Op::STRUCT_GET, 0, key, 1);
    c.emit_op(Op::RETURN, 1);
    c
}
#[test]
#[ignore = "runtime microbenchmark; run explicitly with --ignored --nocapture"]
fn benchmark_struct_operations() {
    let iterations = 100_000;
    for mode in ["get", "set", "get-set"] {
        let mut root = Chunk::new("root");
        root.emit_op(Op::REF_FUNC, 1);
        root.code.push(1);
        root.emit_op(Op::RETURN, 1);
        let mut vm = VM::new();
        let callable = vm.run(vec![root, kernel(mode, iterations)]).unwrap();
        let mut object = Object::new();
        object.properties.insert("value".into(), Value::I32(42));
        let object = Value::Object(vybe_runtime::heap::alloc(object));
        let args = [object];
        let mut times = Vec::new();
        for _ in 0..8 {
            let start = Instant::now();
            let result = vm.invoke(&callable, &args).unwrap();
            let elapsed = start.elapsed();
            assert!(matches!(result, Value::I32(n) if n == if mode == "get" {42} else {iterations - 1}));
            times.push(elapsed.as_secs_f64() * 1000.0);
        }
        times.remove(0); // warmup excluded
        times.sort_by(f64::total_cmp);
        println!("{mode}: iterations={iterations} median_ms={:.3} samples_ms={times:?}", times[3]);
    }
}

#[test]
fn struct_operations_preserve_values() {
    for mode in ["get", "set", "get-set"] {
        let mut root = Chunk::new("root");
        root.emit_op(Op::REF_FUNC, 1);
        root.code.push(1);
        root.emit_op(Op::RETURN, 1);
        let mut vm = VM::new();
        let callable = vm.run(vec![root, kernel(mode, 128)]).unwrap();
        let mut object = Object::new();
        object.properties.insert("value".into(), Value::I32(42));
        let object = vybe_runtime::heap::alloc(object);
        let result = vm.invoke(&callable, &[Value::Object(object.clone())]).unwrap();
        let expected = if mode == "get" { 42 } else { 127 };
        assert!(matches!(result, Value::I32(n) if n == expected));
        assert!(matches!(object.lock().unwrap().properties.get("value"), Some(Value::I32(n)) if *n == expected));
    }
}
