use std::time::Instant;
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, Value, VM};

#[test]
#[ignore = "execution-only microbenchmark; run explicitly with --ignored --nocapture"]
fn benchmark_existing_object_property_set() {
    const ITERATIONS: i32 = 100_000;
    for adapter in [false, true] {
    let mut samples = Vec::new();
    for sample in 0..8 {
        let mut object = Object::new();
        object.properties.insert("value".into(), Value::I32(0));
        let object = vybe_runtime::heap::alloc(object);
        let shared = object.clone();
        let mut vm = VM::new();
        vybe_platform_ecma::object::register(&mut vm);
        vm.register_host_fn("fixture", "noop_set", Box::new(|_, _| Value::Null));
        vm.register_host_fn("fixture", "object", Box::new(move |_, _| Value::Object(shared.clone())));
        vm.register_host_fn("fixture", "key", Box::new(|_, _| Value::String("value".into())));
        let epoch = Instant::now();
        vm.register_host_fn("fixture", "now", Box::new(move |_, _| Value::F64(epoch.elapsed().as_secs_f64())));
        let mut c = Chunk::new("object-set-speed");
        c.local_count = 4;
        let make_object = c.add_import("fixture", "object");
        let make_key = c.add_import("fixture", "key");
        let now = c.add_import("fixture", "now");
        let set = if adapter {
            c.add_import("ecma:object", "set")
        } else {
            c.add_import("fixture", "noop_set")
        };
        c.emit_call(make_object, 0, 1);
        c.emit_op_u16(Op::LOCAL_SET, 0, 1);
        c.emit_call(make_key, 0, 1);
        c.emit_op_u16(Op::LOCAL_SET, 1, 1);
        c.emit_i32_const(0, 1);
        c.emit_op_u16(Op::LOCAL_SET, 2, 1);
        c.emit_call(now, 0, 1);
        c.emit_op_u16(Op::LOCAL_SET, 3, 1);
        c.emit_block(1);
        c.emit_loop_s(1);
        c.emit_op_u16(Op::LOCAL_GET, 2, 1);
        c.emit_i32_const(ITERATIONS, 1);
        c.emit_op(Op::I32_GE_S, 1);
        c.emit_br_if(1, 1);
        c.emit_op_u16(Op::LOCAL_GET, 0, 1);
        c.emit_op_u16(Op::LOCAL_GET, 1, 1);
        c.emit_op_u16(Op::LOCAL_GET, 2, 1);
        c.emit_i32_const(0, 1);
        c.emit_call(set, 4, 1);
        c.emit_op(Op::DROP, 1);
        c.emit_op_u16(Op::LOCAL_GET, 2, 1);
        c.emit_i32_const(1, 1);
        c.emit_op(Op::I32_ADD, 1);
        c.emit_op_u16(Op::LOCAL_SET, 2, 1);
        c.emit_br(0, 1);
        c.emit_end(1);
        c.emit_end(1);
        c.emit_call(now, 0, 1);
        c.emit_op_u16(Op::LOCAL_GET, 3, 1);
        c.emit_op(Op::F64_SUB, 1);
        c.emit_op(Op::RETURN, 1);
        let elapsed = vm.run(vec![c]).unwrap();
        let Value::F64(seconds) = elapsed else { panic!("timer did not return f64") };
        let expected = if adapter { ITERATIONS - 1 } else { 0 };
        assert!(matches!(object.lock().unwrap().properties.get("value"), Some(Value::I32(n)) if *n == expected));
        if sample != 0 { samples.push(seconds * 1000.0); }
    }
    samples.sort_by(f64::total_cmp);
    let label = if adapter { "existing object.set" } else { "no-op host call control" };
    println!("{label}: {ITERATIONS} calls, median {:.3} ms; samples {samples:?}", samples[3]);
    }
    if vybe_platform_ecma::perf::enabled() {
        println!("{}", vybe_platform_ecma::perf::report());
    }
}
