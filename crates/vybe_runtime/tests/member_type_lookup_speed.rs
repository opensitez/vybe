use std::time::Instant;
use vybe_runtime::value::Object;
use vybe_runtime::{Chunk, Op, Value, VM};

#[test]
#[ignore = "runtime-only benchmark; run explicitly with --ignored --nocapture"]
fn benchmark_missing_member_with_type_metadata() {
    const ITERATIONS: i32 = 100_000;
    for metadata in [None, Some("object"), Some("symfony.component.httpfoundation.request"), Some("İ.Σ.中文")] {
        let mut samples = Vec::new();
        for sample in 0..8 {
            let mut object = Object::new();
            if let Some(name) = metadata {
                object.properties.insert("__type".into(), Value::String(name.into()));
            }
            let object = vybe_runtime::heap::alloc(object);
            let shared = object.clone();
            let mut vm = VM::new();
            vm.register_host_fn("fixture", "object", Box::new(move |_, _| Value::Object(shared.clone())));
            let epoch = Instant::now();
            vm.register_host_fn("fixture", "now", Box::new(move |_, _| Value::F64(epoch.elapsed().as_secs_f64())));
            vm.register_host_fn("fixture", "check", Box::new(|_, args| {
                assert!(matches!(args.get(1), Some(Value::Undefined)));
                args[0].clone()
            }));
            let mut c = Chunk::new("member-type-lookup-speed");
            c.local_count = 3;
            let key = c.add_constant(Value::String("missing".into()));
            let make_object = c.add_import("fixture", "object");
            let now = c.add_import("fixture", "now");
            let check = c.add_import("fixture", "check");
            c.emit_call(make_object, 0, 1);
            c.emit_op_u16(Op::LOCAL_SET, 0, 1);
            c.emit_i32_const(0, 1);
            c.emit_op_u16(Op::LOCAL_SET, 1, 1);
            c.emit_call(now, 0, 1);
            c.emit_op_u16(Op::LOCAL_SET, 2, 1);
            c.emit_block(1);
            c.emit_loop_s(1);
            c.emit_op_u16(Op::LOCAL_GET, 1, 1);
            c.emit_i32_const(ITERATIONS, 1);
            c.emit_op(Op::I32_GE_S, 1);
            c.emit_br_if(1, 1);
            c.emit_op_u16(Op::LOCAL_GET, 0, 1);
            c.emit_struct_field_op(Op::STRUCT_GET, 0, key, 1);
            c.emit_op(Op::DROP, 1);
            c.emit_op_u16(Op::LOCAL_GET, 1, 1);
            c.emit_i32_const(1, 1);
            c.emit_op(Op::I32_ADD, 1);
            c.emit_op_u16(Op::LOCAL_SET, 1, 1);
            c.emit_br(0, 1);
            c.emit_end(1);
            c.emit_end(1);
            c.emit_call(now, 0, 1);
            c.emit_op_u16(Op::LOCAL_GET, 2, 1);
            c.emit_op(Op::F64_SUB, 1);
            c.emit_op_u16(Op::LOCAL_GET, 0, 1);
            c.emit_struct_field_op(Op::STRUCT_GET, 0, key, 1);
            c.emit_call(check, 2, 1);
            c.emit_op(Op::RETURN, 1);
            let Value::F64(seconds) = vm.run(vec![c]).unwrap() else { panic!("timer result was not f64") };
            if sample != 0 { samples.push(seconds * 1000.0); }
        }
        samples.sort_by(f64::total_cmp);
        println!("metadata={metadata:?} reads={ITERATIONS} median_ms={:.3} samples_ms={samples:?}", samples[3]);
    }
}
