//! Emission and execution gate for the direct call shape used by class helpers.
use vybe_runtime::{Chunk, Op, Value, VM};

fn chunks(direct: bool) -> Vec<Chunk> {
    let mut helper = Chunk::new("class-helper-add");
    helper.arity = 2;
    helper.param_count = 2;
    helper.local_count = 2;
    helper.emit_op_u16(Op::LOCAL_GET, 0, 1);
    helper.emit_op_u16(Op::LOCAL_GET, 1, 1);
    helper.emit_op(Op::I32_ADD, 1);
    helper.emit_op(Op::RETURN, 1);
    let mut root = Chunk::new("class-helper-root");
    root.emit_i32_const(5, 1);
    root.emit_i32_const(7, 1);
    if direct {
        root.emit_direct_defined_call(1, &helper, false, 1).unwrap();
    } else {
        let first = root.alloc_scratch(1);
        let second = root.alloc_scratch(1);
        root.emit_op_u16(Op::LOCAL_SET, second, 1);
        root.emit_op_u16(Op::LOCAL_SET, first, 1);
        root.emit_op_u16(Op::REF_FUNC, 1, 1);
        root.emit(0, 1);
        root.emit_op_u16(Op::LOCAL_GET, first, 1);
        root.emit_op_u16(Op::LOCAL_GET, second, 1);
        root.emit_op_u8_u8(Op::CALL_REF, 2, 1, 1);
    }
    root.emit_op(Op::RETURN, 1);
    vec![root, helper]
}

#[test]
fn direct_class_helper_call_preserves_arguments_and_reduces_emission() {
    let previous = chunks(false);
    let direct = chunks(true);
    println!("previous_root_bytes={}, direct_root_bytes={}, previous_locals={}, direct_locals={}",
        previous[0].code.len(), direct[0].code.len(),
        previous[0].local_count, direct[0].local_count);
    assert!(direct[0].code.len() < previous[0].code.len());
    assert!(direct[0].local_count < previous[0].local_count);
    assert_eq!(direct[0].direct_defined_calls.len(), 1);
    assert!(matches!(VM::new().run(previous).unwrap(), Value::I32(12)));
    assert!(matches!(VM::new().run(direct).unwrap(), Value::I32(12)));
}

#[test]
fn direct_helper_transport_preserves_reference_identity() {
    use std::sync::Arc;
    use vybe_runtime::value::Object;
    for direct in [false, true] {
        let object = vybe_runtime::heap::alloc(Object::new());
        let supplied = object.clone();
        let mut vm = VM::new();
        vm.register_host_fn("helper-fixture", "value", Box::new(move |_, _| Value::Object(supplied.clone())));
        let mut helper = Chunk::new("class-helper-reference");
        helper.arity = 1;
        helper.param_count = 1;
        helper.local_count = 1;
        helper.emit_op_u16(Op::LOCAL_GET, 0, 1);
        helper.emit_op(Op::RETURN, 1);
        let mut root = Chunk::new("reference-helper-root");
        let factory = root.add_import("helper-fixture", "value");
        root.emit_call(factory, 0, 1);
        if direct {
            root.emit_direct_defined_call(1, &helper, false, 1).unwrap();
        } else {
            let slot = root.alloc_scratch(1);
            root.emit_op_u16(Op::LOCAL_SET, slot, 1);
            root.emit_op_u16(Op::REF_FUNC, 1, 1);
            root.emit(0, 1);
            root.emit_op_u16(Op::LOCAL_GET, slot, 1);
            root.emit_op_u8_u8(Op::CALL_REF, 1, 1, 1);
        }
        root.emit_op(Op::RETURN, 1);
        let Value::Object(actual) = vm.run(vec![root, helper]).unwrap() else {
            panic!("helper did not return the reference");
        };
        assert!(Arc::ptr_eq(&actual, &object));
    }
}

fn timed_call_chunks(direct: bool, iterations: i32) -> Vec<Chunk> {
    let helper = chunks(direct).remove(1);
    let mut root = Chunk::new("timed-helper-root");
    root.local_count = if direct { 3 } else { 5 };
    let clock = root.add_import("helper-benchmark", "clock");
    root.emit_i32_const(0, 1);
    root.emit_op_u16(Op::LOCAL_SET, 0, 1);
    root.emit_call(clock, 0, 1);
    root.emit_op_u16(Op::LOCAL_SET, 2, 1);
    root.emit_loop_s(1);
    root.emit_op_u16(Op::LOCAL_GET, 0, 1);
    root.emit_i32_const(7, 1);
    if direct {
        root.emit_direct_defined_call(1, &helper, false, 1).unwrap();
    } else {
        root.emit_op_u16(Op::LOCAL_SET, 4, 1);
        root.emit_op_u16(Op::LOCAL_SET, 3, 1);
        root.emit_op_u16(Op::REF_FUNC, 1, 1);
        root.emit(0, 1);
        root.emit_op_u16(Op::LOCAL_GET, 3, 1);
        root.emit_op_u16(Op::LOCAL_GET, 4, 1);
        root.emit_op_u8_u8(Op::CALL_REF, 2, 1, 1);
    }
    root.emit_op_u16(Op::LOCAL_SET, 1, 1);
    root.emit_op_u16(Op::LOCAL_GET, 0, 1);
    root.emit_i32_const(1, 1);
    root.emit_op(Op::I32_ADD, 1);
    root.emit_op_u16(Op::LOCAL_TEE, 0, 1);
    root.emit_i32_const(iterations, 1);
    root.emit_op(Op::I32_LT_U, 1);
    root.emit_br_if(0, 1);
    root.emit_end(1);
    root.emit_call(clock, 0, 1);
    root.emit_op_u16(Op::LOCAL_GET, 2, 1);
    root.emit_op(Op::F64_SUB, 1);
    root.emit_op_u16(Op::LOCAL_GET, 1, 1);
    root.emit_i32_const(iterations + 6, 1);
    root.emit_op(Op::I32_NE, 1);
    root.emit_if(1);
    root.emit_op(Op::UNREACHABLE, 1);
    root.emit_end(1);
    root.emit_op(Op::RETURN, 1);
    vec![root, helper]
}

#[test]
#[ignore = "runtime-only call benchmark; excludes linking, predecode and source compilation"]
fn direct_class_helper_execution_benchmark() {
    use std::time::Instant;
    let iterations = 100_000;
    let run = |direct| {
        let mut vm = VM::new();
        let origin = Instant::now();
        vm.register_host_fn("helper-benchmark", "clock", Box::new(move |_, _| Value::F64(origin.elapsed().as_secs_f64())));
        let Value::F64(seconds) = vm.run(timed_call_chunks(direct, iterations)).unwrap() else {
            panic!("benchmark clock did not return elapsed seconds");
        };
        seconds
    };
    run(false);
    run(true);
    let mut previous = Vec::new();
    let mut direct = Vec::new();
    for sample in 0..7 {
        if sample % 2 == 0 {
            previous.push(run(false));
            direct.push(run(true));
        } else {
            direct.push(run(true));
            previous.push(run(false));
        }
    }
    previous.sort_by(f64::total_cmp);
    direct.sort_by(f64::total_cmp);
    println!("calls={iterations}, indirect_median_ms={}, direct_median_ms={}", previous[3] * 1000.0, direct[3] * 1000.0);
}
