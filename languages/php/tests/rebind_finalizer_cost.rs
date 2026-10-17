//! Isolate finalizer emission cost without changing first-registration-wins hooks.
use std::time::Instant;
use vybe_runtime::{Chunk, VM, Value, opcode::Op};

fn original_finalizer(chunk: &mut Chunk, line: u32) {
    let index = chunk.add_import("php:object", "finalizeUnreachable");
    chunk.emit_call(index, 2, line);
    chunk.emit_op(Op::DROP, line);
}

fn kernel(finalizer: fn(&mut Chunk, u32), iterations: i32) -> Chunk {
    let line = 1;
    let mut chunk = Chunk::new("missing_destructor_cost");
    chunk.local_count = 1;
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, line);
    let exit = chunk.emit_block(line);
    let (repeat, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    chunk.emit_i32_const(iterations, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);
    chunk.emit_i32_const(1, line);
    // Obtain an actual Undefined from a fresh VM's absent global property.
    // This identical preparation is outside the finalizer being compared.
    chunk.emit_string_const("__finalizer_cost_absent_20261013", line);
    let get = chunk.add_import("ecma:globalThis", "get");
    chunk.emit_call(get, 1, line);
    finalizer(&mut chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(repeat);
    chunk.emit_end(line);
    chunk.patch_block(exit);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    chunk.emit_op(Op::RETURN, line);
    chunk
}

#[test]
#[ignore = "controlled performance probe; run explicitly without competing builds"]
fn compare_missing_destructor_emission() {
    vybe_language_php::register();
    let guarded = vybe_runtime::registry::hooks("php").rebind_finalizer.unwrap();
    let iterations = 100000;
    let control = kernel(original_finalizer, iterations);
    let candidate = kernel(guarded, iterations);
    for sample in 0..5 {
        let order = if sample % 2 == 0 {
            [ ("original", &control), ("guarded", &candidate) ]
        } else {
            [ ("guarded", &candidate), ("original", &control) ]
        };
        for (name, chunk) in order {
            let chunks = vec![chunk.clone()];
            let mut vm = VM::new();
            vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
            let start = Instant::now();
            let result = vm.run(chunks).unwrap();
            let elapsed = start.elapsed();
            assert!(matches!(result, Value::I32(n) if n == iterations)
                || matches!(result, Value::F64(n) if n == iterations as f64));
            println!("finalizer_cost {name} sample={sample} iterations={iterations} seconds={:.9}",
                elapsed.as_secs_f64());
        }
    }
}
