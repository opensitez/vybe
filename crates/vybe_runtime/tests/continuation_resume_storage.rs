use vybe_runtime::chunk::StackSwitchHandler;
use vybe_runtime::opcode::heaptype::HT_EXTERN;
use vybe_runtime::{Chunk, Op, Value, VM};

// Drive a live continuation through the existing resume handler machinery.
// Input: continuation. Output: yielded/completed value, followed by a flag.
fn emit_next(chunk: &mut Chunk) {
    let continuation = chunk.local_count;
    chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, continuation, 0);
    chunk.emit_block_typed(0, 2);
    chunk.emit_op_u16(Op::LOCAL_GET, continuation, 0);
    chunk.emit_ref_null(HT_EXTERN, 0);
    let resume_ip = chunk.code.len();
    chunk.emit_op_u16(Op::RESUME, 0, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_br(0, 0);
    let handler_ip = chunk.code.len();
    chunk.emit_i32_const(1, 0);
    chunk.emit_end(0);
    chunk.stack_switch_handlers.insert(
        resume_ip,
        vec![StackSwitchHandler {
            kind: 0,
            tag_index: 0,
            label_index: handler_ip as u32,
        }],
    );
}

#[test]
fn repeated_resume_preserves_saved_locals_and_loop_labels() {
    let mut body = Chunk::new("counter");
    body.arity = 1;
    body.local_count = 2;
    body.is_generator = true;
    body.emit_i32_const(0, 0);
    body.emit_op_u16(Op::LOCAL_SET, 1, 0);
    let block = body.emit_block(0);
    let (loop_patch, _) = body.emit_loop_s(0);
    body.emit_op_u16(Op::LOCAL_GET, 1, 0);
    body.emit_i32_const(3, 0);
    body.emit_op(Op::I32_LT_S, 0);
    body.emit_op(Op::I32_EQZ, 0);
    body.emit_br_if(1, 0);
    body.emit_op_u16(Op::LOCAL_GET, 1, 0);
    body.emit_op_u16(Op::SUSPEND, 0, 0);
    body.emit_op(Op::DROP, 0);
    body.emit_op_u16(Op::LOCAL_GET, 1, 0);
    body.emit_i32_const(1, 0);
    body.emit_op(Op::I32_ADD, 0);
    body.emit_op_u16(Op::LOCAL_SET, 1, 0);
    body.emit_br(0, 0);
    body.emit_end(0);
    body.patch_loop(loop_patch);
    body.emit_end(0);
    body.patch_block(block);
    body.emit_ref_null(HT_EXTERN, 0);
    body.emit_op(Op::RETURN, 0);

    let mut script = Chunk::new("<script>");
    script.local_count = 3;
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_op_u8_u8(Op::CALL_REF, 0, 1, 0);
    script.emit_op_u16(Op::LOCAL_SET, 0, 0);
    for expected in 0..3 {
        script.emit_op_u16(Op::LOCAL_GET, 0, 0);
        emit_next(&mut script);
        script.emit_op_u16(Op::LOCAL_SET, 2, 0);
        script.emit_op_u16(Op::LOCAL_SET, 1, 0);
        script.emit_op_u16(Op::LOCAL_GET, 2, 0);
        script.emit_op(Op::I32_EQZ, 0);
        script.emit_if(0);
        script.emit_op(Op::UNREACHABLE, 0);
        script.emit_end(0);
        script.emit_op_u16(Op::LOCAL_GET, 1, 0);
        script.emit_i32_const(expected, 0);
        script.emit_op(Op::I32_NE, 0);
        script.emit_if(0);
        script.emit_op(Op::UNREACHABLE, 0);
        script.emit_end(0);
    }
    script.emit_op_u16(Op::LOCAL_GET, 0, 0);
    emit_next(&mut script);
    script.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    assert_eq!(vm.run(vec![script, body]).unwrap(), Value::I32(0));
}

#[test]
fn host_reentry_does_not_complete_the_running_continuation() {
    let mut vm = VM::new();
    vm.register_host_fn(
        "test",
        "apply_callback",
        Box::new(|ctx: &mut vybe_runtime::HostContext, args: &[Value]| {
            ctx.invoke(&args[0], &[args[1].clone()])
        }),
    );
    let mut callback = Chunk::new("identity");
    callback.arity = 1;
    callback.local_count = 1;
    callback.emit_op_u16(Op::LOCAL_GET, 0, 0);
    callback.emit_op(Op::RETURN, 0);

    let mut body = Chunk::new("generator");
    body.is_generator = true;
    let apply = body.add_import("test", "apply_callback");
    body.emit_op_u16(Op::REF_FUNC, 2, 0);
    body.emit(0, 0);
    body.emit_i32_const(41, 0);
    body.emit_call(apply, 2, 0);
    body.emit_op(Op::DROP, 0);
    body.emit_i32_const(99, 0);
    body.emit_op_u16(Op::SUSPEND, 0, 0);
    body.emit_ref_null(HT_EXTERN, 0);
    body.emit_op(Op::RETURN, 0);

    let mut script = Chunk::new("<script>");
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_op_u8_u8(Op::CALL_REF, 0, 1, 0);
    emit_next(&mut script);
    script.emit_op(Op::DROP, 0);
    script.emit_op(Op::RETURN, 0);
    assert_eq!(vm.run(vec![script, body, callback]).unwrap(), Value::I32(99));
}
