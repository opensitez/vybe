use vybe_runtime::{Chunk, Op, VM};

#[test]
fn host_callback_does_not_swallow_vm_errors() {
    let mut vm = VM::new();
    vm.register_host_fn("test", "invoke", Box::new(|ctx, args| ctx.invoke(&args[0], &[])));
    let mut main = Chunk::new("main");
    main.emit_op_u16(Op::REF_FUNC, 1, 1);
    main.emit(0u8, 1);
    let invoke = main.add_import("test", "invoke");
    main.emit_call(invoke, 1, 1);
    main.emit_op(Op::RETURN, 1);
    let mut callback = Chunk::new("callback");
    callback.emit_op(Op::UNREACHABLE, 2);
    let error = vm.run(vec![main, callback]).expect_err("callback failure must escape its host caller");
    assert!(error.to_string().contains("unreachable"), "{error}");
}
