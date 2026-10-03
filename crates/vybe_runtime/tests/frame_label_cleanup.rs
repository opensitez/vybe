use vybe_runtime::{Chunk, Op, VM};

#[test]
fn callee_fallthrough_discards_its_control_labels() {
    let mut callee = Chunk::new("fallthrough");
    // A linked function may end without an explicit RETURN. Its open label
    // must not become the caller's branch destination.
    callee.emit_loop_s(0);

    let mut main = Chunk::new("<main>");
    let block = main.emit_block(0);
    main.emit_op_u16(Op::REF_FUNC, 1, 0);
    main.emit(0u8, 0); // capture count
    main.emit_op_u8_u8(Op::CALL_REF, 0, 1, 0);
    main.emit_op(Op::DROP, 0);
    main.emit_br(0, 0);
    main.emit_op(Op::UNREACHABLE, 0);
    main.emit_end(0);
    main.patch_block(block);
    main.emit_i32_const(42, 0);
    main.emit_op(Op::RETURN, 0);

    let result = VM::new().run(vec![main, callee]).expect("call and branch should complete");
    assert_eq!(result.as_i32(), 42);
}
