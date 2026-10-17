use vybe_runtime::{Chunk, Op, VM, Value};

#[test]
fn local_load_does_not_fuse_a_direct_call_as_import_zero() {
    for (module, name) in [
        ("wasm:js-undefined", "test"),
        ("wasm:js-string", "test"),
        ("wasm:js-number", "test"),
        ("wasm:js-number", "toI32"),
        ("wasm:js-number", "toF64"),
    ] {
        let mut callee = Chunk::new("actual-defined-target");
        callee.arity = 1;
        callee.local_count = 1;
        callee.emit_i32_const(77, 0);
        callee.emit_op(Op::RETURN, 0);

        let mut caller = Chunk::new("direct-call-with-aliased-import-zero");
        assert_eq!(caller.add_import(module, name), 0);
        caller.local_count = 1;
        caller.emit_i32_const(42, 0);
        caller.emit_op_u16(Op::LOCAL_SET, 0, 0);
        caller.emit_op_u16(Op::LOCAL_GET, 0, 0);
        caller.emit_direct_defined_call(1, &callee, false, 0).unwrap();
        caller.emit_op(Op::RETURN, 0);

        let result = VM::new().run(vec![caller, callee]).unwrap();
        assert!(matches!(result, Value::I32(77)), "{module}.{name}: {result:?}");
    }
}
