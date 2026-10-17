use vybe_runtime::{Chunk, Op, VM, Value};

#[test]
fn unaligned_wide_operands_preserve_float_bits_across_nested_units() {
    let mut vm = VM::new();
    for padding in 0..4 {
        for bits in [0u64, (-0.0f64).to_bits(), 0x7ff8_0000_0000_0042, (-123.25f64).to_bits()] {
            let mut chunk = Chunk::new("<script>");
            // Each i32.const/drop pair occupies nine bytes, exercising every
            // byte alignment of the following twelve-byte float instruction.
            for _ in 0..padding {
                chunk.emit_op(Op::I32_CONST, 0);
                chunk.code.push(0);
                chunk.lines.push(0);
                chunk.emit_op(Op::DROP, 0);
            }
            chunk.emit_op(Op::F64_CONST, 0);
            chunk.code.extend_from_slice(&bits.to_le_bytes());
            chunk.lines.extend_from_slice(&[0; 8]);
            chunk.emit_op(Op::RETURN, 0);
            let result = vm.run_linked_nested(vec![chunk], vec![]).unwrap();
            let Value::F64(value) = result else { panic!("expected float: {result:?}") };
            assert_eq!(value.to_bits(), bits);
        }
    }
}

#[test]
fn local_fusions_preserve_stack_and_slots_at_unaligned_offsets() {
    let mut vm = VM::new();
    for padding in 0..4 {
        let mut chunk = Chunk::new("<script>");
        chunk.local_count = 3;
        for value in 0..padding + 2 {
            chunk.emit_op(Op::I32_CONST, 0);
            chunk.code.push(value as u8);
            chunk.lines.push(0);
            chunk.emit_op_u16(Op::LOCAL_SET, (value % 2) as u16, 0);
        }
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, 2, 0);
        chunk.emit_op(Op::DROP, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 2, 0);
        chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        chunk.emit_op(Op::RETURN, 0);
        let expected = (padding / 2 * 2 + 1) as i32;
        let result = vm.run_linked_nested(vec![chunk], vec![]).unwrap();
        assert_eq!(result, Value::I32(expected));
    }
}

#[test]
fn native_result_fusion_preserves_values_at_unaligned_offsets() {
    let mut vm = VM::new();
    for padding in 0..4 {
        let mut chunk = Chunk::new("<script>");
        for _ in 0..padding {
            chunk.emit_op(Op::I32_CONST, 0);
            chunk.code.push(0);
            chunk.lines.push(0);
            chunk.emit_op(Op::DROP, 0);
        }
        chunk.emit_f64_const(0.0, 0);
        let import = chunk.add_import("wasm:js-number", "toI32");
        chunk.emit_call(import, 1, 0);
        chunk.emit_op(Op::I32_EQZ, 0);
        chunk.emit_op(Op::RETURN, 0);
        let target = vm.resolve_import_target("wasm:js-number", "toI32").unwrap();
        let result = vm.run_linked_nested(vec![chunk], vec![target]).unwrap();
        assert_eq!(result, Value::I32(1));
    }
}

#[test]
fn adjacent_native_conversion_calls_keep_distinct_targets() {
    let mut vm = VM::new();
    for padding in 0..4 {
        let mut chunk = Chunk::new("<script>");
        for _ in 0..padding {
            chunk.emit_i32_const(0, 0);
            chunk.emit_op(Op::DROP, 0);
        }
        chunk.emit_i32_const(7, 0);
        let from_i32 = chunk.add_import("wasm:js-number", "fromI32");
        let to_i32 = chunk.add_import("wasm:js-number", "toI32");
        for import in [from_i32, to_i32, from_i32, to_i32] {
            chunk.emit_call(import, 1, 0);
        }
        chunk.emit_op(Op::RETURN, 0);
        let imports = ["fromI32", "toI32"].into_iter()
            .map(|name| vm.resolve_import_target("wasm:js-number", name).unwrap())
            .collect();
        let result = vm.run_linked_nested(vec![chunk], imports).unwrap();
        assert_eq!(result, Value::I32(7));
    }
}

#[test]
fn function_body_uses_its_unit_imports_without_a_private_import_table() {
    let mut vm = VM::new();
    for value in [-3.0, 0.0, 42.0] {
        let mut script = Chunk::new("<script>");
        script.emit_op_u16(Op::REF_FUNC, 1, 0);
        script.code.push(0);
        script.lines.push(0);
        script.emit_f64_const(value, 0);
        script.emit_op_u8_u8(Op::CALL_REF, 1, 1, 0);
        script.emit_op(Op::RETURN, 0);
        script.add_import("wasm:js-number", "toI32");
        let mut function = Chunk::new("convert");
        function.arity = 1;
        function.param_count = 1;
        function.local_count = 1;
        function.emit_op_u16(Op::LOCAL_GET, 0, 0);
        function.emit_call(0, 1, 0);
        function.emit_op(Op::RETURN, 0);
        assert!(function.imports.is_empty());
        let target = vm.resolve_import_target("wasm:js-number", "toI32").unwrap();
        let result = vm.run_linked_nested(vec![script, function], vec![target]).unwrap();
        assert_eq!(result, Value::I32(value as i32));
    }
}
