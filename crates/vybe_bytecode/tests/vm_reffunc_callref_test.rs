use vybe_bytecode::*;
use vybe_bytecode::chunk::*;
use std::rc::Rc;

#[test]
fn global_init_reffunc_then_callref() {
    let mut vm = VM::new();
    let mut script = Chunk::new("<script>");
    script.local_count = 5;

    // Global init: __test = RefFunc(1) → Function ref to chunk 1
    script.global_inits.push(GlobalInit {
        name: "__test".to_string(),
        init: ConstExpr::RefFunc(1),
    });

    // User code: global_get "__test", push args, call_ref
    let name_c = script.add_constant(Value::String(Rc::from("__test")));
    let v0 = script.add_constant(Value::I32(0));
    let v5 = script.add_constant(Value::I32(5));
    let v1 = script.add_constant(Value::I32(1));
    script.emit_op_u16(opcode::Op::GLOBAL_GET, name_c, 0);
    script.emit_op_u16(opcode::Op::CONST, v0, 0);
    script.emit_op_u16(opcode::Op::CONST, v5, 0);
    script.emit_op_u16(opcode::Op::CONST, v1, 0);
    script.emit_op_u8(opcode::Op::CALL_REF, 3, 0);
    // Result should be an array of length 5
    script.emit_op(opcode::Op::ARRAY_LENGTH, 0);
    script.emit_op(opcode::Op::HALT, 0);

    // Chunk 1: __stdlib_range (start, stop, step) → builds array
    let mut range_chunk = Chunk::new("__stdlib_range");
    range_chunk.arity = 3;
    range_chunk.local_count = 5; // callee(0) + start(1) + stop(2) + step(3) + result(4)

    // result = []
    range_chunk.emit_op_u16(opcode::Op::ARRAY_NEW_FIXED, 0, 0);
    range_chunk.emit_op_u16(opcode::Op::LOCAL_SET, 4, 0);
    range_chunk.emit_op(opcode::Op::DROP, 0);

    // while start < stop
    let loop_start = range_chunk.current_offset();
    range_chunk.emit_op_u16(opcode::Op::LOCAL_GET, 1, 0);
    range_chunk.emit_op_u16(opcode::Op::LOCAL_GET, 2, 0);
    range_chunk.emit_op(opcode::Op::DYN_LT, 0);
    let exit = range_chunk.emit_jump(opcode::Op::BR_IF_FALSE, 0);

    range_chunk.emit_op_u16(opcode::Op::LOCAL_GET, 4, 0);
    range_chunk.emit_op_u16(opcode::Op::LOCAL_GET, 1, 0);
    range_chunk.emit_op(opcode::Op::ARRAY_PUSH, 0);
    range_chunk.emit_op(opcode::Op::DROP, 0);

    range_chunk.emit_op_u16(opcode::Op::LOCAL_GET, 1, 0);
    range_chunk.emit_op_u16(opcode::Op::LOCAL_GET, 3, 0);
    range_chunk.emit_op(opcode::Op::DYN_ADD, 0);
    range_chunk.emit_op_u16(opcode::Op::LOCAL_SET, 1, 0);
    range_chunk.emit_op(opcode::Op::DROP, 0);

    range_chunk.emit_loop(loop_start, 0);
    range_chunk.patch_jump(exit);

    range_chunk.emit_op_u16(opcode::Op::LOCAL_GET, 4, 0);
    range_chunk.emit_op(opcode::Op::RETURN, 0);

    let result = vm.run(vec![script, range_chunk]).unwrap();
    assert_eq!(result.as_i32(), 5, "range(0,5,1) via RefFunc global → call_ref");
}
