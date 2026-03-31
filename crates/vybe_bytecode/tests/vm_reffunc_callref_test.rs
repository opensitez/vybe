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
    script.emit_op_u16(opcode::Op::global_get, name_c, 0);
    script.emit_op_u16(opcode::Op::r#const, v0, 0);
    script.emit_op_u16(opcode::Op::r#const, v5, 0);
    script.emit_op_u16(opcode::Op::r#const, v1, 0);
    script.emit_op_u8(opcode::Op::call_ref, 3, 0);
    // Result should be an array of length 5
    script.emit_op(opcode::Op::array_length, 0);
    script.emit_op(opcode::Op::halt, 0);

    // Chunk 1: __stdlib_range (start, stop, step) → builds array
    let mut range_chunk = Chunk::new("__stdlib_range");
    range_chunk.arity = 3;
    range_chunk.local_count = 5; // callee(0) + start(1) + stop(2) + step(3) + result(4)

    // result = []
    range_chunk.emit_op_u16(opcode::Op::array_new, 0, 0);
    range_chunk.emit_op_u16(opcode::Op::local_set, 4, 0);
    range_chunk.emit_op(opcode::Op::drop, 0);

    // while start < stop
    let loop_start = range_chunk.current_offset();
    range_chunk.emit_op_u16(opcode::Op::local_get, 1, 0);
    range_chunk.emit_op_u16(opcode::Op::local_get, 2, 0);
    range_chunk.emit_op(opcode::Op::dyn_lt, 0);
    let exit = range_chunk.emit_jump(opcode::Op::br_if_false, 0);

    range_chunk.emit_op_u16(opcode::Op::local_get, 4, 0);
    range_chunk.emit_op_u16(opcode::Op::local_get, 1, 0);
    range_chunk.emit_op(opcode::Op::array_push, 0);
    range_chunk.emit_op(opcode::Op::drop, 0);

    range_chunk.emit_op_u16(opcode::Op::local_get, 1, 0);
    range_chunk.emit_op_u16(opcode::Op::local_get, 3, 0);
    range_chunk.emit_op(opcode::Op::dyn_add, 0);
    range_chunk.emit_op_u16(opcode::Op::local_set, 1, 0);
    range_chunk.emit_op(opcode::Op::drop, 0);

    range_chunk.emit_loop(loop_start, 0);
    range_chunk.patch_jump(exit);

    range_chunk.emit_op_u16(opcode::Op::local_get, 4, 0);
    range_chunk.emit_op(opcode::Op::r#return, 0);

    let result = vm.run(vec![script, range_chunk]).unwrap();
    assert_eq!(result.as_i32(), 5, "range(0,5,1) via RefFunc global → call_ref");
}
