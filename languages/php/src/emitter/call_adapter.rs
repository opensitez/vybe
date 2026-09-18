use std::sync::Arc;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

pub fn emit_php_dynamic_method_call(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 3 {
        return;
    }
    let chunk = &mut chunks[current];
    let args_slot = chunk.alloc_scratch(1);
    let method_slot = chunk.alloc_scratch(1);
    let receiver_slot = chunk.alloc_scratch(1);
    let target_slot = chunk.alloc_scratch(1);
    let call_slot = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, args_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);

    let lookup = chunk.add_import("ecma:value", "getMethodForCall");
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_bool_const(false, line);
    chunk.emit_call(lookup, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, target_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, target_slot, line);
    let typeof_fn = chunk.add_import("ecma:value", "typeof");
    chunk.emit_call(typeof_fn, 1, line);
    chunk.emit_string_const(&Arc::<str>::from("function"), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, target_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, args_slot, line);
    let apply = chunk.add_import("ecma:function", "apply");
    chunk.emit_call(apply, 3, line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_string_const(&Arc::<str>::from("__call"), line);
    chunk.emit_bool_const(false, line);
    chunk.emit_call(lookup, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, call_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, call_slot, line);
    chunk.emit_call(typeof_fn, 1, line);
    chunk.emit_string_const(&Arc::<str>::from("function"), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, call_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, args_slot, line);
    chunk.emit_array_new_fixed(0, 2, line);
    chunk.emit_call(apply, 3, line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, target_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, args_slot, line);
    chunk.emit_call(apply, 3, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_php_method_exists(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 2 {
        return;
    }

    let chunk = &mut chunks[current];
    let method_slot = chunk.alloc_scratch(1);
    let receiver_slot = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, method_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);

    let lookup = chunk.add_import("ecma:value", "getMethodForCall");
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, method_slot, line);
    chunk.emit_bool_const(true, line);
    chunk.emit_call(lookup, 3, line);

    let typeof_fn = chunk.add_import("ecma:value", "typeof");
    chunk.emit_call(typeof_fn, 1, line);
    chunk.emit_string_const(&Arc::<str>::from("function"), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}
