use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::{Chunk, opcode::Op};

pub fn emit_range(chunks: &mut [Chunk], current: usize, arg_count: u8, inclusive: bool, line: u32) {
    let base = chunks[current].local_count;
    chunks[current].alloc_scratch(6);
    let (start_s, stop_s, step_s, i_s, result_s, isstr_s) =
        (base, base + 1, base + 2, base + 3, base + 4, base + 5);

    // Unpack stack args → (start, stop, step) by arity.
    match arg_count {
        1 => {
            // [stop] → start = 0, step = 1
            chunks[current].emit_op_u16(Op::LOCAL_SET, stop_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, step_s, line);
        }
        3 => {
            // [start, stop, step]
            chunks[current].emit_op_u16(Op::LOCAL_SET, step_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, stop_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
        }
        _ => {
            // [start, stop] (arg_count == 2) → step = 1
            chunks[current].emit_op_u16(Op::LOCAL_SET, stop_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, step_s, line);
        }
    }

    // isstr = js-string.test(start): char range → iterate over code units.
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    let str_test = chunks[current].add_import("wasm:js-string", "test");
    chunks[current].emit_call(str_test, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, isstr_s, line);

    // Char range: convert both bounds to code units.
    chunks[current].emit_op_u16(Op::LOCAL_GET, isstr_s, line);
    chunks[current].emit_if(line);
    let cca = chunks[current].add_import("ecma:string", "charCodeAt");
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_call(cca, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stop_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_call(cca, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stop_s, line);
    chunks[current].emit_end(line);

    // result = []; i = start
    chunks[current].emit_array_new_fixed(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    // cond — continue while:
    //   arity 3: step > 0 ? i < stop : i > stop   (runtime step sign)
    //   else:    i < stop   (or i <= stop when inclusive)
    if arg_count == 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, step_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line); // step > 0
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, stop_s, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, stop_s, line);
        vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, stop_s, line);
        if inclusive {
            vybe_compiler::primitives::ops::emit_dyn_le(&mut chunks[current], line);
        } else {
            vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
        }
    }
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    // result.push(isstr ? String.fromCharCode(i) : i)
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, isstr_s, line);
    chunks[current].emit_if_value(line);
    let fcc = chunks[current].add_import("ecma:string", "fromCharCode");
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    chunks[current].emit_call(fcc, 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    chunks[current].emit_end(line);
    let push = chunks[current].add_import("ecma:array", "push");
    chunks[current].emit_call(push, 2, line);
    chunks[current].emit_op(Op::DROP, line);

    // i += step
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, step_s, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
}
