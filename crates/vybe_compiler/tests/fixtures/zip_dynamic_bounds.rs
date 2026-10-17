use vybe_runtime::{Chunk,Op};
use vybe_compiler::primitives::{collections::{ZipLen,emit_len,emit_array_new,emit_get,emit_push},instructions::core_wasm};
pub fn emit_zip(chunks: &mut [Chunk], current: usize, n: u8, mode: ZipLen, line: u32) {
    let n = n.max(1) as u16;
    let base = chunks[current].local_count;
    chunks[current].alloc_scratch(n + 5);
    let (result_s, i_s, len_s, tuple_s, tmp_s) = (
        base + n,
        base + n + 1,
        base + n + 2,
        base + n + 3,
        base + n + 4,
    );

    // Pop the n arrays into base..base+n (top of stack = last array).
    for k in (0..n).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + k, line);
    }
    // len = arrays[0].length
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    // Shortest/Longest: fold the remaining arrays' lengths into `len`.
    if mode != ZipLen::First {
        for k in 1..n {
            chunks[current].emit_op_u16(Op::LOCAL_GET, base + k, line);
            emit_len(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, tmp_s, line);
            // if (tmp <op> len) len = tmp
            chunks[current].emit_op_u16(Op::LOCAL_GET, tmp_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            if mode == ZipLen::Shortest {
                vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
            } else {
                vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
            }
            vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, tmp_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
            chunks[current].emit_end(line);
        }
    }
    // result = []; i = 0
    emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);

    let block_p = chunks[current].emit_block(line);
    let (loop_p, _) = chunks[current].emit_loop_s(line);
    // break when !(i < len)
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    // Each row is a real tuple, not a plain list: `zip`/`enumerate` pair
    // heterogeneous values, and tuple-typed languages (Python/Dart/C#) must
    // repr it as `(a, b)`. `tuples::emit_tuple` is the one canonical builder;
    // array-typed languages (PHP) simply ignore the inert `__tuple` tag.
    let _ = tuple_s;
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    for k in 0..n {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + k, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
        emit_get(chunks, current, line);
    }
    vybe_compiler::primitives::tuples::emit_tuple(chunks, current, n as u16, line);
    // result.push(tuple)
    emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // i += 1
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_p);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_p);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
}

