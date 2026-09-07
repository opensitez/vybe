use vybe_compiler::primitives::collections;
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use crate::emitter::adapter_util::{lget, stash_args};
use crate::emitter::collections_adapter::emit_py_iter_array;

pub fn emit_list_iadd(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 2 {
        return;
    }
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let src = base + 1;
    lget(&mut chunks[current], recv, line);
    core_wasm::dup(&mut chunks[current], line);
    collections::emit_len(chunks, current, line);
    lget(&mut chunks[current], src, line);
    emit_py_iter_array(chunks, current, 1, line);
    collections::emit_insert_range(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], recv, line);
}
