use std::sync::Arc;

use vybe_compiler::primitives::{callable, functions::create_function_chunk, instructions::host};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};
use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames,
};

const PY_THREAD_HANDLE_KEY: &str = "__py_thread_handle";
const PY_THREAD_TARGET_KEY: &str = "__py_thread_target";

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn set_object_prop_from_local(
    chunks: &mut [Chunk],
    current: usize,
    object: u16,
    key: &str,
    value: u16,
    line: u32,
) {
    get(&mut chunks[current], object, line);
    chunks[current].emit_string_const(key, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn get_object_prop(chunks: &mut [Chunk], current: usize, object: u16, key: &str, line: u32) {
    get(&mut chunks[current], object, line);
    chunks[current].emit_string_const(key, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
}

pub fn emit_thread_start_with(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let target = chunks[current].alloc_scratch(1);
    let thread = chunks[current].alloc_scratch(1);
    let task = chunks[current].alloc_scratch(1);

    set(&mut chunks[current], target, line);
    set(&mut chunks[current], thread, line);

    set_object_prop_from_local(chunks, current, thread, PY_THREAD_TARGET_KEY, target, line);

    let mut worker = create_function_chunk("__py_thread_worker", 1);
    let target_key = class_slots::resolve_interned(&mut worker, &ClassSlot::internal(PY_THREAD_TARGET_KEY), &PlainNames);
    worker.emit_op_u16(Op::LOCAL_GET, 0, line);
    worker.emit_op_u16(Op::TABLE_GET, 0, line);
    worker.emit_op_u16(Op::LOCAL_SET, 0, line);
    worker.emit_op_u16(Op::LOCAL_GET, 0, line);
    class_slots::emit_class_get(&mut worker, ObjSource::Stack, &target_key, Dest::Stack, line);
    callable::emit_direct_invoke_chunk(&mut worker, 0, line);
    worker.emit_op(Op::RETURN, line);
    worker.local_count = 1;
    chunks.push(worker);
    let worker_idx = chunks.len() - 1;

    get(&mut chunks[current], thread, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::TABLE_GROW, 0, line);
    chunks[current].emit_op_u16(Op::REF_FUNC, worker_idx as u16, line);
    chunks[current].emit(0, line);
    vybe_compiler::primitives::threading::emit_thread_spawn(chunks, current, line);
    set(&mut chunks[current], task, line);

    set_object_prop_from_local(chunks, current, thread, PY_THREAD_HANDLE_KEY, task, line);

    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_thread_join(chunks: &mut [Chunk], current: usize, line: u32) {
    let thread = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], thread, line);
    get_object_prop(chunks, current, thread, PY_THREAD_HANDLE_KEY, line);
    vybe_compiler::primitives::threading::emit_thread_join(&mut chunks[current], line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}
