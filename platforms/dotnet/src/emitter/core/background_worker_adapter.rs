//! System.ComponentModel.BackgroundWorker: .NET delegates on the web task queue.

use vybe_compiler::primitives::{class_slots, delegates, errors, functions, ops};
use vybe_compiler::primitives::class_slots::{Dest, ObjSource, ValueSource};
use vybe_runtime::opcode::{Op, heaptype::HT_EXTERN};
use vybe_runtime::Chunk;

use super::object_fields::field_slot;

fn get(c: &mut Chunk, object: u16, key: &str, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, object, line);
    class_slots::emit_class_get(c, ObjSource::Stack, &field_slot(key), Dest::Stack, line);
}

fn set(c: &mut Chunk, object: u16, key: &str, value: u16, line: u32) {
    class_slots::emit_class_set(
        c,
        ObjSource::Local(object),
        &field_slot(key),
        ValueSource::Local(value),
        line,
    );
}

fn event_get(c: &mut Chunk, object: u16, key: &str, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, object, line);
    c.emit_string_const(key, line);
    let f = c.add_import("ecma:object", "get");
    c.emit_call(f, 2, line);
}

fn event_set(c: &mut Chunk, object: u16, key: &str, value: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, object, line);
    c.emit_string_const(key, line);
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    let f = c.add_import("ecma:object", "set");
    c.emit_call(f, 3, line);
    c.emit_op(Op::DROP, line);
}

fn set_bool(c: &mut Chunk, object: u16, key: &str, value: bool, line: u32) {
    c.emit_bool_const(value, line);
    let slot = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
    set(c, object, key, slot, line);
}

fn event(c: &mut Chunk, kind: &str, line: u32) -> u16 {
    let f = c.add_import("ecma:object", "new");
    c.emit_call(f, 0, line);
    let slot = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
    let class_name = match kind {
        "dowork" => "DoWorkEventArgs",
        "progresschanged" => "ProgressChangedEventArgs",
        _ => "RunWorkerCompletedEventArgs",
    };
    c.emit_string_const(class_name, line);
    let type_slot = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, type_slot, line);
    event_set(c, slot, "__type", type_slot, line);
    slot
}

fn invoke(c: &mut Chunk, owner: u16, key: &str, event: u16, line: u32) {
    get(c, owner, key, line);
    c.emit_op_u16(Op::LOCAL_GET, owner, line);
    c.emit_op_u16(Op::LOCAL_GET, event, line);
    delegates::emit_invoke(std::slice::from_mut(c), 0, 3, line);
    c.emit_op(Op::DROP, line);
}

fn work_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let index = chunks.len();
    let mut c = functions::create_function_chunk("__dotnet_background_worker_run", 2);
    c.module_receiver_abi = chunks[0].module_receiver_abi;
    c.local_count = 2;
    let owner = 0;
    let argument = 1;
    let work = event(&mut c, "dowork", line);
    event_set(&mut c, work, "argument", argument, line);
    c.emit_bool_const(false, line);
    let not_cancelled = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, not_cancelled, line);
    event_set(&mut c, work, "cancel", not_cancelled, line);
    c.emit_ref_null(HT_EXTERN, line);
    let error = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, error, line);
    let after_work = c.emit_block(line);
    errors::emit_try_start(&mut c, line);
    invoke(&mut c, owner, "dowork", work, line);
    errors::emit_try_end(&mut c, line);
    c.emit_br(1, line);
    errors::emit_handler_block_end(&mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, error, line);
    c.emit_end(line);
    c.patch_block(after_work);

    let completed = event(&mut c, "runworkercompleted", line);
    event_get(&mut c, work, "result", line);
    let result = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, result, line);
    event_set(&mut c, completed, "result", result, line);
    event_get(&mut c, work, "cancel", line);
    let cancelled = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, cancelled, line);
    event_set(&mut c, completed, "cancelled", cancelled, line);
    event_set(&mut c, completed, "error", error, line);
    set_bool(&mut c, owner, "isbusy", false, line);
    invoke(&mut c, owner, "runworkercompleted", completed, line);
    c.emit_ref_null(HT_EXTERN, line);
    c.emit_op(Op::RETURN, line);
    chunks.push(c);
    index
}

/// Stack: [] -> [worker]. A component has no DOM node or document lifecycle.
pub fn emit_new(c: &mut Chunk, line: u32) {
    class_slots::emit_class_construct(
        c,
        "System.ComponentModel.BackgroundWorker",
        &[
            (field_slot("isbusy"), ValueSource::ConstBool(false)),
            (field_slot("cancellationpending"), ValueSource::ConstBool(false)),
            (field_slot("workerreportsprogress"), ValueSource::ConstBool(false)),
            (field_slot("workersupportscancellation"), ValueSource::ConstBool(false)),
        ],
        line,
    );
}

/// Stack: [worker, optional argument] -> [null].
pub fn emit_run(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let c = &mut chunks[current];
    if argc == 1 {
        c.emit_ref_null(HT_EXTERN, line);
    }
    let argument = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, argument, line);
    let worker = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, worker, line);
    get(c, worker, "isbusy", line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_if(line);
    super::exceptions::emit_throw_typed(
        chunks,
        current,
        "InvalidOperationException",
        "BackgroundWorker is already running.",
        line,
    );
    let c = &mut chunks[current];
    c.emit_end(line);
    set_bool(c, worker, "isbusy", true, line);
    set_bool(c, worker, "cancellationpending", false, line);
    let index = work_chunk(chunks, line);
    let c = &mut chunks[current];
    c.emit_op_u16(Op::REF_FUNC, index as u16, line);
    c.emit(0, line);
    c.emit_ref_null(HT_EXTERN, line);
    c.emit_op_u16(Op::LOCAL_GET, worker, line);
    c.emit_op_u16(Op::LOCAL_GET, argument, line);
    let bind = c.add_import("ecma:function", "bindWithArgs");
    c.emit_call(bind, 4, line);
    c.emit_i32_const(0, line);
    let timeout = c.add_import("web:timers", "setTimeout");
    c.emit_call(timeout, 2, line);
    c.emit_op(Op::DROP, line);
    c.emit_ref_null(HT_EXTERN, line);
}

/// Stack: [worker] -> [null].
pub fn emit_cancel(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let worker = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, worker, line);
    get(c, worker, "workersupportscancellation", line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    super::exceptions::emit_throw_typed(
        chunks,
        current,
        "InvalidOperationException",
        "BackgroundWorker does not support cancellation.",
        line,
    );
    let c = &mut chunks[current];
    c.emit_end(line);
    set_bool(c, worker, "cancellationpending", true, line);
    c.emit_ref_null(HT_EXTERN, line);
}

/// Stack: [worker, percentage, optional state] -> [null].
pub fn emit_progress(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let c = &mut chunks[current];
    if argc == 2 {
        c.emit_ref_null(HT_EXTERN, line);
    }
    let state = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, state, line);
    let percentage = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, percentage, line);
    let worker = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, worker, line);
    get(c, worker, "workerreportsprogress", line);
    ops::emit_dyn_to_bool(c, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    super::exceptions::emit_throw_typed(
        chunks,
        current,
        "InvalidOperationException",
        "BackgroundWorker does not report progress.",
        line,
    );
    let c = &mut chunks[current];
    c.emit_end(line);
    let progress = event(c, "progresschanged", line);
    event_set(c, progress, "progresspercentage", percentage, line);
    event_set(c, progress, "userstate", state, line);
    get(c, worker, "progresschanged", line);
    c.emit_op_u16(Op::LOCAL_GET, worker, line);
    c.emit_op_u16(Op::LOCAL_GET, progress, line);
    delegates::emit_invoke(chunks, current, 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
}
