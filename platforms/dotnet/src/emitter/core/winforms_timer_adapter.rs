//! WinForms Timer scheduling over the shared web event loop.

use vybe_compiler::primitives::{functions, ops};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::{Op, heaptype::HT_EXTERN};

fn get(c: &mut Chunk, owner: u16, key: &str, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, owner, line);
    c.emit_string_const(key, line);
    let f = c.add_import("ecma:object", "get");
    c.emit_call(f, 2, line);
}

fn set(c: &mut Chunk, owner: u16, key: &str, value: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, owner, line);
    c.emit_string_const(key, line);
    c.emit_op_u16(Op::LOCAL_GET, value, line);
    let f = c.add_import("ecma:object", "set");
    c.emit_call(f, 3, line);
    c.emit_op(Op::DROP, line);
}

fn tick_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let index = chunks.len();
    let mut tick = functions::create_function_chunk("__dotnet_winforms_timer_tick", 0);
    tick.capture_base = 0;
    tick.capture_count = 1;
    tick.local_count = 1;
    let doc = tick.add_import("web:html", "activeDocument");
    tick.emit_call(doc, 0, line);
    tick.emit_op_u16(Op::LOCAL_GET, 0, line);
    tick.emit_string_const("tick", line);
    let event = tick.add_import("web:ui-events", "newEvent");
    tick.emit_call(event, 1, line);
    let dispatch = tick.add_import("web:dom", "dispatchEvent");
    tick.emit_call(dispatch, 3, line);
    tick.emit_op(Op::DROP, line);
    tick.emit_ref_null(HT_EXTERN, line);
    tick.emit_op(Op::RETURN, line);
    chunks.push(tick);
    index
}

fn stop(c: &mut Chunk, owner: u16, line: u32) {
    get(c, owner, "__timer_id", line);
    let id = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, id, line);
    c.emit_op_u16(Op::LOCAL_GET, id, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, id, line);
    let clear = c.add_import("web:timers", "clearInterval");
    c.emit_call(clear, 1, line);
    c.emit_op(Op::DROP, line);
    c.emit_end(line);
    c.emit_ref_null(HT_EXTERN, line);
    let empty = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, empty, line);
    set(c, owner, "__timer_id", empty, line);
}

fn start(chunks: &mut Vec<Chunk>, current: usize, owner: u16, line: u32) {
    let index = tick_chunk(chunks, line);
    let c = &mut chunks[current];
    c.emit_op_u16(Op::REF_FUNC, index as u16, line);
    c.emit(1, line);
    functions::emit_closure_upvalue(c, true, owner, line);
    get(c, owner, "interval", line);
    let interval = c.add_import("web:timers", "setInterval");
    c.emit_call(interval, 2, line);
    let id = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, id, line);
    set(c, owner, "__timer_id", id, line);
}

/// Stack: [timer] -> [null].
pub fn emit_init(c: &mut Chunk, line: u32) {
    let owner = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, owner, line);
    c.emit_bool_const(false, line);
    let enabled = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, enabled, line);
    set(c, owner, "enabled", enabled, line);
    c.emit_i32_const(100, line);
    let interval = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, interval, line);
    set(c, owner, "interval", interval, line);
    c.emit_ref_null(HT_EXTERN, line);
}

/// Stack: [timer, enabled] -> [null].
pub fn emit_enabled_set(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let c = &mut chunks[current];
    ops::emit_dyn_to_bool(c, line);
    let enabled = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, enabled, line);
    let owner = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, owner, line);
    stop(c, owner, line);
    c.emit_op_u16(Op::LOCAL_GET, enabled, line);
    ops::emit_i32_to_bool(c, line);
    let enabled_value = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, enabled_value, line);
    set(c, owner, "enabled", enabled_value, line);
    c.emit_op_u16(Op::LOCAL_GET, enabled, line);
    c.emit_if(line);
    start(chunks, current, owner, line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
}

/// Stack: [timer, interval] -> [null].
pub fn emit_interval_set(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let c = &mut chunks[current];
    let interval = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, interval, line);
    let owner = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, owner, line);
    get(c, owner, "enabled", line);
    let enabled = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, enabled, line);
    stop(c, owner, line);
    set(c, owner, "interval", interval, line);
    c.emit_op_u16(Op::LOCAL_GET, enabled, line);
    c.emit_if(line);
    start(chunks, current, owner, line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
}

/// Stack: [timer] -> [null].
pub fn emit_stop(c: &mut Chunk, line: u32) {
    let owner = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, owner, line);
    stop(c, owner, line);
    c.emit_bool_const(false, line);
    let enabled = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, enabled, line);
    set(c, owner, "enabled", enabled, line);
    c.emit_ref_null(HT_EXTERN, line);
}

/// Stack: [timer] -> [null].
pub fn emit_start(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    chunks[current].emit_bool_const(true, line);
    emit_enabled_set(chunks, current, line);
}
