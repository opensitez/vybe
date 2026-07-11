//! PHP error-handler runtime — `set_error_handler` / `trigger_error` /
//! `error_get_last` / `restore_error_handler` / `error_clear_last` /
//! `set_exception_handler` / `error_reporting`.
//!
//! "PHP over JS": the handler is a JS-style callback stored in a global and
//! invoked via `CALL_REF`; error state (`error_get_last`) lives in globals too.
//! Handlers form a linked stack via a `prev` field so `restore_*` pops cleanly.
//! Control flow uses the shared structured `emit_if_value`/`emit_else`/`emit_end`
//! helpers rather than hand-managed jumps.

use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

/// Current error handler: `{cb, mask, prev}` object, or null.
const HANDLER_G: &str = "__php_err_handler";
/// Last error array (`{type, message, file, line}`), or null.
const LAST_G: &str = "__php_last_error";
/// Current uncaught-exception handler callback, or null.
const EXC_G: &str = "__php_exc_handler";
/// Active `error_reporting()` bitmask (defaults to `E_ALL` when unset).
const REPORTING_G: &str = "__php_err_reporting";

const E_USER_NOTICE: f64 = 1024.0;
const E_ALL: f64 = 32767.0;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    chunk.emit_string_const(value, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    chunks[current].emit_call(idx, argc, line);
}

fn struct_get_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_GET, idx, line);
}

fn struct_set_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_SET, idx, line);
    chunk.emit_op(Op::DROP, line);
}

fn global_get(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::GLOBAL_GET, idx, line);
}

fn global_set(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::GLOBAL_SET, idx, line);
}

/// Set one string→value entry on the map in `map_slot` (`ecma:map.set` returns
/// the map, which we drop — the Map is mutated in place).
fn map_set_slot(
    chunks: &mut [Chunk],
    current: usize,
    map_slot: u16,
    key: &str,
    val_slot: u16,
    line: u32,
) {
    {
        let chunk = &mut chunks[current];
        lget(chunk, map_slot, line);
        push_str(chunk, key, line);
        lget(chunk, val_slot, line);
    }
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Build a PHP error array `{type, message, file, line}` and store it in a
/// fresh slot, returning that slot. `type_slot`/`msg_slot` supply the values.
fn emit_error_map(
    chunks: &mut [Chunk],
    current: usize,
    type_slot: u16,
    msg_slot: u16,
    line: u32,
) -> u16 {
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let map_slot = {
        let chunk = &mut chunks[current];
        let s = alloc_local(chunk);
        lset(chunk, s, line);
        s
    };
    map_set_slot(chunks, current, map_slot, "type", type_slot, line);
    map_set_slot(chunks, current, map_slot, "message", msg_slot, line);
    // file/line: constant placeholders that still satisfy shape probes.
    let file_slot = {
        let chunk = &mut chunks[current];
        let s = alloc_local(chunk);
        push_str(chunk, "php", line);
        lset(chunk, s, line);
        s
    };
    map_set_slot(chunks, current, map_slot, "file", file_slot, line);
    let line_slot = {
        let chunk = &mut chunks[current];
        let s = alloc_local(chunk);
        chunk.emit_f64_const(1.0, line);
        lset(chunk, s, line);
        s
    };
    map_set_slot(chunks, current, map_slot, "line", line_slot, line);
    map_slot
}

/// PHP `set_error_handler($cb [, $mask])` — push a new handler frame.
/// Stack: `[cb]` or `[cb, mask]` → `[prev_cb|null]`.
pub fn emit_set_error_handler(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let mask_slot = alloc_local(chunk);
    if argc >= 2 {
        lset(chunk, mask_slot, line);
    } else {
        chunk.emit_f64_const(E_ALL, line);
        lset(chunk, mask_slot, line);
    }
    let cb_slot = alloc_local(chunk);
    lset(chunk, cb_slot, line);

    // prev = current handler (for the return value)
    let prev_slot = alloc_local(chunk);
    global_get(chunk, HANDLER_G, line);
    lset(chunk, prev_slot, line);

    // handler = {cb, mask, prev}
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_dup(line);
    lget(chunk, cb_slot, line);
    struct_set_key(chunk, "cb", line);
    chunk.emit_dup(line);
    lget(chunk, mask_slot, line);
    struct_set_key(chunk, "mask", line);
    chunk.emit_dup(line);
    lget(chunk, prev_slot, line);
    struct_set_key(chunk, "prev", line);
    global_set(chunk, HANDLER_G, line);

    // return the previous handler's callback (or null).
    lget(chunk, prev_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_op(Op::NULL, line);
    chunk.emit_else(line);
    lget(chunk, prev_slot, line);
    struct_get_key(chunk, "cb", line);
    chunk.emit_end(line);
}

/// PHP `restore_error_handler()` — pop the current handler frame. → `true`.
pub fn emit_restore_error_handler(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let cur_slot = alloc_local(chunk);
    global_get(chunk, HANDLER_G, line);
    lset(chunk, cur_slot, line);
    // if current != null: handler = current.prev
    lget(chunk, cur_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    lget(chunk, cur_slot, line);
    struct_get_key(chunk, "prev", line);
    global_set(chunk, HANDLER_G, line);
    chunk.emit_end(line);
    chunk.emit_bool_const(true, line);
}

/// PHP `error_get_last()` → the last error array or null.
pub fn emit_error_get_last(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    global_get(chunk, LAST_G, line);
    // Coerce an unset (undefined) global to null.
    let v = alloc_local(chunk);
    lset(chunk, v, line);
    lget(chunk, v, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_op(Op::NULL, line);
    chunk.emit_else(line);
    lget(chunk, v, line);
    chunk.emit_end(line);
}

/// PHP `error_clear_last()` — reset the last error to null.
pub fn emit_error_clear_last(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::NULL, line);
    global_set(chunk, LAST_G, line);
    chunk.emit_op(Op::NULL, line);
}

/// Push the active reporting mask, treating an unset global as `E_ALL`.
fn emit_get_reporting(chunk: &mut Chunk, line: u32) {
    let cur = alloc_local(chunk);
    global_get(chunk, REPORTING_G, line);
    lset(chunk, cur, line);
    lget(chunk, cur, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(E_ALL, line);
    chunk.emit_else(line);
    lget(chunk, cur, line);
    chunk.emit_end(line);
}

/// PHP `error_reporting([$level])` — get, or set-and-return-previous, the mask.
pub fn emit_error_reporting(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let new_slot = if argc >= 1 {
        let s = alloc_local(chunk);
        lset(chunk, s, line);
        Some(s)
    } else {
        None
    };
    // old = current mask (default E_ALL)
    let old_slot = alloc_local(chunk);
    emit_get_reporting(chunk, line);
    lset(chunk, old_slot, line);
    if let Some(s) = new_slot {
        lget(chunk, s, line);
        global_set(chunk, REPORTING_G, line);
    }
    lget(chunk, old_slot, line);
}

/// PHP `set_exception_handler($cb)` → previous handler (or null).
pub fn emit_set_exception_handler(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let cb_slot = alloc_local(chunk);
    lset(chunk, cb_slot, line);
    let prev_slot = alloc_local(chunk);
    global_get(chunk, EXC_G, line);
    lset(chunk, prev_slot, line);
    lget(chunk, cb_slot, line);
    global_set(chunk, EXC_G, line);
    lget(chunk, prev_slot, line);
}

/// PHP `trigger_error($msg [, $level])` — dispatch to the user handler when its
/// mask matches; otherwise (or when the handler returns false) record the error
/// for `error_get_last`. → `true`.
pub fn emit_trigger_error(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (level_slot, msg_slot) = {
        let chunk = &mut chunks[current];
        let level = alloc_local(chunk);
        if argc >= 2 {
            lset(chunk, level, line);
        } else {
            chunk.emit_f64_const(E_USER_NOTICE, line);
            lset(chunk, level, line);
        }
        let msg = alloc_local(chunk);
        lset(chunk, msg, line);
        (level, msg)
    };

    let map_slot = emit_error_map(chunks, current, level_slot, msg_slot, line);
    let handler_slot = {
        let chunk = &mut chunks[current];
        let s = alloc_local(chunk);
        global_get(chunk, HANDLER_G, line);
        lset(chunk, s, line);
        s
    };

    let chunk = &mut chunks[current];
    // Reporting gate: skip entirely when `(level & error_reporting()) == 0`.
    lget(chunk, level_slot, line);
    emit_get_reporting(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);

    // if handler is set …
    lget(chunk, handler_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);

    //   … and (level & handler.mask) != 0 …
    lget(chunk, level_slot, line);
    lget(chunk, handler_slot, line);
    struct_get_key(chunk, "mask", line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);

    //     invoke cb(errno, errstr, errfile, errline) → ret
    lget(chunk, handler_slot, line);
    struct_get_key(chunk, "cb", line);
    lget(chunk, level_slot, line);
    lget(chunk, msg_slot, line);
    push_str(chunk, "php", line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op_u8(Op::CALL_REF, 4, line);
    let ret_slot = alloc_local(chunk);
    lset(chunk, ret_slot, line);
    //     if ret === false → record for error_get_last
    lget(chunk, ret_slot, line);
    chunk.emit_bool_const(false, line);
    crate::emitter::ops::emit_js_strict_eq(chunk, line);
    chunk.emit_if(line);
    lget(chunk, map_slot, line);
    global_set(chunk, LAST_G, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    //   handler mask excludes the level → default: record it.
    lget(chunk, map_slot, line);
    global_set(chunk, LAST_G, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    // no handler → default: record it.
    lget(chunk, map_slot, line);
    global_set(chunk, LAST_G, line);
    chunk.emit_end(line);

    chunk.emit_end(line); // end reporting gate
    chunk.emit_bool_const(true, line);
}
