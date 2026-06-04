//! PHP autoloading — Rust inline opcode emitters.
//!
//! When a class constructor global is `undefined` at runtime, PHP invokes
//! the registered `spl_autoload_register` callback (stored in the
//! `__php_autoload_callback` / `__php_autoload_callback_receiver` globals)
//! with the class name, then re-reads the constructor global. These
//! adapters emit that fallback sequence straight into the chunk.
//!
//! Mirrors the other `languages/php/emitter` adapters: chunk-based, core
//! ops only. The shared compiler routes here via the `supports_autoload`
//! profile flag — no `profile.name == "php"` branch.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use crate::emitter::ops::{emit_dyn_eq, emit_dyn_to_bool};

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let s = chunk.local_count;
    chunk.local_count = s + 1;
    s
}
fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    let idx = chunk.add_constant(val);
    chunk.emit_op_u16(Op::CONST, idx, line);
}
fn push_str(chunk: &mut Chunk, v: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(v)), line);
}
fn str_idx(chunk: &mut Chunk, v: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(v)))
}

/// Push a reference to `ctor_global`, autoloading the class first if the
/// global is still undefined. Stack on exit: `[ctor_ref]`.
pub fn emit_constructor_ref_with_autoload(
    chunk: &mut Chunk,
    ctor_global: &str,
    autoload_name: &str,
    line: u32,
) {
    let idx = str_idx(chunk, ctor_global);
    let ctor_slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::GLOBAL_GET, idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ctor_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, ctor_slot, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    push_str(chunk, "undefined", line);
    emit_dyn_eq(chunk, line);
    emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    emit_autoload_invoke(chunk, autoload_name, line);

    chunk.emit_op_u16(Op::GLOBAL_GET, idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ctor_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ctor_slot, line);
}

/// Like [`emit_constructor_ref_with_autoload`] but resolves a primary
/// constructor global, then an optional fallback global, before
/// autoloading. Stack on exit: `[ctor_ref]`.
pub fn emit_dynamic_constructor_ref_with_autoload(
    chunk: &mut Chunk,
    primary_ctor_global: &str,
    fallback_ctor_global: Option<&str>,
    autoload_name: &str,
    line: u32,
) {
    let ctor_slot = alloc_local(chunk);
    let primary_idx = str_idx(chunk, primary_ctor_global);
    chunk.emit_op_u16(Op::GLOBAL_GET, primary_idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ctor_slot, line);
    chunk.emit_op(Op::DROP, line);

    if let Some(fallback) = fallback_ctor_global {
        emit_fallback_if_undefined(chunk, ctor_slot, fallback, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, ctor_slot, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    push_str(chunk, "undefined", line);
    emit_dyn_eq(chunk, line);
    emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    emit_autoload_invoke(chunk, autoload_name, line);

    chunk.emit_op_u16(Op::GLOBAL_GET, primary_idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ctor_slot, line);
    chunk.emit_op(Op::DROP, line);
    if let Some(fallback) = fallback_ctor_global {
        emit_fallback_if_undefined(chunk, ctor_slot, fallback, line);
    }

    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ctor_slot, line);
}

/// `if ctor_slot is undefined { ctor_slot = GLOBAL_GET fallback }`.
fn emit_fallback_if_undefined(chunk: &mut Chunk, ctor_slot: u16, fallback: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, ctor_slot, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    push_str(chunk, "undefined", line);
    emit_dyn_eq(chunk, line);
    emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    let fallback_idx = str_idx(chunk, fallback);
    chunk.emit_op_u16(Op::GLOBAL_GET, fallback_idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ctor_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
}

/// Invoke the registered autoload callback with `autoload_name` (passing
/// the receiver as `this` when the callback is a method). No-op when no
/// callback is registered.
fn emit_autoload_invoke(chunk: &mut Chunk, autoload_name: &str, line: u32) {
    let autoload_slot = alloc_local(chunk);
    let autoload_idx = str_idx(chunk, "__php_autoload_callback");
    chunk.emit_op_u16(Op::GLOBAL_GET, autoload_idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, autoload_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, autoload_slot, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    push_str(chunk, "undefined", line);
    emit_dyn_eq(chunk, line);
    emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);

    let receiver_slot = alloc_local(chunk);
    let receiver_idx = str_idx(chunk, "__php_autoload_callback_receiver");
    chunk.emit_op_u16(Op::GLOBAL_GET, receiver_idx, line);
    chunk.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    push_str(chunk, "undefined", line);
    emit_dyn_eq(chunk, line);
    emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    // Plain function callback: call with just the class name.
    chunk.emit_op_u16(Op::LOCAL_GET, autoload_slot, line);
    push_str(chunk, autoload_name, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_else(line);
    // Method callback: call with (receiver, class name).
    chunk.emit_op_u16(Op::LOCAL_GET, autoload_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    push_str(chunk, autoload_name, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}
