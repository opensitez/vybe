//! PHP Fiber helpers — Rust inline opcode emitters using the VM's
//! existing stack-switching primitives (`CONT_NEW`, `RESUME`, `SUSPEND`).
//!
//! PHP 8.1 `Fiber` is a synchronous stack-switching coroutine — the
//! exact shape WASM continuation opcodes were designed for. Mapping:
//!
//!   `new Fiber($cb)`     → `CONT_NEW $cb`
//!   `Fiber::suspend(v)`  → `SUSPEND v`
//!   `$f->start(args...)` → `RESUME $f val`     (binds args via __bound_args)
//!   `$f->resume(v)`      → `RESUME $f v`
//!   `$f->getReturn()`    → read from continuation's saved value
//!
//! All four PHP forms compose to the same continuation Object the
//! generator infrastructure already uses, so this is bytecode-level
//! reuse — no new VM ops, no new host fns.

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ValueSource,
};
use vybe_runtime::chunk::StackSwitchHandler;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

/// `new Fiber($cb)` — wrap the callback as a continuation Object via
/// `CONT_NEW`. Stack on entry: `[$cb]`. Stack on exit: `[continuation]`.
pub fn emit_php_fiber_new(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    chunks[current].emit_op(Op::CONT_NEW, line);
}

/// `Fiber::suspend($v)` — yield from the current continuation, returning
/// `$v` to the caller's `RESUME`. Stack on entry: `[$v]`. Stack on
/// exit: `[resume_value]` (whatever the next RESUME passes back).
pub fn emit_php_fiber_suspend(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        // PHP `Fiber::suspend()` with no arg yields null.
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    let tag = 0;
    vybe_compiler::primitives::generators::emit_suspend_tagged(chunk, tag, line);
}

/// `$fiber->start($v)` — start the fiber with `$v` as the initial
/// value. Stack on entry: `[$fiber, $v?]`. Stack on exit: `[yielded_value]`.
///
/// PHP `start()` may receive multiple args; for the test surface we
/// support up to 1 explicit arg (the rest are passed via the closure's
/// param defaults). When called with 0 args, push null first.
pub fn emit_php_fiber_start(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    // argc counts user args. After `[$fiber, args...]` we need
    // `[$fiber, val]` for RESUME (arity-2). Pad with null when called
    // with no args.
    if argc == 1 {
        // Just `[$fiber]` — push null as the resume value.
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    } else if argc > 2 {
        // Multi-arg start: drop extras (rest pattern would require
        // bound-args wiring — pending tests).
        for _ in 2..argc {
            chunk.emit_op(Op::DROP, line);
        }
    }
    let value_slot = alloc_local(chunk);
    let fiber_slot = alloc_local(chunk);
    let ret_slot = alloc_local(chunk);
    let started_key = class_slots::resolve_interned(chunk, &ClassSlot::internal("__started"), &PlainNames);
    let suspended_key = class_slots::resolve_interned(chunk, &ClassSlot::internal("__suspended"), &PlainNames);
    let running_key = class_slots::resolve_interned(chunk, &ClassSlot::internal("__running"), &PlainNames);
    let terminated_key = class_slots::resolve_interned(chunk, &ClassSlot::internal("__terminated"), &PlainNames);
    let return_key = class_slots::resolve_interned(chunk, &ClassSlot::internal("__return"), &PlainNames);

    lset(chunk, value_slot, line); // [$fiber]
    chunk.emit_op_u16(Op::LOCAL_TEE, fiber_slot, line); // [$fiber]
    chunk.emit_bool_const(true, line); // [$fiber, true]
    class_slots::emit_class_set(chunk, ObjSource::Stack, &started_key, ValueSource::Stack, line); // [true]
    lget(chunk, fiber_slot, line);
    chunk.emit_bool_const(false, line);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &suspended_key, ValueSource::Stack, line);
    lget(chunk, fiber_slot, line);
    chunk.emit_bool_const(false, line);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &terminated_key, ValueSource::Stack, line);
    lget(chunk, fiber_slot, line);
    chunk.emit_bool_const(true, line);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &running_key, ValueSource::Stack, line);
    lget(chunk, fiber_slot, line);
    vybe_compiler::primitives::globals::emit_write(chunk, "__php_current_fiber", line);

    let block_p = chunk.emit_block_typed(line, 1);
    lget(chunk, fiber_slot, line);
    lget(chunk, value_slot, line);
    let tag = 0;
    let resume_ip = chunk.code.len();
    vybe_compiler::primitives::generators::emit_resume_tagged(chunk, tag, line);

    // Completion arm: RESUME fell through because the fiber returned.
    chunk.emit_op_u16(Op::LOCAL_TEE, ret_slot, line); // [ret]
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    vybe_compiler::primitives::globals::emit_write(chunk, "__php_current_fiber", line);
    lget(chunk, fiber_slot, line);
    chunk.emit_bool_const(false, line);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &running_key, ValueSource::Stack, line);
    lget(chunk, fiber_slot, line); // [ret, fiber]
    chunk.emit_bool_const(false, line); // [ret, fiber, false]
    class_slots::emit_class_set(chunk, ObjSource::Stack, &suspended_key, ValueSource::Stack, line); // [ret, true]
    lget(chunk, fiber_slot, line); // [ret, fiber]
    chunk.emit_bool_const(true, line); // [ret, fiber, true]
    class_slots::emit_class_set(chunk, ObjSource::Stack, &terminated_key, ValueSource::Stack, line); // [ret, true]
    lget(chunk, fiber_slot, line); // [ret, fiber]
    lget(chunk, ret_slot, line); // [ret, fiber, ret]
    class_slots::emit_class_set(chunk, ObjSource::Stack, &return_key, ValueSource::Stack, line); // [ret, ret]
    chunk.emit_br(0, line);

    // Yield arm: VM jumps here from SUSPEND with [yielded_value].
    let handler_ip = chunk.code.len();
    chunk.emit_op_u16(Op::LOCAL_TEE, ret_slot, line); // [yielded]
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    vybe_compiler::primitives::globals::emit_write(chunk, "__php_current_fiber", line);
    lget(chunk, fiber_slot, line);
    chunk.emit_bool_const(false, line);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &running_key, ValueSource::Stack, line);
    lget(chunk, fiber_slot, line);
    chunk.emit_bool_const(true, line);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &suspended_key, ValueSource::Stack, line);
    lget(chunk, fiber_slot, line);
    chunk.emit_bool_const(false, line);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &terminated_key, ValueSource::Stack, line);
    lget(chunk, fiber_slot, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &return_key, ValueSource::Stack, line);
    chunk.emit_end(line);
    chunk.patch_block(block_p);
    chunk.stack_switch_handlers.insert(
        resume_ip,
        vec![StackSwitchHandler {
            kind: 0,
            tag_index: tag as u32,
            label_index: handler_ip as u32,
        }],
    );
}

/// `$fiber->resume($v)` — resume with `$v`. Same shape as start;
/// distinct only at the AST level so future state checks can branch.
pub fn emit_php_fiber_resume(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_php_fiber_start(chunks, current, argc, line);
}

/// `$fiber->throw($exn)` — inject an exception into a suspended fiber.
/// Stack on entry: `[$fiber, $exn]`. Stack on exit follows `RESUME_THROW`.
pub fn emit_php_fiber_throw(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 1 {
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    } else if argc > 2 {
        for _ in 2..argc {
            chunk.emit_op(Op::DROP, line);
        }
    }
    let exn_slot = alloc_local(chunk);
    let fiber_slot = alloc_local(chunk);
    let running_key = class_slots::resolve_interned(chunk, &ClassSlot::internal("__running"), &PlainNames);
    let suspended_key = class_slots::resolve_interned(chunk, &ClassSlot::internal("__suspended"), &PlainNames);
    let cs_slot_1 = class_slots::resolve_interned(chunk, &ClassSlot::Internal(("__return").to_string()), &PlainNames);
    lset(chunk, exn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, fiber_slot, line);
    chunk.emit_bool_const(false, line);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &suspended_key, ValueSource::Stack, line);
    lget(chunk, fiber_slot, line);
    chunk.emit_bool_const(true, line);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &running_key, ValueSource::Stack, line);
    lget(chunk, fiber_slot, line);
    vybe_compiler::primitives::globals::emit_write(chunk, "__php_current_fiber", line);
    lget(chunk, fiber_slot, line);
    lget(chunk, exn_slot, line);
    let tag = 0;
    vybe_compiler::primitives::generators::emit_resume_throw_tagged(chunk, tag, line);
    let ret_slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_TEE, ret_slot, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    vybe_compiler::primitives::globals::emit_write(chunk, "__php_current_fiber", line);
    lget(chunk, fiber_slot, line);
    chunk.emit_bool_const(false, line);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &running_key, ValueSource::Stack, line);
    lget(chunk, fiber_slot, line);
    chunk.emit_bool_const(true, line);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &suspended_key, ValueSource::Stack, line);
    lget(chunk, fiber_slot, line);
    lget(chunk, ret_slot, line);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &cs_slot_1, ValueSource::Stack, line);
}

/// `$fiber->getReturn()` — return the fiber's return value. After the
/// fiber's body returns normally, the value is on the calling frame's
/// stack (RESUME left it there). PHP's getReturn() should be called
/// after fiber completion to retrieve that value, but we also stash
/// it in `__return` for the cases where the user calls it after
/// reading the value via the last `resume`.
///
/// Stack on entry: `[$fiber]`. Stack on exit: `[return_value]`.
///
/// MVP: read `__return` property; if absent, return null. The
/// `__return` is populated by the walker rewrite of resume that
/// stashes RESUME's return value before re-pushing it for the caller.
pub fn emit_php_fiber_get_return(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let cs_slot = class_slots::resolve(&ClassSlot::Internal(("__return").to_string()), &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &cs_slot, Dest::Stack, line);
}

/// State-check helpers — minimal MVP, all default to false (the
/// continuation Object doesn't currently expose phase queries through
/// public properties; pending VM-level state accessor opcodes).
pub fn emit_php_fiber_is_started(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let cs_slot = class_slots::resolve(&ClassSlot::Internal(("__started").to_string()), &PlainNames);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &cs_slot, Dest::Stack, line);
}
pub fn emit_php_fiber_is_suspended(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let cs_slot = class_slots::resolve(&ClassSlot::Internal(("__suspended").to_string()), &PlainNames);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &cs_slot, Dest::Stack, line);
}
pub fn emit_php_fiber_is_running(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let cs_slot = class_slots::resolve(&ClassSlot::Internal(("__running").to_string()), &PlainNames);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &cs_slot, Dest::Stack, line);
}
pub fn emit_php_fiber_is_terminated(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let cs_slot = class_slots::resolve(&ClassSlot::Internal(("__terminated").to_string()), &PlainNames);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &cs_slot, Dest::Stack, line);
}

// Suppress unused-import warnings if some helpers grow.
#[allow(dead_code)]
fn _touch(_c: &mut Chunk, _s: u16, _l: u32) {
    let _ = lset;
    let _ = lget;
    let _ = alloc_local;
}
