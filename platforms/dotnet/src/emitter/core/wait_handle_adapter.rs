//! `System.Threading.AutoResetEvent` / `ManualResetEvent` — the `WaitHandle`
//! event pair.
//!
//! Neither was registered anywhere: `New AutoResetEvent(False)` produced an
//! object with no methods and `signal.WaitOne(1)` died as "undefined is not
//! callable". They account for the largest single cluster in
//! `vb_system_threading_matrix`.
//!
//! ## The signalling boundary — stated, not hidden
//!
//! A real `WaitOne(timeout)` BLOCKS until another thread signals or the
//! timeout expires. That needs a futex on a shared address:
//! `primitives::threading::emit_atomic_wait` / `emit_atomic_notify` are the
//! right primitives and they exist. What does NOT exist yet is a sound shared
//! WORD to point them at — the same gap `Interlocked` is currently trapping on
//! (`atomic unaligned access`), because a WASM atomic is ADDRESS-based and an
//! object field is a VALUE. Building on that today would inherit the trap.
//!
//! So the event's state lives on the INSTANCE, and a wait reports the state
//! rather than blocking on it. That is exact for a signal set before the wait
//! (which is what `Set()` then `WaitOne(ms)` does, and what the corpus
//! exercises) and for the timeout expiry of an unsignalled handle, which is
//! `False` either way. It is NOT exact for a wait that must OUTLIVE the call
//! until another thread signals: that returns `False` immediately instead of
//! blocking. When the shared-word allocation lands, `emit_wait_one` is the one
//! function to move onto `atomic_wait`.

use vybe_compiler::primitives::collections;
use vybe_compiler::primitives::ops;
use vybe_runtime::chunk::Chunk;
use vybe_runtime::opcode::Op;

/// Instance state keys.
///
/// ⛔ Prefixed and lowercase. A dotnet type with no property accessor resolves
/// its properties as an ordinary lowercased struct-field read, so a PascalCase
/// key is unreadable from a case-insensitive frontend — the convention
/// `thread_adapter`'s `CANCELLED_KEY` already documents.
const SIGNALED_KEY: &str = "__dotnet_signaled";
const AUTO_RESET_KEY: &str = "__dotnet_auto_reset";

fn set_flag_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    object: u16,
    key: &str,
    value: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn set_flag_const(chunks: &mut [Chunk], current: usize, object: u16, key: &str, on: bool, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_bool_const(on, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn get_flag(chunks: &mut [Chunk], current: usize, object: u16, key: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
    chunks[current].emit_string_const(key, line);
    collections::emit_get(chunks, current, line);
}

/// `New AutoResetEvent(initialState)` / `New ManualResetEvent(initialState)`.
///
/// `auto_reset` is the ONLY difference between the two classes: a successful
/// wait on an `AutoResetEvent` consumes the signal, a `ManualResetEvent` stays
/// signalled until `Reset()`. One body, one flag — not two adapters that could
/// drift.
///
/// Stack: `[initial_state]` → `[event]`.
fn emit_event_new(chunks: &mut [Chunk], current: usize, argc: u8, auto_reset: bool, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let (initial, event) = (base, base + 1);

    if argc > 0 {
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, initial, line);
    } else {
        // `New ManualResetEvent()` is not legal .NET, but a frontend that drops
        // the argument must not leave the slot holding whatever aliased it.
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, initial, line);
    }

    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, event, line);
    set_flag_from_slot(chunks, current, event, SIGNALED_KEY, initial, line);
    set_flag_const(chunks, current, event, AUTO_RESET_KEY, auto_reset, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, event, line);
}

pub fn emit_auto_reset_event_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_event_new(chunks, current, argc, true, line);
}

pub fn emit_manual_reset_event_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_event_new(chunks, current, argc, false, line);
}

/// `.Set()` — signal. Returns `True`, as .NET's does.
/// Stack: `[event]` → `[bool]`.
pub fn emit_event_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let event = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, event, line);
    set_flag_const(chunks, current, event, SIGNALED_KEY, true, line);
    chunks[current].emit_bool_const(true, line);
}

/// `.Reset()` — unsignal. Returns `True`, as .NET's does.
/// Stack: `[event]` → `[bool]`.
pub fn emit_event_reset(chunks: &mut [Chunk], current: usize, line: u32) {
    let event = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, event, line);
    set_flag_const(chunks, current, event, SIGNALED_KEY, false, line);
    chunks[current].emit_bool_const(true, line);
}

/// `.WaitOne()` / `.WaitOne(timeoutMs)` — did the handle come back signalled?
///
/// An `AutoResetEvent` CONSUMES the signal on a successful wait; a
/// `ManualResetEvent` does not. The consume is driven by the instance's own
/// `__dotnet_auto_reset` flag, so both classes share this body.
///
/// Stack: `[event]` or `[event, timeout]` → `[bool]`.
pub fn emit_wait_one(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let (event, signaled) = (base, base + 1);
    if argc > 0 {
        // The timeout is read and discarded: with the state on the instance
        // there is nothing to wait FOR. See the module header.
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, event, line);

    get_flag(chunks, current, event, SIGNALED_KEY, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, signaled, line);

    // if signaled AND auto_reset { signaled = false }
    chunks[current].emit_op_u16(Op::LOCAL_GET, signaled, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get_flag(chunks, current, event, AUTO_RESET_KEY, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    set_flag_const(chunks, current, event, SIGNALED_KEY, false, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    // ⛔ `emit_dyn_to_bool` leaves an i32 — the branch form. Returning it raw
    // printed `0`/`1` where .NET prints `False`/`True`; the value must be
    // MATERIALIZED back into a Bool for a frontend that renders one.
    chunks[current].emit_op_u16(Op::LOCAL_GET, signaled, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `WaitHandle.WaitAll(handles)` — every handle signalled?
///
/// Consumes the signal on each `AutoResetEvent` exactly as `WaitOne` does, so
/// a handle cannot be waited twice on one signal. The per-handle logic is the
/// same three reads as `emit_wait_one`; it is inlined per element rather than
/// factored into a called chunk because the loop already owns the slots.
///
/// Stack: `[handles]` → `[bool]`.
pub fn emit_wait_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let (handles, index, len, all, item) = (base, base + 1, base + 2, base + 3, base + 4);
    chunks[current].emit_op_u16(Op::LOCAL_SET, handles, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, handles, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
    // .NET's `WaitAll` on an EMPTY array is `True` — the vacuous case.
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, all, line);

    let block = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, handles, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item, line);

    get_flag(chunks, current, item, SIGNALED_KEY, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get_flag(chunks, current, item, AUTO_RESET_KEY, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    set_flag_const(chunks, current, item, SIGNALED_KEY, false, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, all, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, all, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}
