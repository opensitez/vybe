//! ONE lowering for the channel model — `vybe_ast::ChanOp` /
//! `StmtKind::Select` → runtime-helper calls.
//!
//! CSP semantics live in HELPER CHUNKS (`__stdlib_chan_*`), linked once per
//! program like every other runtime helper — not expanded inline per site
//! (the old builder inlined thousands of instructions of property
//! manipulation into a 10-line channel program). Go-spec anchored:
//!
//! - receive on a closed channel drains the buffer, then yields the element
//!   ZERO VALUE with `ok == false` (the zero travels WITH the channel — the
//!   walker stored it at `make()`, where the element type was known);
//! - send on a closed channel and close of nil/closed panic;
//! - a nil channel is never ready (`select` skips it; len/cap are 0);
//! - `select` readiness: receive = buffered value present OR closed;
//!   send = open with buffer room.
//!
//! Blocking `Send`/`Recv` (empty-buffer rendezvous) is fiber + scheduler
//! territory on the `DeferredSource` seam; until that lands, receive on an
//! open empty channel yields the zero value non-blockingly — the historical
//! behaviour, now in exactly one place.
//!
//! The channel VALUE is `{queue: cell([]), closed, capacity, __zero}` — the
//! same shape the retired AST builders produced (the queue cell keeps the
//! buffer shared across go-pointer copies), plus `__zero`.

use std::sync::Arc;

use vybe_ast::{ChanOp, ExprKind, Expression, ObjectProperty, SelectArm, Statement};
use vybe_runtime::Chunk;
use vybe_runtime::Value;
use vybe_runtime::opcode::Op;

use crate::primitives::collections;
use crate::primitives::errors;
use crate::primitives::instructions::core_wasm;
use crate::primitives::ops;

fn key(c: &mut Chunk, name: &str) -> u16 {
    c.add_constant(Value::String(Arc::from(name)))
}

/// TOS: [maybe-cell] → [value]. A go pointer / the queue field wraps its
/// target in `{__ref_kind: "cell", __value}`; unwrap if present. Imports
/// register on `imports` (chunks[0]) — helper-chunk convention.
fn deref_cell_into(imports: &mut Chunk, c: &mut Chunk, line: u32) {
    let obj = c.local_count;
    let out = c.local_count + 1;
    c.local_count += 2;
    c.emit_op_u16(Op::LOCAL_SET, obj, line);
    c.emit_op_u16(Op::LOCAL_GET, obj, line);
    c.emit_op_u16(Op::LOCAL_SET, out, line);

    // object-like = non-null and not undefined/number/string/boolean/bigint
    c.emit_op_u16(Op::LOCAL_GET, obj, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op(Op::I32_EQZ, line);
    for (module, func) in [
        ("wasm:js-undefined", "test"),
        ("wasm:js-number", "test"),
        ("wasm:js-string", "test"),
        ("wasm:js-boolean", "test"),
        ("wasm:js-bigint", "test"),
    ] {
        c.emit_op_u16(Op::LOCAL_GET, obj, line);
        collections::emit_import_call_into(imports, c, module, func, 1, line);
        c.emit_op(Op::I32_EQZ, line);
        c.emit_op(Op::I32_AND, line);
    }
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, obj, line);
    let kind_key = key(c, "__ref_kind");
    c.emit_struct_field_op(Op::STRUCT_GET, 0, kind_key, line);
    c.emit_string_const("cell", line);
    ops::emit_dyn_eq_into(imports, c, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, obj, line);
    let value_key = key(c, "__value");
    c.emit_struct_field_op(Op::STRUCT_GET, 0, value_key, line);
    c.emit_op_u16(Op::LOCAL_SET, out, line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_GET, out, line);
}

/// TOS: [ch] → [queue-array]. Assumes ch already deref'd and non-null.
fn queue_into(imports: &mut Chunk, c: &mut Chunk, line: u32) {
    let queue_key = key(c, "queue");
    c.emit_struct_field_op(Op::STRUCT_GET, 0, queue_key, line);
    deref_cell_into(imports, c, line);
}

/// Stack: [] → diverges. Throw a string-payload panic.
fn throw_msg(c: &mut Chunk, msg: &str, line: u32) {
    c.emit_string_const(msg, line);
    errors::emit_throw(c, line);
}

/// Emit `if ch is null { throw msg }` with the deref'd channel in `slot`.
fn nil_check(c: &mut Chunk, slot: u16, msg: &str, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if(line);
    throw_msg(c, msg, line);
    c.emit_end(line);
}

/// Stack: [] → [i32 bool]. Read the `closed` flag of the channel in `slot`.
fn closed_flag(imports: &mut Chunk, c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
    let closed_key = key(c, "closed");
    c.emit_struct_field_op(Op::STRUCT_GET, 0, closed_key, line);
    ops::emit_dyn_to_bool_into(imports, c, line);
}

/// Stack: [] → [i32 bool]. Buffered-value-present test for the channel in
/// `slot` — reads the futex COUNT word, the authoritative buffered count
/// under the blocking protocol (the array lags mid-reservation).
fn has_buffered(imports: &mut Chunk, c: &mut Chunk, slot: u16, line: u32) {
    let _ = imports;
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
    let k = key(c, "__futex");
    c.emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
    atomic(c, Op::I32_ATOMIC_LOAD, line);
    core_wasm::i32_const(c, line, 0);
    c.emit_op(Op::I32_GT_S, line);
}

const DEADLOCK: &str = "all goroutines are asleep - deadlock!";

/// One futex poll interval. Deadlock is DETECTED, not assumed: after a
/// timed-out slice the waiter asks `wasm:threads.all_parked` whether every
/// other live thread is blocked in `wait32`; PARKED_STREAK consecutive true
/// readings (each a slice apart — a single reading can race a thread that
/// is between slices) → the Go runtime's deadlock panic. A sibling that is
/// COMPUTING keeps the reading false, so long pre-send compute can never
/// false-panic. WAIT_MAX_ATTEMPTS (~2min, beyond any harness timeout) is
/// pure insurance for detector-blind states; OS joins are counted as
/// parked, so none are currently known.
const WAIT_SLICE_NS: i64 = 20_000_000;
const WAIT_MAX_ATTEMPTS: i32 = 6000;
const PARKED_STREAK: i32 = 3;

/// Emit an atomic op with its memarg, GRID-ALIGNED and SPEC-ALIGNED: the
/// VM's structural scanners walk a 4-byte instruction grid, so both memarg
/// fields are padded to two bytes with non-minimal LEBs. The align field
/// is 2 (`0x82 0x00`) — the threads spec REQUIRES an atomic's natural
/// alignment (log2; 2 for the 32-bit class this file uses, wait32 and
/// notify included) and a real engine's validator rejects anything else.
/// Offset stays 0 (`0x80 0x00`).
fn atomic(c: &mut Chunk, op: Op, line: u32) {
    c.emit_op(op, line);
    c.emit(0x82, line);
    c.emit(0x00, line);
    c.emit(0x80, line);
    c.emit(0x00, line);
}

/// [ ] → [ ] — load ch.__futex into `addr_slot` (i32-valued property).
fn load_futex_addr(c: &mut Chunk, ch: u16, addr_slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    let k = key(c, "__futex");
    c.emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
    c.emit_op_u16(Op::LOCAL_SET, addr_slot, line);
}

/// [ ] → [ ] — atomically add `delta` to the channel's RECV-WAITER word
/// (`addr+4`, the second half of the 8-byte reservation). A nonzero count
/// means a receiver is blocked — the signal that makes an UNBUFFERED send
/// ready (Go: rendezvous readiness is "a receiver is waiting").
fn bump_recv_waiters(c: &mut Chunk, addr_slot: u16, delta: i32, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, addr_slot, line);
    c.emit_i32_const(4, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_i32_const(delta, line);
    atomic(c, Op::I32_ATOMIC_RMW_ADD, line);
    c.emit_op(Op::DROP, line);
}

/// [ ] → [ i32 ] — load the RECV-WAITER count.
fn load_recv_waiters(c: &mut Chunk, addr_slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, addr_slot, line);
    c.emit_i32_const(4, line);
    c.emit_op(Op::I32_ADD, line);
    atomic(c, Op::I32_ATOMIC_LOAD, line);
}

/// [ ] → [ ] — wake every waiter on the channel's futex word.
fn notify_all(c: &mut Chunk, addr_slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, addr_slot, line);
    c.emit_i32_const(i32::MAX, line);
    atomic(c, Op::MEMORY_ATOMIC_NOTIFY, line);
    c.emit_op(Op::DROP, line);
}

/// [ ] → [ ] — one bounded wait slice on the futex word (expected value in
/// `expected_slot`) with DETECTED deadlock: a timed-out (fully quiet) slice
/// where `wasm:threads.all_parked` also reads true bumps `streak_slot`;
/// PARKED_STREAK consecutive such slices → the Go deadlock panic. Any wake,
/// value change, or runnable sibling resets the streak. `attempts_slot` is
/// the coarse safety net for states the detector cannot see.
fn wait_slice(
    imports: &mut Chunk,
    c: &mut Chunk,
    addr_slot: u16,
    expected_slot: u16,
    attempts_slot: u16,
    streak_slot: u16,
    line: u32,
) {
    c.emit_op_u16(Op::LOCAL_GET, attempts_slot, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_TEE, attempts_slot, line);
    c.emit_i32_const(WAIT_MAX_ATTEMPTS, line);
    c.emit_op(Op::I32_GT_S, line);
    c.emit_if(line);
    throw_msg(c, DEADLOCK, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_GET, addr_slot, line);
    c.emit_op_u16(Op::LOCAL_GET, expected_slot, line);
    c.emit_i64_const(WAIT_SLICE_NS, line);
    atomic(c, Op::MEMORY_ATOMIC_WAIT32, line);
    // 2 = timed-out: the whole slice passed with no wake and no change.
    c.emit_i32_const(2, line);
    c.emit_op(Op::I32_EQ, line);
    c.emit_if(line);
    collections::emit_import_call_into(imports, c, "wasm:threads", "all_parked", 0, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, streak_slot, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_TEE, streak_slot, line);
    c.emit_i32_const(PARKED_STREAK - 1, line);
    c.emit_op(Op::I32_GT_S, line);
    c.emit_if(line);
    throw_msg(c, DEADLOCK, line);
    c.emit_end(line);
    c.emit_else(line);
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, streak_slot, line);
    c.emit_end(line);
    c.emit_else(line);
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_SET, streak_slot, line);
    c.emit_end(line);
}

/// `__stdlib_chan_send(ch, v)` → null. Panics on closed/nil.
pub fn build_chan_send(imports: &mut Chunk) -> Chunk {
    // BLOCKING send — the count-word futex protocol. The channel's `__futex`
    // word holds the buffered COUNT; reservations are atomic cmpxchg so
    // concurrent goroutines (OS threads sharing the memory) never overfill.
    // An unbuffered channel is capacity-1 storage plus "wait until the value
    // is taken" — the Go rendezvous.
    let mut c = Chunk::new("__stdlib_chan_send");
    c.arity = 2;
    c.local_count = 9; // ch, v, addr, cap, c, c2, attempts, eff, streak
    let (ch, v, line) = (0u16, 1u16, 0u32);
    let addr = 2u16;
    let cap = 3u16;
    let cnt = 4u16;
    let cnt2 = 5u16;
    let attempts = 6u16;
    let eff = 7u16;
    let streak = 8u16;

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, ch, line);
    nil_check(&mut c, ch, DEADLOCK, line);
    load_futex_addr(&mut c, ch, addr, line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    let cap_key = key(&mut c, "capacity");
    c.emit_struct_field_op(Op::STRUCT_GET, 0, cap_key, line);
    c.emit_op_u16(Op::LOCAL_SET, cap, line);
    // eff = cap == 0 ? 1 : cap  (unbuffered = one-slot rendezvous)
    c.emit_op_u16(Op::LOCAL_GET, cap, line);
    core_wasm::i32_const(&mut c, line, 0);
    ops::emit_dyn_eq_into(imports, &mut c, line);
    c.emit_if_value(line);
    core_wasm::i32_const(&mut c, line, 1);
    c.emit_else(line);
    c.emit_op_u16(Op::LOCAL_GET, cap, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_SET, eff, line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_op_u16(Op::LOCAL_SET, attempts, line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_op_u16(Op::LOCAL_SET, streak, line);

    let out_block = c.emit_block(line); // $out
    let (retry_loop, _) = c.emit_loop_s(line); // $retry
    // closed → panic (send on closed channel)
    closed_flag(imports, &mut c, ch, line);
    c.emit_if(line);
    throw_msg(&mut c, "send on closed channel", line);
    c.emit_end(line);
    // c = atomic.load(addr)
    c.emit_op_u16(Op::LOCAL_GET, addr, line);
    atomic(&mut c, Op::I32_ATOMIC_LOAD, line);
    c.emit_op_u16(Op::LOCAL_SET, cnt, line);
    // room? cnt < eff (dyn compare handles i32 vs f64 capacity)
    c.emit_op_u16(Op::LOCAL_GET, cnt, line);
    c.emit_op_u16(Op::LOCAL_GET, eff, line);
    ops::emit_dyn_lt_into(imports, &mut c, line);
    ops::emit_dyn_to_bool_into(imports, &mut c, line);
    c.emit_if(line); // $if1
    // reserve: cmpxchg(addr, cnt, cnt+1)
    c.emit_op_u16(Op::LOCAL_GET, addr, line);
    c.emit_op_u16(Op::LOCAL_GET, cnt, line);
    c.emit_op_u16(Op::LOCAL_GET, cnt, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    atomic(&mut c, Op::I32_ATOMIC_RMW_CMPXCHG, line);
    c.emit_op_u16(Op::LOCAL_GET, cnt, line);
    c.emit_op(Op::I32_EQ, line);
    c.emit_if(line); // $if2 — reserved
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    queue_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_GET, v, line);
    collections::emit_push_into(imports, &mut c, line);
    c.emit_op(Op::DROP, line);
    notify_all(&mut c, addr, line);
    // unbuffered: wait until the value is TAKEN (count back to 0)
    c.emit_op_u16(Op::LOCAL_GET, cap, line);
    core_wasm::i32_const(&mut c, line, 0);
    ops::emit_dyn_eq_into(imports, &mut c, line);
    c.emit_if(line); // $if3
    let (drain_loop, _) = c.emit_loop_s(line); // $drain
    c.emit_op_u16(Op::LOCAL_GET, addr, line);
    atomic(&mut c, Op::I32_ATOMIC_LOAD, line);
    c.emit_op_u16(Op::LOCAL_TEE, cnt2, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line); // taken
    c.emit_br(6, line); // → $out (if(0) $drain(1) $if3(2) $if2(3) $if1(4) $retry(5) $out(6))
    c.emit_end(line);
    closed_flag(imports, &mut c, ch, line);
    c.emit_if(line);
    throw_msg(&mut c, "send on closed channel", line);
    c.emit_end(line);
    wait_slice(imports, &mut c, addr, cnt2, attempts, streak, line);
    c.emit_br(0, line); // → $drain
    c.emit_end(line);
    c.patch_loop(drain_loop);
    c.emit_end(line); // $if3
    c.emit_br(3, line); // → $out ($if2(0) $if1(1) $retry(2) $out(3))
    c.emit_end(line); // $if2 — reservation lost: fall to retry
    c.emit_br(1, line); // → $retry ($if1(0) $retry(1))
    c.emit_end(line); // $if1
    // no room: wait for a receiver
    wait_slice(imports, &mut c, addr, cnt, attempts, streak, line);
    c.emit_br(0, line); // → $retry
    c.emit_end(line);
    c.patch_loop(retry_loop);
    c.emit_end(line);
    c.patch_block(out_block);

    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    c.emit_op(Op::RETURN, line);
    c
}

/// Shared receive body — BLOCKING. `with_ok` selects `[value, ok]` vs the
/// bare value. Protocol: reserve one buffered value via cmpxchg on the count
/// word, spin the paired `shift` until the sender's push lands, notify (a
/// blocked unbuffered sender is waiting for count→0); on empty: closed →
/// zero(+false), open → futex-wait a slice and retry (deadlock cap applies).
fn build_chan_recv_impl(
    imports: &mut Chunk,
    name: &str,
    with_ok: bool,
    throw_err_arg: bool,
) -> Chunk {
    // `throw_err_arg`: arity-2 variant `(ch, err)` — a closed+drained
    // channel THROWS `err` instead of yielding the zero value (.NET
    // ReadAsync / Rust recv() semantics; the error is the CALLER's).
    let mut c = Chunk::new(name);
    let base = if throw_err_arg { 2u16 } else { 1u16 };
    c.arity = base as u8;
    c.local_count = base + 7; // addr, cnt, attempts, v, result, scratch, streak
    let (ch, line) = (0u16, 0u32);
    let addr = base;
    let cnt = base + 1;
    let attempts = base + 2;
    let v = base + 3;
    let result = base + 4;
    let streak = base + 6;

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, ch, line);
    nil_check(&mut c, ch, DEADLOCK, line);
    load_futex_addr(&mut c, ch, addr, line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_op_u16(Op::LOCAL_SET, attempts, line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_op_u16(Op::LOCAL_SET, streak, line);

    let done_block = c.emit_block(line); // $done
    let (retry_loop, _) = c.emit_loop_s(line); // $retry
    c.emit_op_u16(Op::LOCAL_GET, addr, line);
    atomic(&mut c, Op::I32_ATOMIC_LOAD, line);
    c.emit_op_u16(Op::LOCAL_SET, cnt, line);
    c.emit_op_u16(Op::LOCAL_GET, cnt, line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_op(Op::I32_GT_S, line);
    c.emit_if(line); // $if1 — buffered value claimed?
    c.emit_op_u16(Op::LOCAL_GET, addr, line);
    c.emit_op_u16(Op::LOCAL_GET, cnt, line);
    c.emit_op_u16(Op::LOCAL_GET, cnt, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_SUB, line);
    atomic(&mut c, Op::I32_ATOMIC_RMW_CMPXCHG, line);
    c.emit_op_u16(Op::LOCAL_GET, cnt, line);
    c.emit_op(Op::I32_EQ, line);
    c.emit_if(line); // $if2 — reserved
    let (take_loop, _) = c.emit_loop_s(line); // $take — push may be in flight
    // The reservation (count cmpxchg) can land BEFORE the sender's paired
    // queue.push: gate the take on the ARRAY length, never on the shifted
    // value (the polymorphic shift's empty-case return is not a reliable
    // sentinel — measured: a raced receiver took `false` as the value).
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    queue_into(imports, &mut c, line);
    collections::emit_len_into(imports, &mut c, line);
    ops::emit_dyn_lt_into(imports, &mut c, line); // 0 < len
    ops::emit_dyn_to_bool_into(imports, &mut c, line);
    c.emit_if(line); // $if3
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    queue_into(imports, &mut c, line);
    collections::emit_shift_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, v, line);
    c.emit_op_u16(Op::LOCAL_GET, v, line);
    if with_ok {
        c.emit_bool_const(true, line);
        collections::emit_array_new_into(imports, &mut c, 2, line);
    }
    c.emit_op_u16(Op::LOCAL_SET, result, line);
    notify_all(&mut c, addr, line);
    c.emit_br(5, line); // → $done ($if3(0) $take(1) $if2(2) $if1(3) $retry(4) $done(5))
    c.emit_end(line); // $if3
    c.emit_br(0, line); // → $take
    c.emit_end(line);
    c.patch_loop(take_loop);
    c.emit_end(line); // $if2 — reservation lost
    c.emit_br(1, line); // → $retry ($if1(0) $retry(1))
    c.emit_end(line); // $if1
    // empty: closed → zero(+false), or THROW the caller's error value
    closed_flag(imports, &mut c, ch, line);
    c.emit_if(line); // $if4
    if throw_err_arg {
        c.emit_op_u16(Op::LOCAL_GET, 1, line); // err arg
        c.emit_op(Op::THROW, line);
    } else {
        c.emit_op_u16(Op::LOCAL_GET, ch, line);
        let zero_key = key(&mut c, "__zero");
        c.emit_struct_field_op(Op::STRUCT_GET, 0, zero_key, line);
        if with_ok {
            c.emit_bool_const(false, line);
            collections::emit_array_new_into(imports, &mut c, 2, line);
        }
        c.emit_op_u16(Op::LOCAL_SET, result, line);
        c.emit_br(2, line); // → $done ($if4(0) $retry(1) $done(2))
    }
    c.emit_end(line); // $if4
    // open + empty: register as a WAITING RECEIVER (enables the paired
    // unbuffered send's readiness), wait a slice, deregister, retry
    bump_recv_waiters(&mut c, addr, 1, line);
    wait_slice(imports, &mut c, addr, cnt, attempts, streak, line);
    bump_recv_waiters(&mut c, addr, -1, line);
    c.emit_br(0, line); // → $retry
    c.emit_end(line);
    c.patch_loop(retry_loop);
    c.emit_end(line);
    c.patch_block(done_block);

    c.emit_op_u16(Op::LOCAL_GET, result, line);
    c.emit_op(Op::RETURN, line);
    c
}

pub fn build_chan_recv(imports: &mut Chunk) -> Chunk {
    build_chan_recv_impl(imports, "__stdlib_chan_recv", false, false)
}

pub fn build_chan_recv_ok(imports: &mut Chunk) -> Chunk {
    build_chan_recv_impl(imports, "__stdlib_chan_recv_ok", true, false)
}

/// `__stdlib_chan_len(ch)` → number (nil → 0).
pub fn build_chan_len(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_chan_len");
    c.arity = 1;
    c.local_count = 1;
    let (ch, line) = (0u16, 0u32);

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, ch, line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if_value(line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_else(line);
    // The COUNT word — authoritative under the blocking protocol.
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    let k = key(&mut c, "__futex");
    c.emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
    atomic(&mut c, Op::I32_ATOMIC_LOAD, line);
    c.emit_end(line);
    c.emit_op(Op::RETURN, line);
    c
}

/// `__stdlib_chan_cap(ch)` → number (nil → 0).
pub fn build_chan_cap(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_chan_cap");
    c.arity = 1;
    c.local_count = 1;
    let (ch, line) = (0u16, 0u32);

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, ch, line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if_value(line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_else(line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    let cap_key = key(&mut c, "capacity");
    c.emit_struct_field_op(Op::STRUCT_GET, 0, cap_key, line);
    c.emit_end(line);
    c.emit_op(Op::RETURN, line);
    c
}

/// `__stdlib_chan_close(ch)` → null. Panics on nil/double close.
pub fn build_chan_close(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_chan_close");
    c.arity = 1;
    c.local_count = 1;
    let (ch, line) = (0u16, 0u32);

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, ch, line);
    nil_check(&mut c, ch, "close of nil channel", line);
    closed_flag(imports, &mut c, ch, line);
    c.emit_if(line);
    throw_msg(&mut c, "close of closed channel", line);
    c.emit_end(line);

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    let closed_key = key(&mut c, "closed");
    c.emit_bool_const(true, line);
    c.emit_struct_field_op(Op::STRUCT_SET, 0, closed_key, line);
    // Wake every blocked receiver/sender: receivers observe `closed` and
    // yield the zero value; a blocked sender panics (Go semantics).
    let addr = c.local_count;
    c.local_count += 1;
    load_futex_addr(&mut c, ch, addr, line);
    notify_all(&mut c, addr, line);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    c.emit_op(Op::RETURN, line);
    c
}

/// `__stdlib_chan_ready_recv(ch)` → bool: non-nil && (buffered || closed).
pub fn build_chan_ready_recv(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_chan_ready_recv");
    c.arity = 1;
    c.local_count = 1;
    let (ch, line) = (0u16, 0u32);

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, ch, line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if_value(line);
    c.emit_bool_const(false, line);
    c.emit_else(line);
    has_buffered(imports, &mut c, ch, line);
    closed_flag(imports, &mut c, ch, line);
    c.emit_op(Op::I32_OR, line);
    ops::emit_i32_to_bool(&mut c, line);
    c.emit_end(line);
    c.emit_op(Op::RETURN, line);
    c
}

/// `__stdlib_chan_ready_send(ch)` → bool: non-nil && open && len < cap.
pub fn build_chan_ready_send(imports: &mut Chunk) -> Chunk {
    // non-nil && open && (count < cap || (cap == 0 && a receiver WAITS)) —
    // the second disjunct is Go's unbuffered rendezvous readiness.
    let mut c = Chunk::new("__stdlib_chan_ready_send");
    c.arity = 1;
    c.local_count = 2; // ch, addr
    let (ch, line) = (0u16, 0u32);
    let addr = 1u16;

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, ch, line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if_value(line);
    c.emit_bool_const(false, line);
    c.emit_else(line);
    closed_flag(imports, &mut c, ch, line);
    c.emit_if_value(line);
    c.emit_bool_const(false, line);
    c.emit_else(line);
    load_futex_addr(&mut c, ch, addr, line);
    c.emit_op_u16(Op::LOCAL_GET, addr, line);
    atomic(&mut c, Op::I32_ATOMIC_LOAD, line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    let cap_key = key(&mut c, "capacity");
    c.emit_struct_field_op(Op::STRUCT_GET, 0, cap_key, line);
    ops::emit_dyn_lt_into(imports, &mut c, line); // count < cap
    ops::emit_dyn_to_bool_into(imports, &mut c, line);
    // OR: unbuffered with a parked receiver
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    c.emit_struct_field_op(Op::STRUCT_GET, 0, cap_key, line);
    core_wasm::i32_const(&mut c, line, 0);
    ops::emit_dyn_eq_into(imports, &mut c, line);
    ops::emit_dyn_to_bool_into(imports, &mut c, line);
    c.emit_if_value(line);
    load_recv_waiters(&mut c, addr, line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_op(Op::I32_GT_S, line);
    c.emit_else(line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_end(line);
    c.emit_op(Op::I32_OR, line);
    ops::emit_i32_to_bool(&mut c, line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_op(Op::RETURN, line);
    c
}

/// `__stdlib_chan_wait_slice(ch, is_recv)` → null: ONE bounded futex wait
/// on the channel's count word (20ms slice; wakes early on any notify of
/// that word). `is_recv` marks a RECEIVE arm: the wait registers in the
/// channel's recv-waiter word so a paired unbuffered send sees readiness.
/// A nil channel returns immediately — the CALLER owns the attempts cap
/// that turns spinning into the Go deadlock panic. Blocking select's wait
/// primitive: one slice per arm per poll round.
pub fn build_chan_wait_slice(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_chan_wait_slice");
    c.arity = 2;
    c.local_count = 5; // ch, is_recv, addr, expected, recv_flag
    let (ch, is_recv, line) = (0u16, 1u16, 0u32);
    let addr = 2u16;
    let expected = 3u16;
    let recv_flag = 4u16;

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, ch, line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    load_futex_addr(&mut c, ch, addr, line);
    c.emit_op_u16(Op::LOCAL_GET, is_recv, line);
    ops::emit_dyn_to_bool_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_TEE, recv_flag, line);
    c.emit_if(line);
    bump_recv_waiters(&mut c, addr, 1, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_GET, addr, line);
    atomic(&mut c, Op::I32_ATOMIC_LOAD, line);
    c.emit_op_u16(Op::LOCAL_SET, expected, line);
    c.emit_op_u16(Op::LOCAL_GET, addr, line);
    c.emit_op_u16(Op::LOCAL_GET, expected, line);
    c.emit_i64_const(WAIT_SLICE_NS, line);
    atomic(&mut c, Op::MEMORY_ATOMIC_WAIT32, line);
    c.emit_op(Op::DROP, line);
    c.emit_op_u16(Op::LOCAL_GET, recv_flag, line);
    c.emit_if(line);
    bump_recv_waiters(&mut c, addr, -1, line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    c.emit_op(Op::RETURN, line);
    c
}

/// `__stdlib_chan_try_send(ch, v)` → bool: non-suspending send. False on
/// nil/closed/full; true iff the value was buffered. UNBUFFERED channels
/// report false unconditionally — rendezvous-readiness needs a waiter
/// count (same limit as select-send; only .NET consumes TrySend today and
/// it never creates capacity-0 channels).
pub fn build_chan_try_send(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_chan_try_send");
    c.arity = 2;
    c.local_count = 5; // ch, v, addr, cap, cnt
    let (ch, v, line) = (0u16, 1u16, 0u32);
    let addr = 2u16;
    let cap = 3u16;
    let cnt = 4u16;

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, ch, line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if(line);
    c.emit_bool_const(false, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    let cap_key = key(&mut c, "capacity");
    c.emit_struct_field_op(Op::STRUCT_GET, 0, cap_key, line);
    c.emit_op_u16(Op::LOCAL_SET, cap, line);
    load_futex_addr(&mut c, ch, addr, line);
    c.emit_op_u16(Op::LOCAL_GET, cap, line);
    core_wasm::i32_const(&mut c, line, 0);
    ops::emit_dyn_eq_into(imports, &mut c, line);
    c.emit_if(line);
    // UNBUFFERED: deliverable iff a receiver is parked. Reserve the
    // rendezvous slot (0→1) and hand the value off — the parked receiver
    // takes it on its next slice; the try-sender does NOT drain-wait.
    load_recv_waiters(&mut c, addr, line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_op(Op::I32_GT_S, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, addr, line);
    c.emit_i32_const(0, line);
    c.emit_i32_const(1, line);
    atomic(&mut c, Op::I32_ATOMIC_RMW_CMPXCHG, line);
    c.emit_op(Op::I32_EQZ, line); // old == 0 → reserved
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    queue_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_GET, v, line);
    collections::emit_push_into(imports, &mut c, line);
    c.emit_op(Op::DROP, line);
    notify_all(&mut c, addr, line);
    c.emit_bool_const(true, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_bool_const(false, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    let (retry_loop, _) = c.emit_loop_s(line); // $retry
    closed_flag(imports, &mut c, ch, line);
    c.emit_if(line);
    c.emit_bool_const(false, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_GET, addr, line);
    atomic(&mut c, Op::I32_ATOMIC_LOAD, line);
    c.emit_op_u16(Op::LOCAL_SET, cnt, line);
    c.emit_op_u16(Op::LOCAL_GET, cnt, line);
    c.emit_op_u16(Op::LOCAL_GET, cap, line);
    ops::emit_dyn_lt_into(imports, &mut c, line);
    ops::emit_dyn_to_bool_into(imports, &mut c, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line); // full
    c.emit_bool_const(false, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_GET, addr, line);
    c.emit_op_u16(Op::LOCAL_GET, cnt, line);
    c.emit_op_u16(Op::LOCAL_GET, cnt, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    atomic(&mut c, Op::I32_ATOMIC_RMW_CMPXCHG, line);
    c.emit_op_u16(Op::LOCAL_GET, cnt, line);
    c.emit_op(Op::I32_EQ, line);
    c.emit_if(line); // reserved
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    queue_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_GET, v, line);
    collections::emit_push_into(imports, &mut c, line);
    c.emit_op(Op::DROP, line);
    notify_all(&mut c, addr, line);
    c.emit_bool_const(true, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    // reservation lost to a racer: retry
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(retry_loop);
    c.emit_bool_const(false, line);
    c.emit_op(Op::RETURN, line);
    c
}

/// [ ] → [ [zero(ch), false] ] on the stack — the not-ready TryRecv/TryPeek
/// result (`ch` must be non-nil and deref'd).
fn emit_zero_false_pair(imports: &mut Chunk, c: &mut Chunk, ch: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    let zero_key = key(c, "__zero");
    c.emit_struct_field_op(Op::STRUCT_GET, 0, zero_key, line);
    c.emit_bool_const(false, line);
    collections::emit_array_new_into(imports, c, 2, line);
}

/// `__stdlib_chan_try_recv(ch)` → `[value, ok]`: non-suspending receive.
/// `ok == false` (value = the channel's zero) when nothing is buffered —
/// empty-open and closed-drained look the same, exactly the .NET
/// `TryRead(out v)` contract. Nil → `[null, false]`.
pub fn build_chan_try_recv(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_chan_try_recv");
    c.arity = 1;
    c.local_count = 4; // ch, addr, cnt, v
    let (ch, line) = (0u16, 0u32);
    let addr = 1u16;
    let cnt = 2u16;
    let v = 3u16;

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, ch, line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if(line);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    c.emit_bool_const(false, line);
    collections::emit_array_new_into(imports, &mut c, 2, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    load_futex_addr(&mut c, ch, addr, line);
    let (retry_loop, _) = c.emit_loop_s(line); // $retry
    c.emit_op_u16(Op::LOCAL_GET, addr, line);
    atomic(&mut c, Op::I32_ATOMIC_LOAD, line);
    c.emit_op_u16(Op::LOCAL_TEE, cnt, line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_op(Op::I32_GT_S, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line); // nothing buffered
    emit_zero_false_pair(imports, &mut c, ch, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_GET, addr, line);
    c.emit_op_u16(Op::LOCAL_GET, cnt, line);
    c.emit_op_u16(Op::LOCAL_GET, cnt, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_SUB, line);
    atomic(&mut c, Op::I32_ATOMIC_RMW_CMPXCHG, line);
    c.emit_op_u16(Op::LOCAL_GET, cnt, line);
    c.emit_op(Op::I32_EQ, line);
    c.emit_if(line); // reserved
    let (take_loop, _) = c.emit_loop_s(line); // $take — push may be in flight
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    queue_into(imports, &mut c, line);
    collections::emit_len_into(imports, &mut c, line);
    ops::emit_dyn_lt_into(imports, &mut c, line);
    ops::emit_dyn_to_bool_into(imports, &mut c, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    queue_into(imports, &mut c, line);
    collections::emit_shift_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, v, line);
    notify_all(&mut c, addr, line);
    c.emit_op_u16(Op::LOCAL_GET, v, line);
    c.emit_bool_const(true, line);
    collections::emit_array_new_into(imports, &mut c, 2, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    c.emit_br(0, line); // → $take
    c.emit_end(line);
    c.patch_loop(take_loop);
    c.emit_end(line); // reservation lost: retry
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(retry_loop);
    emit_zero_false_pair(imports, &mut c, ch, line);
    c.emit_op(Op::RETURN, line);
    c
}

/// `__stdlib_chan_try_peek(ch)` → `[value, ok]`: read the buffered head
/// WITHOUT consuming it. Not-ready (or a push still in flight) → zero+false.
pub fn build_chan_try_peek(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_chan_try_peek");
    c.arity = 1;
    c.local_count = 2; // ch, addr
    let (ch, line) = (0u16, 0u32);
    let addr = 1u16;

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, ch, line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if(line);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    c.emit_bool_const(false, line);
    collections::emit_array_new_into(imports, &mut c, 2, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    load_futex_addr(&mut c, ch, addr, line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    queue_into(imports, &mut c, line);
    collections::emit_len_into(imports, &mut c, line);
    ops::emit_dyn_lt_into(imports, &mut c, line);
    ops::emit_dyn_to_bool_into(imports, &mut c, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    queue_into(imports, &mut c, line);
    core_wasm::i32_const(&mut c, line, 0);
    collections::emit_get_into(imports, &mut c, line);
    c.emit_bool_const(true, line);
    collections::emit_array_new_into(imports, &mut c, 2, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    emit_zero_false_pair(imports, &mut c, ch, line);
    c.emit_op(Op::RETURN, line);
    c
}

/// `__stdlib_chan_drained(ch)` → bool: closed AND no buffered values — the
/// consumer's definitive "done" (.NET `Completion.IsCompleted`, Kotlin
/// `isClosedForReceive`). Nil → false.
pub fn build_chan_drained(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_chan_drained");
    c.arity = 1;
    c.local_count = 1;
    let (ch, line) = (0u16, 0u32);

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, ch, line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if_value(line);
    c.emit_bool_const(false, line);
    c.emit_else(line);
    closed_flag(imports, &mut c, ch, line);
    has_buffered(imports, &mut c, ch, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_op(Op::I32_AND, line);
    ops::emit_i32_to_bool(&mut c, line);
    c.emit_end(line);
    c.emit_op(Op::RETURN, line);
    c
}

/// `__stdlib_chan_closed(ch)` → bool: closed for WRITING (the `closed` flag
/// alone — buffered values may remain readable, unlike `drained`). .NET
/// `Writer.TryWrite` after `Complete()`, Kotlin `isClosedForSend`. Nil → false.
pub fn build_chan_closed(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_chan_closed");
    c.arity = 1;
    c.local_count = 1;
    let (ch, line) = (0u16, 0u32);

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, ch, line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if_value(line);
    c.emit_bool_const(false, line);
    c.emit_else(line);
    closed_flag(imports, &mut c, ch, line);
    ops::emit_i32_to_bool(&mut c, line);
    c.emit_end(line);
    c.emit_op(Op::RETURN, line);
    c
}

/// `__stdlib_chan_recv_or_throw(ch, err)` → value, or THROWS `err` when the
/// channel is closed and drained (blocking otherwise — rides the full
/// rendezvous). The failure value is the CALLER's policy: .NET ReadAsync
/// passes its ChannelClosedException message, Rust would pass RecvError.
pub fn build_chan_recv_or_throw(imports: &mut Chunk) -> Chunk {
    build_chan_recv_impl(imports, "__stdlib_chan_recv_or_throw", false, true)
}

/// `__stdlib_chan_wait_readable(ch)` → bool: block until a read would
/// succeed (buffered value) or the channel is definitively done (closed
/// AND drained → false). Registers as a waiting receiver so a paired
/// unbuffered send sees readiness. Deadlock detection as in recv.
pub fn build_chan_wait_readable(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_chan_wait_readable");
    c.arity = 1;
    c.local_count = 5; // ch, addr, cnt, attempts, streak
    let (ch, line) = (0u16, 0u32);
    let addr = 1u16;
    let cnt = 2u16;
    let attempts = 3u16;
    let streak = 4u16;

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, ch, line);
    nil_check(&mut c, ch, DEADLOCK, line);
    load_futex_addr(&mut c, ch, addr, line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_op_u16(Op::LOCAL_SET, attempts, line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_op_u16(Op::LOCAL_SET, streak, line);
    let (poll_loop, _) = c.emit_loop_s(line); // $poll
    has_buffered(imports, &mut c, ch, line);
    c.emit_if(line);
    c.emit_bool_const(true, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    closed_flag(imports, &mut c, ch, line);
    c.emit_if(line);
    c.emit_bool_const(false, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_GET, addr, line);
    atomic(&mut c, Op::I32_ATOMIC_LOAD, line);
    c.emit_op_u16(Op::LOCAL_SET, cnt, line);
    bump_recv_waiters(&mut c, addr, 1, line);
    wait_slice(imports, &mut c, addr, cnt, attempts, streak, line);
    bump_recv_waiters(&mut c, addr, -1, line);
    c.emit_br(0, line); // → $poll
    c.emit_end(line);
    c.patch_loop(poll_loop);
    c.emit_bool_const(false, line);
    c.emit_op(Op::RETURN, line);
    c
}

/// `__stdlib_futex_alloc16()` → i32 base: reserve 16 bytes in the shared
/// futex page (same bump global the channels use — a bump allocator
/// tolerates mixed 8/16-byte reservations) and zero the STATUS word at
/// base+4. The thread-start record: `{fn_table_index, status, user_arg}`.
pub fn build_futex_alloc16(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_futex_alloc16");
    c.arity = 0;
    c.local_count = 1; // base
    let (base, line) = (0u16, 0u32);

    crate::primitives::globals::emit_read(&mut c, "__vybe_chan_futex_next", line);
    c.emit_op_u16(Op::LOCAL_TEE, base, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op_u16(Op::LOCAL_GET, base, line);
    collections::emit_import_call_into(imports, &mut c, "wasm:js-undefined", "test", 1, line);
    c.emit_op(Op::I32_OR, line);
    c.emit_if(line);
    c.emit_i32_const(1, line);
    c.emit_op_u16(Op::MEMORY_GROW, 0, line);
    c.emit_i32_const(65536, line);
    c.emit_op(Op::I32_MUL, line);
    c.emit_op_u16(Op::LOCAL_SET, base, line);
    c.emit_end(line);
    c.emit_op_u16(Op::LOCAL_GET, base, line);
    c.emit_i32_const(16, line);
    c.emit_op(Op::I32_ADD, line);
    crate::primitives::globals::emit_write(&mut c, "__vybe_chan_futex_next", line);
    c.emit_op_u16(Op::LOCAL_GET, base, line);
    c.emit_i32_const(4, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_i32_const(0, line);
    atomic(&mut c, Op::I32_ATOMIC_STORE, line);
    c.emit_op_u16(Op::LOCAL_GET, base, line);
    c.emit_op(Op::RETURN, line);
    c
}

/// `__stdlib_task_new(tid, base)` → the Task object. The VM no longer
/// builds task objects — they are ordinary bytecode-built objects whose
/// `__futex` names the thread-start record (status word at +4).
pub fn build_task_new(imports: &mut Chunk) -> Chunk {
    let _ = imports;
    let mut c = Chunk::new("__stdlib_task_new");
    c.arity = 2;
    c.local_count = 3; // tid, base, obj
    let (tid, base, obj, line) = (0u16, 1u16, 2u16, 0u32);

    c.emit_struct_new(0, 0, line);
    c.emit_op_u16(Op::LOCAL_SET, obj, line);
    let mut set = |c: &mut Chunk, name: &str, push: &dyn Fn(&mut Chunk)| {
        c.emit_op_u16(Op::LOCAL_GET, obj, line);
        push(c);
        let k = key(c, name);
        c.emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
    };
    set(&mut c, "__type", &|c| c.emit_string_const("Task", line));
    set(&mut c, "__thread_id", &|c| {
        c.emit_op_u16(Op::LOCAL_GET, tid, line)
    });
    set(&mut c, "__futex", &|c| c.emit_op_u16(Op::LOCAL_GET, base, line));
    set(&mut c, "isalive", &|c| c.emit_bool_const(true, line));
    set(&mut c, "iscompleted", &|c| c.emit_bool_const(false, line));
    set(&mut c, "status", &|c| c.emit_string_const("Running", line));
    set(&mut c, "result", &|c| {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line)
    });
    c.emit_op_u16(Op::LOCAL_GET, obj, line);
    c.emit_op(Op::RETURN, line);
    c
}

/// `__stdlib_task_wait(task)` → i32 (0 ok / 1 faulted): the JOIN — a
/// spec-bytecode futex wait on the task's status word, wasi-threads'
/// sanctioned user-code join (the proposal deliberately has no join
/// primitive). Deadlock detection rides `wait_slice` as everywhere.
pub fn build_task_wait(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_task_wait");
    c.arity = 1;
    c.local_count = 6; // task, addr, s, attempts, streak, scratch
    let (task, line) = (0u16, 0u32);
    let addr = 1u16;
    let s = 2u16;
    let attempts = 3u16;
    let streak = 4u16;

    c.emit_op_u16(Op::LOCAL_GET, task, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, task, line);
    nil_check(&mut c, task, "Task.Wait on null task", line);
    c.emit_op_u16(Op::LOCAL_GET, task, line);
    let fk = key(&mut c, "__futex");
    c.emit_struct_field_op(Op::STRUCT_GET, 0, fk, line);
    c.emit_i32_const(4, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, addr, line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_op_u16(Op::LOCAL_SET, attempts, line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_op_u16(Op::LOCAL_SET, streak, line);
    let (poll, _) = c.emit_loop_s(line);
    c.emit_op_u16(Op::LOCAL_GET, addr, line);
    atomic(&mut c, Op::I32_ATOMIC_LOAD, line);
    c.emit_op_u16(Op::LOCAL_TEE, s, line);
    core_wasm::i32_const(&mut c, line, 0);
    c.emit_op(Op::I32_GT_S, line);
    c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, s, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_EQ, line);
    c.emit_if_value(line);
    c.emit_i32_const(0, line);
    c.emit_else(line);
    c.emit_i32_const(1, line);
    c.emit_end(line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);
    wait_slice(imports, &mut c, addr, s, attempts, streak, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(poll);
    c.emit_i32_const(1, line);
    c.emit_op(Op::RETURN, line);
    c
}

// ── ChanOp / Select lowering ────────────────────────────────────────────────

/// The channel value's construction AST — the ONE place the shape is spelled.
fn channel_literal(capacity: Option<&Expression>, zero: &Expression) -> Expression {
    Expression::new(ExprKind::Object(vec![
        // The queue is a PLAIN array: the channel OBJECT is the shared
        // identity (Arc), so the old pointer-cell wrapper bought nothing and
        // its extra deref hop mis-read under cross-thread capture cloning.
        ObjectProperty::KeyValue {
            key: Expression::string("queue"),
            value: Expression::new(ExprKind::Array(Vec::new())) },
        ObjectProperty::KeyValue {
            key: Expression::string("closed"),
            value: Expression::bool(false) },
        ObjectProperty::KeyValue {
            key: Expression::string("capacity"),
            value: capacity.cloned().unwrap_or_else(|| Expression::int(0)) },
        ObjectProperty::KeyValue {
            key: Expression::string("__zero"),
            value: zero.clone() },
    ]))
}

impl crate::primitives::Compiler {
    /// Call the named channel helper with the given argument expressions.
    fn chan_helper_call(
        &mut self,
        helper: &'static str,
        args: &[&Expression],
    ) -> Result<(), String> {
        let line = self.line;
        crate::primitives::bundle::emit_call_push_func(
            &mut self.chunks[self.current],
            helper,
            line,
        );
        for arg in args {
            self.compile_expr(arg)?;
        }
        crate::primitives::bundle::emit_call_invoke(
            &mut self.chunks[self.current],
            args.len() as u8,
            line,
        );
        Ok(())
    }

    pub(crate) fn emit_chan(&mut self, op: &ChanOp) -> Result<(), String> {
        match op {
            ChanOp::New { capacity, zero } => {
                let literal = channel_literal(capacity.as_deref(), zero);
                self.compile_expr(&literal)?;
                // Allocate the channel's futex COUNT word in shared linear
                // memory (bump allocator over a dedicated page; the word
                // doubles as buffered-count and wait/notify address —
                // `memory.atomic.wait32/notify`, the threads-proposal
                // primitives). Each spawned thread shares the memory, so
                // blocked goroutines wake cross-thread.
                let line = self.line;
                let addr_slot = self.define_local("__chan_futex_addr");
                crate::primitives::globals::emit_read(
                    &mut self.chunks[self.current],
                    "__vybe_chan_futex_next",
                    line,
                );
                self.emit_u16(Op::LOCAL_TEE, addr_slot);
                self.emit_u16(Op::LOCAL_GET, addr_slot);
                self.chunks[self.current].emit_op(Op::REF_IS_NULL, line);
                {
                    let undef = self.import("wasm:js-undefined", "test");
                    self.emit_u16(Op::LOCAL_GET, addr_slot);
                    self.emit_host_call(undef, 1);
                }
                self.chunks[self.current].emit_op(Op::I32_OR, line);
                self.chunks[self.current].emit_if(line);
                // First channel: claim a fresh page; base = old_pages * 64KiB.
                self.chunks[self.current].emit_i32_const(1, line);
                self.chunks[self.current].emit_op_u16(Op::MEMORY_GROW, 0, line);
                self.chunks[self.current].emit_i32_const(65536, line);
                self.chunks[self.current].emit_op(Op::I32_MUL, line);
                self.emit_u16(Op::LOCAL_SET, addr_slot);
                self.chunks[self.current].emit_end(line);
                self.chunks[self.current].emit_op(Op::DROP, line); // the TEE'd read
                // bump: next = addr + 8 (padded)
                self.emit_u16(Op::LOCAL_GET, addr_slot);
                self.chunks[self.current].emit_i32_const(8, line);
                self.chunks[self.current].emit_op(Op::I32_ADD, line);
                crate::primitives::globals::emit_write(
                    &mut self.chunks[self.current],
                    "__vybe_chan_futex_next",
                    line,
                );
                // zero the count word (memarg: natural align 2, offset 0 —
                // grid-padded LEBs, spec-valid)
                self.emit_u16(Op::LOCAL_GET, addr_slot);
                self.chunks[self.current].emit_i32_const(0, line);
                self.chunks[self.current].emit_op(Op::I32_ATOMIC_STORE, line);
                self.chunks[self.current].emit(0x82, line);
                self.chunks[self.current].emit(0x00, line);
                self.chunks[self.current].emit(0x80, line);
                self.chunks[self.current].emit(0x00, line);
                // zero the recv-waiter word (addr+4 — the second half of the
                // channel's 8-byte reservation; fresh pages arrive zeroed,
                // but the explicit store keeps the invariant local)
                self.emit_u16(Op::LOCAL_GET, addr_slot);
                self.chunks[self.current].emit_i32_const(4, line);
                self.chunks[self.current].emit_op(Op::I32_ADD, line);
                self.chunks[self.current].emit_i32_const(0, line);
                self.chunks[self.current].emit_op(Op::I32_ATOMIC_STORE, line);
                self.chunks[self.current].emit(0x82, line);
                self.chunks[self.current].emit(0x00, line);
                self.chunks[self.current].emit(0x80, line);
                self.chunks[self.current].emit(0x00, line);
                // stamp: [obj] dup; addr; STRUCT_SET __futex
                crate::primitives::instructions::core_wasm::dup(
                    &mut self.chunks[self.current],
                    line,
                );
                self.emit_u16(Op::LOCAL_GET, addr_slot);
                let futex_key = self.str_const("__futex");
                self.emit_struct_field_op(Op::STRUCT_SET, 0, futex_key);
                Ok(())
            }
            ChanOp::Send { channel, value } => {
                self.chan_helper_call("__vybe_chan_send", &[channel, value])
            }
            ChanOp::Recv(ch) => self.chan_helper_call("__vybe_chan_recv", &[ch]),
            ChanOp::RecvOk(ch) => self.chan_helper_call("__vybe_chan_recv_ok", &[ch]),
            ChanOp::Len(ch) => self.chan_helper_call("__vybe_chan_len", &[ch]),
            ChanOp::Cap(ch) => self.chan_helper_call("__vybe_chan_cap", &[ch]),
            ChanOp::Close(ch) => self.chan_helper_call("__vybe_chan_close", &[ch]),
            ChanOp::TrySend { channel, value } => {
                self.chan_helper_call("__vybe_chan_try_send", &[channel, value])
            }
            ChanOp::TryRecv(ch) => self.chan_helper_call("__vybe_chan_try_recv", &[ch]),
            ChanOp::TryPeek(ch) => self.chan_helper_call("__vybe_chan_try_peek", &[ch]),
            ChanOp::Drained(ch) => self.chan_helper_call("__vybe_chan_drained", &[ch]),
            ChanOp::Closed(ch) => self.chan_helper_call("__vybe_chan_closed", &[ch]),
            ChanOp::RecvOrFail { channel, error } => {
                self.chan_helper_call("__vybe_chan_recv_or_throw", &[channel, error])
            }
            ChanOp::WaitReadable(ch) => {
                self.chan_helper_call("__vybe_chan_wait_readable", &[ch])
            } }
    }

    /// `select` — readiness choice (Go §Select statements). Test each arm's
    /// communication for readiness in source order; run the first ready arm's
    /// body (whose FIRST statement performs the communication), else the
    /// default. With no default the whole chain sits in a POLL LOOP: nothing
    /// ready → one bounded futex slice per arm (a wake on any arm's count
    /// word re-polls early; a nil arm returns immediately) → retry, with the
    /// shared attempts cap turning forever into the Go deadlock panic.
    /// Deterministic first-ready instead of Go's uniform-random pick.
    /// KNOWN LIMIT: select-SEND on an UNBUFFERED channel is never ready
    /// (readiness is `count < cap`; true rendezvous-readiness needs a waiter
    /// count), so it blocks to the deadlock cap even with a live receiver.
    pub(crate) fn emit_select(
        &mut self,
        arms: &[SelectArm],
        default: Option<&[Statement]>,
    ) -> Result<(), String> {
        let line = self.line;
        let blocking = default.is_none();
        if arms.is_empty() {
            // Go: `select {}` with no default blocks forever — the deadlock
            // panic is its only observable behavior.
            if let Some(default) = default {
                for stmt in default {
                    self.compile_stmt(stmt)?;
                }
            } else {
                throw_msg(&mut self.chunks[self.current], DEADLOCK, line);
            }
            return Ok(());
        }
        // Channel operands evaluate exactly ONCE on entry (Go §Select) —
        // the poll loop depends on this: re-evaluating per round would
        // repeat side effects.
        let mut ready: Vec<(&'static str, u16, bool)> = Vec::with_capacity(arms.len());
        for arm in arms {
            let (helper, ch, is_recv): (&'static str, &Expression, bool) = match &arm.comm {
                ChanOp::Send { channel, .. } => ("__vybe_chan_ready_send", channel, false),
                ChanOp::Recv(ch) | ChanOp::RecvOk(ch) => ("__vybe_chan_ready_recv", ch, true),
                other => {
                    return Err(format!("select arm cannot communicate via {other:?}"));
                }
            };
            self.compile_expr(ch)?;
            let c = &mut self.chunks[self.current];
            let slot = c.alloc_scratch(1);
            c.emit_op_u16(Op::LOCAL_SET, slot, line);
            ready.push((helper, slot, is_recv));
        }
        let (attempts, poll_loop) = if blocking {
            let c = &mut self.chunks[self.current];
            let attempts = c.alloc_scratch(2); // attempts, parked-streak
            c.emit_i32_const(0, line);
            c.emit_op_u16(Op::LOCAL_SET, attempts, line);
            c.emit_i32_const(0, line);
            c.emit_op_u16(Op::LOCAL_SET, attempts + 1, line);
            let (l, _) = c.emit_loop_s(line);
            (attempts, Some(l))
        } else {
            (0u16, None)
        };
        let mut open_ifs = 0usize;
        for (i, arm) in arms.iter().enumerate() {
            {
                let (helper, slot, _) = ready[i];
                let c = &mut self.chunks[self.current];
                crate::primitives::bundle::emit_call_push_func(c, helper, line);
                c.emit_op_u16(Op::LOCAL_GET, slot, line);
                crate::primitives::bundle::emit_call_invoke(c, 1, line);
                ops::emit_dyn_to_bool(c, line);
                c.emit_if(line);
            }
            for stmt in &arm.body {
                self.compile_stmt(stmt)?;
            }
            self.chunks[self.current].emit_else(line);
            open_ifs += 1;
        }
        if let Some(default) = default {
            for stmt in default {
                self.compile_stmt(stmt)?;
            }
        } else {
            let all_parked = self.import("wasm:threads", "all_parked");
            let streak = attempts + 1;
            let c = &mut self.chunks[self.current];
            // Safety net for states the parked-detector cannot see.
            c.emit_op_u16(Op::LOCAL_GET, attempts, line);
            c.emit_i32_const(1, line);
            c.emit_op(Op::I32_ADD, line);
            c.emit_op_u16(Op::LOCAL_TEE, attempts, line);
            c.emit_i32_const(WAIT_MAX_ATTEMPTS, line);
            c.emit_op(Op::I32_GT_S, line);
            c.emit_if(line);
            throw_msg(c, DEADLOCK, line);
            c.emit_end(line);
            // One bounded futex slice per arm (nil arms return immediately);
            // recv arms register as waiting receivers for the slice so a
            // paired unbuffered send can see readiness.
            for &(_, slot, is_recv) in &ready {
                crate::primitives::bundle::emit_call_push_func(
                    c,
                    "__vybe_chan_wait_slice",
                    line,
                );
                c.emit_op_u16(Op::LOCAL_GET, slot, line);
                c.emit_bool_const(is_recv, line);
                crate::primitives::bundle::emit_call_invoke(c, 2, line);
                c.emit_op(Op::DROP, line);
            }
            // Detected deadlock: PARKED_STREAK consecutive rounds in which
            // every other live thread sat parked in wait32. A computing
            // sibling keeps the reading false — no compute-time false panic.
            self.emit_host_call(all_parked, 0);
            let c = &mut self.chunks[self.current];
            c.emit_if(line);
            c.emit_op_u16(Op::LOCAL_GET, streak, line);
            c.emit_i32_const(1, line);
            c.emit_op(Op::I32_ADD, line);
            c.emit_op_u16(Op::LOCAL_TEE, streak, line);
            c.emit_i32_const(PARKED_STREAK - 1, line);
            c.emit_op(Op::I32_GT_S, line);
            c.emit_if(line);
            throw_msg(c, DEADLOCK, line);
            c.emit_end(line);
            c.emit_else(line);
            c.emit_i32_const(0, line);
            c.emit_op_u16(Op::LOCAL_SET, streak, line);
            c.emit_end(line);
            // Innermost else: every arm's if is a label, the loop sits just
            // past the outermost.
            c.emit_br(open_ifs as u32, line);
        }
        for _ in 0..open_ifs {
            self.chunks[self.current].emit_end(line);
        }
        if let Some(l) = poll_loop {
            let c = &mut self.chunks[self.current];
            c.emit_end(line);
            c.patch_loop(l);
        }
        Ok(())
    }
}
