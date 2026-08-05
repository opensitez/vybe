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

use vybe_ast::{ChanOp, ExprKind, Expression, ObjectProperty, SelectArm, Statement, UnaryOp};
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

/// Stack: [] → [i32 bool]. Buffered-value-present test for the channel in `slot`.
fn has_buffered(imports: &mut Chunk, c: &mut Chunk, slot: u16, line: u32) {
    core_wasm::i32_const(c, line, 0);
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
    queue_into(imports, c, line);
    collections::emit_len_into(imports, c, line);
    ops::emit_dyn_lt_into(imports, c, line); // 0 < len
    ops::emit_dyn_to_bool_into(imports, c, line);
}

const DEADLOCK: &str = "all goroutines are asleep - deadlock!";

/// `__stdlib_chan_send(ch, v)` → null. Panics on closed/nil.
pub fn build_chan_send(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_chan_send");
    c.arity = 2;
    c.local_count = 2;
    let (ch, v, line) = (0u16, 1u16, 0u32);

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, ch, line);
    nil_check(&mut c, ch, DEADLOCK, line);
    closed_flag(imports, &mut c, ch, line);
    c.emit_if(line);
    throw_msg(&mut c, "send on closed channel", line);
    c.emit_end(line);

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    queue_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_GET, v, line);
    collections::emit_push_into(imports, &mut c, line); // [new_len]
    c.emit_op(Op::DROP, line);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    c.emit_op(Op::RETURN, line);
    c
}

/// Shared receive body. `with_ok` selects `[value, ok]` vs the bare value.
fn build_chan_recv_impl(imports: &mut Chunk, name: &str, with_ok: bool) -> Chunk {
    let mut c = Chunk::new(name);
    c.arity = 1;
    c.local_count = 1;
    let (ch, line) = (0u16, 0u32);

    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    deref_cell_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_SET, ch, line);
    nil_check(&mut c, ch, DEADLOCK, line);

    has_buffered(imports, &mut c, ch, line);
    c.emit_if_value(line);
    // buffered: shift the queue
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    queue_into(imports, &mut c, line);
    collections::emit_shift_into(imports, &mut c, line);
    if with_ok {
        c.emit_bool_const(true, line);
        collections::emit_array_new_into(imports, &mut c, 2, line);
    }
    c.emit_else(line);
    // drained (closed, or open-and-empty until blocking lands): the zero.
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    let zero_key = key(&mut c, "__zero");
    c.emit_struct_field_op(Op::STRUCT_GET, 0, zero_key, line);
    if with_ok {
        c.emit_bool_const(false, line);
        collections::emit_array_new_into(imports, &mut c, 2, line);
    }
    c.emit_end(line);
    c.emit_op(Op::RETURN, line);
    c
}

pub fn build_chan_recv(imports: &mut Chunk) -> Chunk {
    build_chan_recv_impl(imports, "__stdlib_chan_recv", false)
}

pub fn build_chan_recv_ok(imports: &mut Chunk) -> Chunk {
    build_chan_recv_impl(imports, "__stdlib_chan_recv_ok", true)
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
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    queue_into(imports, &mut c, line);
    collections::emit_len_into(imports, &mut c, line);
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
    let mut c = Chunk::new("__stdlib_chan_ready_send");
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
    c.emit_if_value(line);
    c.emit_bool_const(false, line);
    c.emit_else(line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    queue_into(imports, &mut c, line);
    collections::emit_len_into(imports, &mut c, line);
    c.emit_op_u16(Op::LOCAL_GET, ch, line);
    let cap_key = key(&mut c, "capacity");
    c.emit_struct_field_op(Op::STRUCT_GET, 0, cap_key, line);
    ops::emit_dyn_lt_into(imports, &mut c, line); // len < cap
    c.emit_end(line);
    c.emit_end(line);
    c.emit_op(Op::RETURN, line);
    c
}

// ── ChanOp / Select lowering ────────────────────────────────────────────────

/// The channel value's construction AST — the ONE place the shape is spelled.
fn channel_literal(capacity: Option<&Expression>, zero: &Expression) -> Expression {
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("queue"),
            value: Expression::new(ExprKind::Unary {
                op: UnaryOp::AddrOf,
                expr: Box::new(Expression::new(ExprKind::Array(Vec::new()))) }) },
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
                self.compile_expr(&literal)
            }
            ChanOp::Send { channel, value } => {
                self.chan_helper_call("__vybe_chan_send", &[channel, value])
            }
            ChanOp::Recv(ch) => self.chan_helper_call("__vybe_chan_recv", &[ch]),
            ChanOp::RecvOk(ch) => self.chan_helper_call("__vybe_chan_recv_ok", &[ch]),
            ChanOp::Len(ch) => self.chan_helper_call("__vybe_chan_len", &[ch]),
            ChanOp::Cap(ch) => self.chan_helper_call("__vybe_chan_cap", &[ch]),
            ChanOp::Close(ch) => self.chan_helper_call("__vybe_chan_close", &[ch]) }
    }

    /// `select` — readiness choice (Go §Select statements). Test each arm's
    /// communication for readiness in source order; run the first ready arm's
    /// body (whose FIRST statement performs the communication), else the
    /// default. Deterministic first-ready instead of Go's uniform-random pick
    /// — observable only in programs that race arms, which the non-blocking
    /// model cannot yet express. With nothing ready and no default this falls
    /// through (blocking select lands with the fiber rendezvous).
    pub(crate) fn emit_select(
        &mut self,
        arms: &[SelectArm],
        default: Option<&[Statement]>,
    ) -> Result<(), String> {
        let mut open_ifs = 0usize;
        for arm in arms {
            let (helper, ch): (&'static str, &Expression) = match &arm.comm {
                ChanOp::Send { channel, .. } => ("__vybe_chan_ready_send", channel),
                ChanOp::Recv(ch) | ChanOp::RecvOk(ch) => ("__vybe_chan_ready_recv", ch),
                other => {
                    return Err(format!("select arm cannot communicate via {other:?}"));
                }
            };
            self.chan_helper_call(helper, &[ch])?;
            let line = self.line;
            ops::emit_dyn_to_bool(&mut self.chunks[self.current], line);
            self.chunks[self.current].emit_if(line);
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
        }
        let line = self.line;
        for _ in 0..open_ifs {
            self.chunks[self.current].emit_end(line);
        }
        Ok(())
    }
}
