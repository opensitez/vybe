//! Component Model canonical built-in definitions.
//!
//! Binary encoding from proposals/component-model/design/mvp/Binary.md §Canon Definitions.
//! Prefix 0xF0 is the VM's encoding for the CM3 canon section.
//! Sub-values match the spec byte values exactly.

use super::Op;
use super::opcode_category;

impl Op {
    // ── canon lift / lower ──────────────────────────────────────────
    pub const CANON_LIFT: Op = Op::new(0xF0, 0x00);
    pub const CANON_LOWER: Op = Op::new(0xF0, 0x01);

    // ── resource ────────────────────────────────────────────────────
    pub const RESOURCE_NEW: Op = Op::new(0xF0, 0x02);
    pub const RESOURCE_DROP: Op = Op::new(0xF0, 0x03);
    pub const RESOURCE_REP: Op = Op::new(0xF0, 0x04);

    // ── task ────────────────────────────────────────────────────────
    pub const TASK_CANCEL: Op = Op::new(0xF0, 0x05);
    pub const SUBTASK_CANCEL: Op = Op::new(0xF0, 0x06);
    pub const BACKPRESSURE_SET: Op = Op::new(0xF0, 0x08);
    pub const TASK_RETURN: Op = Op::new(0xF0, 0x09);
    pub const CONTEXT_GET: Op = Op::new(0xF0, 0x0A);
    pub const CONTEXT_SET: Op = Op::new(0xF0, 0x0B);
    pub const THREAD_YIELD: Op = Op::new(0xF0, 0x0C);
    pub const SUBTASK_DROP: Op = Op::new(0xF0, 0x0D);

    // ── stream ──────────────────────────────────────────────────────
    pub const STREAM_NEW: Op = Op::new(0xF0, 0x0E);
    pub const STREAM_READ: Op = Op::new(0xF0, 0x0F);
    pub const STREAM_WRITE: Op = Op::new(0xF0, 0x10);
    pub const STREAM_CANCEL_READ: Op = Op::new(0xF0, 0x11);
    pub const STREAM_CANCEL_WRITE: Op = Op::new(0xF0, 0x12);
    pub const STREAM_DROP_RD: Op = Op::new(0xF0, 0x13);
    pub const STREAM_DROP_WR: Op = Op::new(0xF0, 0x14);

    // ── future ──────────────────────────────────────────────────────
    pub const FUTURE_NEW: Op = Op::new(0xF0, 0x15);
    pub const FUTURE_READ: Op = Op::new(0xF0, 0x16);
    pub const FUTURE_WRITE: Op = Op::new(0xF0, 0x17);
    pub const FUTURE_CANCEL_READ: Op = Op::new(0xF0, 0x18);
    pub const FUTURE_CANCEL_WRITE: Op = Op::new(0xF0, 0x19);
    pub const FUTURE_DROP_RD: Op = Op::new(0xF0, 0x1A);
    pub const FUTURE_DROP_WR: Op = Op::new(0xF0, 0x1B);

    // ── error-context ───────────────────────────────────────────────
    pub const ERROR_CONTEXT_NEW: Op = Op::new(0xF0, 0x1C);
    pub const ERROR_CONTEXT_DEBUG_MESSAGE: Op = Op::new(0xF0, 0x1D);
    pub const ERROR_CONTEXT_DROP: Op = Op::new(0xF0, 0x1E);

    // ── waitable ────────────────────────────────────────────────────
    pub const WAITABLE_SET_NEW: Op = Op::new(0xF0, 0x1F);
    pub const WAITABLE_SET_WAIT: Op = Op::new(0xF0, 0x20);
    pub const WAITABLE_SET_POLL: Op = Op::new(0xF0, 0x21);
    pub const WAITABLE_SET_DROP: Op = Op::new(0xF0, 0x22);
    pub const WAITABLE_JOIN: Op = Op::new(0xF0, 0x23);

    // ── backpressure (new) ──────────────────────────────────────────
    pub const BACKPRESSURE_INC: Op = Op::new(0xF0, 0x24);
    pub const BACKPRESSURE_DEC: Op = Op::new(0xF0, 0x25);

    // ── thread ──────────────────────────────────────────────────────
    pub const THREAD_INDEX: Op = Op::new(0xF0, 0x26);
    pub const THREAD_NEW_INDIRECT: Op = Op::new(0xF0, 0x27);
}

opcode_category! {
    [0x00] canon_lift => U16, "canon.lift";
    [0x01] canon_lower => U16, "canon.lower";
    [0x02] resource_new => U16, "resource.new";
    [0x03] resource_drop => U16, "resource.drop";
    [0x04] resource_rep => U16, "resource.rep";
    [0x05] task_cancel => None, "task.cancel";
    [0x06] subtask_cancel => None, "subtask.cancel";
    [0x08] backpressure_set => None, "backpressure.set";
    [0x09] task_return => None, "task.return";
    [0x0A] context_get => None, "context.get";
    [0x0B] context_set => None, "context.set";
    [0x0C] thread_yield => None, "thread.yield";
    [0x0D] subtask_drop => None, "subtask.drop";
    [0x0E] stream_new => None, "stream.new";
    [0x0F] stream_read => None, "stream.read";
    [0x10] stream_write => None, "stream.write";
    [0x11] stream_cancel_read => None, "stream.cancel-read";
    [0x12] stream_cancel_write => None, "stream.cancel-write";
    [0x13] stream_drop_rd => None, "stream.drop-readable";
    [0x14] stream_drop_wr => None, "stream.drop-writable";
    [0x15] future_new => None, "future.new";
    [0x16] future_read => None, "future.read";
    [0x17] future_write => None, "future.write";
    [0x18] future_cancel_read => None, "future.cancel-read";
    [0x19] future_cancel_write => None, "future.cancel-write";
    [0x1A] future_drop_rd => None, "future.drop-readable";
    [0x1B] future_drop_wr => None, "future.drop-writable";
    [0x1C] error_context_new => None, "error-context.new";
    [0x1D] error_context_debug_message => None, "error-context.debug-message";
    [0x1E] error_context_drop => None, "error-context.drop";
    [0x1F] waitable_set_new => None, "waitable-set.new";
    [0x20] waitable_set_wait => None, "waitable-set.wait";
    [0x21] waitable_set_poll => None, "waitable-set.poll";
    [0x22] waitable_set_drop => None, "waitable-set.drop";
    [0x23] waitable_join => None, "waitable.join";
    [0x24] backpressure_inc => None, "backpressure.inc";
    [0x25] backpressure_dec => None, "backpressure.dec";
    [0x26] thread_index => None, "thread.index";
    [0x27] thread_new_indirect => U16, "thread.new-indirect";
}
