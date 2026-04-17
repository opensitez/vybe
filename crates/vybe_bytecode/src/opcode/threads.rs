//! Threads/atomics proposal opcodes (prefix 0xFE).
//! Byte values match the WASM threads specification.

use super::Op;
use super::opcode_category;

impl Op {
    pub const MEMORY_ATOMIC_NOTIFY: Op   = Op::new(0xFE, 0x00);
    pub const MEMORY_ATOMIC_WAIT32: Op   = Op::new(0xFE, 0x01);
    pub const ATOMIC_FENCE: Op           = Op::new(0xFE, 0x03);
    pub const I32_ATOMIC_LOAD: Op        = Op::new(0xFE, 0x10);
    pub const I64_ATOMIC_LOAD: Op        = Op::new(0xFE, 0x11);
    pub const I32_ATOMIC_STORE: Op       = Op::new(0xFE, 0x17);
    pub const I64_ATOMIC_STORE: Op       = Op::new(0xFE, 0x18);
    pub const I32_ATOMIC_RMW_ADD: Op     = Op::new(0xFE, 0x1E);
    pub const I32_ATOMIC_RMW_SUB: Op     = Op::new(0xFE, 0x1F);
    pub const I32_ATOMIC_RMW_AND: Op     = Op::new(0xFE, 0x20);
    pub const I32_ATOMIC_RMW_OR: Op      = Op::new(0xFE, 0x21);
    pub const I32_ATOMIC_RMW_XOR: Op     = Op::new(0xFE, 0x22);
    pub const I32_ATOMIC_RMW_XCHG: Op    = Op::new(0xFE, 0x23);
    pub const I32_ATOMIC_RMW_CMPXCHG: Op = Op::new(0xFE, 0x24);
    pub const I64_ATOMIC_RMW_ADD: Op     = Op::new(0xFE, 0x25);
    pub const I64_ATOMIC_RMW_SUB: Op     = Op::new(0xFE, 0x26);
    pub const I64_ATOMIC_RMW_CMPXCHG: Op = Op::new(0xFE, 0x2E);
    pub const THREAD_SPAWN: Op           = Op::new(0xFE, 0x80);
    pub const THREAD_JOIN: Op            = Op::new(0xFE, 0x81);
}

opcode_category! {
    [0x00] memory_atomic_notify => None, "memory.atomic.notify";
    [0x01] memory_atomic_wait32 => None, "memory.atomic.wait32";
    [0x03] atomic_fence => None, "atomic.fence";
    [0x10] i32_atomic_load => None, "i32.atomic.load";
    [0x11] i64_atomic_load => None, "i64.atomic.load";
    [0x17] i32_atomic_store => None, "i32.atomic.store";
    [0x18] i64_atomic_store => None, "i64.atomic.store";
    [0x1E] i32_atomic_rmw_add => None, "i32.atomic.rmw.add";
    [0x1F] i32_atomic_rmw_sub => None, "i32.atomic.rmw.sub";
    [0x20] i32_atomic_rmw_and => None, "i32.atomic.rmw.and";
    [0x21] i32_atomic_rmw_or => None, "i32.atomic.rmw.or";
    [0x22] i32_atomic_rmw_xor => None, "i32.atomic.rmw.xor";
    [0x23] i32_atomic_rmw_xchg => None, "i32.atomic.rmw.xchg";
    [0x24] i32_atomic_rmw_cmpxchg => None, "i32.atomic.rmw.cmpxchg";
    [0x25] i64_atomic_rmw_add => None, "i64.atomic.rmw.add";
    [0x26] i64_atomic_rmw_sub => None, "i64.atomic.rmw.sub";
    [0x2E] i64_atomic_rmw_cmpxchg => None, "i64.atomic.rmw.cmpxchg";
    [0x80] thread_spawn => None, "thread.spawn";
    [0x81] thread_join => None, "thread.join";
}
