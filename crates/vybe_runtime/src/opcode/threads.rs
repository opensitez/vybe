//! Threads/atomics proposal opcodes (prefix 0xFE).
//! Byte values match the WASM threads specification.

use super::Op;
use super::opcode_category;

impl Op {
    pub const MEMORY_ATOMIC_NOTIFY: Op = Op::new(0xFE, 0x00);
    pub const MEMORY_ATOMIC_WAIT32: Op = Op::new(0xFE, 0x01);
    pub const MEMORY_ATOMIC_WAIT64: Op = Op::new(0xFE, 0x02);
    pub const ATOMIC_FENCE: Op = Op::new(0xFE, 0x03);
    pub const I32_ATOMIC_LOAD: Op = Op::new(0xFE, 0x10);
    pub const I64_ATOMIC_LOAD: Op = Op::new(0xFE, 0x11);
    pub const I32_ATOMIC_LOAD8_U: Op = Op::new(0xFE, 0x12);
    pub const I32_ATOMIC_LOAD16_U: Op = Op::new(0xFE, 0x13);
    pub const I64_ATOMIC_LOAD8_U: Op = Op::new(0xFE, 0x14);
    pub const I64_ATOMIC_LOAD16_U: Op = Op::new(0xFE, 0x15);
    pub const I64_ATOMIC_LOAD32_U: Op = Op::new(0xFE, 0x16);
    pub const I32_ATOMIC_STORE: Op = Op::new(0xFE, 0x17);
    pub const I64_ATOMIC_STORE: Op = Op::new(0xFE, 0x18);
    pub const I32_ATOMIC_STORE8: Op = Op::new(0xFE, 0x19);
    pub const I32_ATOMIC_STORE16: Op = Op::new(0xFE, 0x1A);
    pub const I64_ATOMIC_STORE8: Op = Op::new(0xFE, 0x1B);
    pub const I64_ATOMIC_STORE16: Op = Op::new(0xFE, 0x1C);
    pub const I64_ATOMIC_STORE32: Op = Op::new(0xFE, 0x1D);
    pub const I32_ATOMIC_RMW_ADD: Op = Op::new(0xFE, 0x1E);
    pub const I64_ATOMIC_RMW_ADD: Op = Op::new(0xFE, 0x1F);
    pub const I32_ATOMIC_RMW8_ADD_U: Op = Op::new(0xFE, 0x20);
    pub const I32_ATOMIC_RMW16_ADD_U: Op = Op::new(0xFE, 0x21);
    pub const I64_ATOMIC_RMW8_ADD_U: Op = Op::new(0xFE, 0x22);
    pub const I64_ATOMIC_RMW16_ADD_U: Op = Op::new(0xFE, 0x23);
    pub const I64_ATOMIC_RMW32_ADD_U: Op = Op::new(0xFE, 0x24);
    pub const I32_ATOMIC_RMW_SUB: Op = Op::new(0xFE, 0x25);
    pub const I64_ATOMIC_RMW_SUB: Op = Op::new(0xFE, 0x26);
    pub const I32_ATOMIC_RMW8_SUB_U: Op = Op::new(0xFE, 0x27);
    pub const I32_ATOMIC_RMW16_SUB_U: Op = Op::new(0xFE, 0x28);
    pub const I64_ATOMIC_RMW8_SUB_U: Op = Op::new(0xFE, 0x29);
    pub const I64_ATOMIC_RMW16_SUB_U: Op = Op::new(0xFE, 0x2A);
    pub const I64_ATOMIC_RMW32_SUB_U: Op = Op::new(0xFE, 0x2B);
    pub const I32_ATOMIC_RMW_AND: Op = Op::new(0xFE, 0x2C);
    pub const I64_ATOMIC_RMW_AND: Op = Op::new(0xFE, 0x2D);
    pub const I32_ATOMIC_RMW8_AND_U: Op = Op::new(0xFE, 0x2E);
    pub const I32_ATOMIC_RMW16_AND_U: Op = Op::new(0xFE, 0x2F);
    pub const I64_ATOMIC_RMW8_AND_U: Op = Op::new(0xFE, 0x30);
    pub const I64_ATOMIC_RMW16_AND_U: Op = Op::new(0xFE, 0x31);
    pub const I64_ATOMIC_RMW32_AND_U: Op = Op::new(0xFE, 0x32);
    pub const I32_ATOMIC_RMW_OR: Op = Op::new(0xFE, 0x33);
    pub const I64_ATOMIC_RMW_OR: Op = Op::new(0xFE, 0x34);
    pub const I32_ATOMIC_RMW8_OR_U: Op = Op::new(0xFE, 0x35);
    pub const I32_ATOMIC_RMW16_OR_U: Op = Op::new(0xFE, 0x36);
    pub const I64_ATOMIC_RMW8_OR_U: Op = Op::new(0xFE, 0x37);
    pub const I64_ATOMIC_RMW16_OR_U: Op = Op::new(0xFE, 0x38);
    pub const I64_ATOMIC_RMW32_OR_U: Op = Op::new(0xFE, 0x39);
    pub const I32_ATOMIC_RMW_XOR: Op = Op::new(0xFE, 0x3A);
    pub const I64_ATOMIC_RMW_XOR: Op = Op::new(0xFE, 0x3B);
    pub const I32_ATOMIC_RMW8_XOR_U: Op = Op::new(0xFE, 0x3C);
    pub const I32_ATOMIC_RMW16_XOR_U: Op = Op::new(0xFE, 0x3D);
    pub const I64_ATOMIC_RMW8_XOR_U: Op = Op::new(0xFE, 0x3E);
    pub const I64_ATOMIC_RMW16_XOR_U: Op = Op::new(0xFE, 0x3F);
    pub const I64_ATOMIC_RMW32_XOR_U: Op = Op::new(0xFE, 0x40);
    pub const I32_ATOMIC_RMW_XCHG: Op = Op::new(0xFE, 0x41);
    pub const I64_ATOMIC_RMW_XCHG: Op = Op::new(0xFE, 0x42);
    pub const I32_ATOMIC_RMW8_XCHG_U: Op = Op::new(0xFE, 0x43);
    pub const I32_ATOMIC_RMW16_XCHG_U: Op = Op::new(0xFE, 0x44);
    pub const I64_ATOMIC_RMW8_XCHG_U: Op = Op::new(0xFE, 0x45);
    pub const I64_ATOMIC_RMW16_XCHG_U: Op = Op::new(0xFE, 0x46);
    pub const I64_ATOMIC_RMW32_XCHG_U: Op = Op::new(0xFE, 0x47);
    pub const I32_ATOMIC_RMW_CMPXCHG: Op = Op::new(0xFE, 0x48);
    pub const I64_ATOMIC_RMW_CMPXCHG: Op = Op::new(0xFE, 0x49);
    pub const I32_ATOMIC_RMW8_CMPXCHG_U: Op = Op::new(0xFE, 0x4A);
    pub const I32_ATOMIC_RMW16_CMPXCHG_U: Op = Op::new(0xFE, 0x4B);
    pub const I64_ATOMIC_RMW8_CMPXCHG_U: Op = Op::new(0xFE, 0x4C);
    pub const I64_ATOMIC_RMW16_CMPXCHG_U: Op = Op::new(0xFE, 0x4D);
    pub const I64_ATOMIC_RMW32_CMPXCHG_U: Op = Op::new(0xFE, 0x4E);
    // THREAD_SPAWN (0xFE 0x80) / THREAD_JOIN (0xFE 0x81) RETIRED 2026-08-06:
    // spawning is the `wasi:threads/thread-spawn` IMPORT (the VM is the
    // embedder implementation) and join is helper bytecode futex-waiting the
    // task's status word — wasi-threads deliberately has no join primitive.
}

opcode_category! {
    // Spec: notify/wait carry a memarg exactly like every other atomic —
    // and dispatch has always read one (`pop_atomic_addr`). These were
    // declared `None`, which made every operand_format-driven walk lie
    // about the instruction size.
    [0x00] memory_atomic_notify => MemArg, "memory.atomic.notify";
    [0x01] memory_atomic_wait32 => MemArg, "memory.atomic.wait32";
    [0x02] memory_atomic_wait64 => MemArg, "memory.atomic.wait64";
    [0x03] atomic_fence => U8, "atomic.fence";
    [0x10] i32_atomic_load => MemArg, "i32.atomic.load";
    [0x11] i64_atomic_load => MemArg, "i64.atomic.load";
    [0x12] i32_atomic_load8_u => MemArg, "i32.atomic.load8_u";
    [0x13] i32_atomic_load16_u => MemArg, "i32.atomic.load16_u";
    [0x14] i64_atomic_load8_u => MemArg, "i64.atomic.load8_u";
    [0x15] i64_atomic_load16_u => MemArg, "i64.atomic.load16_u";
    [0x16] i64_atomic_load32_u => MemArg, "i64.atomic.load32_u";
    [0x17] i32_atomic_store => MemArg, "i32.atomic.store";
    [0x18] i64_atomic_store => MemArg, "i64.atomic.store";
    [0x19] i32_atomic_store8 => MemArg, "i32.atomic.store8";
    [0x1A] i32_atomic_store16 => MemArg, "i32.atomic.store16";
    [0x1B] i64_atomic_store8 => MemArg, "i64.atomic.store8";
    [0x1C] i64_atomic_store16 => MemArg, "i64.atomic.store16";
    [0x1D] i64_atomic_store32 => MemArg, "i64.atomic.store32";
    [0x1E] i32_atomic_rmw_add => MemArg, "i32.atomic.rmw.add";
    [0x1F] i64_atomic_rmw_add => MemArg, "i64.atomic.rmw.add";
    [0x20] i32_atomic_rmw8_add_u => MemArg, "i32.atomic.rmw8.add_u";
    [0x21] i32_atomic_rmw16_add_u => MemArg, "i32.atomic.rmw16.add_u";
    [0x22] i64_atomic_rmw8_add_u => MemArg, "i64.atomic.rmw8.add_u";
    [0x23] i64_atomic_rmw16_add_u => MemArg, "i64.atomic.rmw16.add_u";
    [0x24] i64_atomic_rmw32_add_u => MemArg, "i64.atomic.rmw32.add_u";
    [0x25] i32_atomic_rmw_sub => MemArg, "i32.atomic.rmw.sub";
    [0x26] i64_atomic_rmw_sub => MemArg, "i64.atomic.rmw.sub";
    [0x27] i32_atomic_rmw8_sub_u => MemArg, "i32.atomic.rmw8.sub_u";
    [0x28] i32_atomic_rmw16_sub_u => MemArg, "i32.atomic.rmw16.sub_u";
    [0x29] i64_atomic_rmw8_sub_u => MemArg, "i64.atomic.rmw8.sub_u";
    [0x2A] i64_atomic_rmw16_sub_u => MemArg, "i64.atomic.rmw16.sub_u";
    [0x2B] i64_atomic_rmw32_sub_u => MemArg, "i64.atomic.rmw32.sub_u";
    [0x2C] i32_atomic_rmw_and => MemArg, "i32.atomic.rmw.and";
    [0x2D] i64_atomic_rmw_and => MemArg, "i64.atomic.rmw.and";
    [0x2E] i32_atomic_rmw8_and_u => MemArg, "i32.atomic.rmw8.and_u";
    [0x2F] i32_atomic_rmw16_and_u => MemArg, "i32.atomic.rmw16.and_u";
    [0x30] i64_atomic_rmw8_and_u => MemArg, "i64.atomic.rmw8.and_u";
    [0x31] i64_atomic_rmw16_and_u => MemArg, "i64.atomic.rmw16.and_u";
    [0x32] i64_atomic_rmw32_and_u => MemArg, "i64.atomic.rmw32.and_u";
    [0x33] i32_atomic_rmw_or => MemArg, "i32.atomic.rmw.or";
    [0x34] i64_atomic_rmw_or => MemArg, "i64.atomic.rmw.or";
    [0x35] i32_atomic_rmw8_or_u => MemArg, "i32.atomic.rmw8.or_u";
    [0x36] i32_atomic_rmw16_or_u => MemArg, "i32.atomic.rmw16.or_u";
    [0x37] i64_atomic_rmw8_or_u => MemArg, "i64.atomic.rmw8.or_u";
    [0x38] i64_atomic_rmw16_or_u => MemArg, "i64.atomic.rmw16.or_u";
    [0x39] i64_atomic_rmw32_or_u => MemArg, "i64.atomic.rmw32.or_u";
    [0x3A] i32_atomic_rmw_xor => MemArg, "i32.atomic.rmw.xor";
    [0x3B] i64_atomic_rmw_xor => MemArg, "i64.atomic.rmw.xor";
    [0x3C] i32_atomic_rmw8_xor_u => MemArg, "i32.atomic.rmw8.xor_u";
    [0x3D] i32_atomic_rmw16_xor_u => MemArg, "i32.atomic.rmw16.xor_u";
    [0x3E] i64_atomic_rmw8_xor_u => MemArg, "i64.atomic.rmw8.xor_u";
    [0x3F] i64_atomic_rmw16_xor_u => MemArg, "i64.atomic.rmw16.xor_u";
    [0x40] i64_atomic_rmw32_xor_u => MemArg, "i64.atomic.rmw32.xor_u";
    [0x41] i32_atomic_rmw_xchg => MemArg, "i32.atomic.rmw.xchg";
    [0x42] i64_atomic_rmw_xchg => MemArg, "i64.atomic.rmw.xchg";
    [0x43] i32_atomic_rmw8_xchg_u => MemArg, "i32.atomic.rmw8.xchg_u";
    [0x44] i32_atomic_rmw16_xchg_u => MemArg, "i32.atomic.rmw16.xchg_u";
    [0x45] i64_atomic_rmw8_xchg_u => MemArg, "i64.atomic.rmw8.xchg_u";
    [0x46] i64_atomic_rmw16_xchg_u => MemArg, "i64.atomic.rmw16.xchg_u";
    [0x47] i64_atomic_rmw32_xchg_u => MemArg, "i64.atomic.rmw32.xchg_u";
    [0x48] i32_atomic_rmw_cmpxchg => MemArg, "i32.atomic.rmw.cmpxchg";
    [0x49] i64_atomic_rmw_cmpxchg => MemArg, "i64.atomic.rmw.cmpxchg";
    [0x4A] i32_atomic_rmw8_cmpxchg_u => MemArg, "i32.atomic.rmw8.cmpxchg_u";
    [0x4B] i32_atomic_rmw16_cmpxchg_u => MemArg, "i32.atomic.rmw16.cmpxchg_u";
    [0x4C] i64_atomic_rmw8_cmpxchg_u => MemArg, "i64.atomic.rmw8.cmpxchg_u";
    [0x4D] i64_atomic_rmw16_cmpxchg_u => MemArg, "i64.atomic.rmw16.cmpxchg_u";
    [0x4E] i64_atomic_rmw32_cmpxchg_u => MemArg, "i64.atomic.rmw32.cmpxchg_u";
    // 0x80/0x81 (`thread.spawn`/`thread.join`) DELETED — they were custom
    // opcodes squatting on the spec 0xFE prefix. Spawning is the
    // `wasi:threads/thread-spawn` import; join is futex helper bytecode.
    // Rows removed so stale bytecode fails decode loudly.
}
