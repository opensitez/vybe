//! `System.Runtime.InteropServices.Marshal` — unmanaged allocation.
//!
//! ⛔ ONE ALLOCATOR, TWO SPELLINGS. `Marshal.AllocHGlobal(n)` is the same
//! operation as C's `malloc(n)`, and C reaches it through the shared
//! `primitives/memory.rs` — whose `heap_zeroed_bytes_sized` builds
//! `new Uint8Array(n)`, compiling to the host import `ecma:uint8array.new`.
//! This emits that same import, so a .NET allocation and a C allocation are the
//! SAME runtime representation rather than two parallel ones.
//!
//! `memory.rs` itself is an AST-constructor layer (libc's adapter returns
//! `Expression`s), and this is an emitter, so the shared thing here is the
//! representation and the host entry point, not the Rust function.
//!
//! `FreeHGlobal` answers null, matching `memory::free_value()`: allocations are
//! GC-owned, so freeing is an accounting no-op, and .NET's method returns void.

use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

/// `Marshal.AllocHGlobal(cb)` / `AllocCoTaskMem(cb)` — a zeroed byte block.
pub fn emit_alloc(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    if argc == 0 {
        chunk.emit_i32_const(0, line);
    }
    let idx = chunk.add_import("ecma:uint8array", "new");
    chunk.emit_call(idx, 1, line);
}

/// `Marshal.FreeHGlobal(ptr)` / `FreeCoTaskMem(ptr)`.
pub fn emit_free(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}
