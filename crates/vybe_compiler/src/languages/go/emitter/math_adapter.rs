//! Go-specific math helpers routed via `common:go.*`.
//!
//! `math/bits` population/leading-zero counts and `math.Remainder` compose
//! standard WASM opcodes. No host fns, no custom opcodes. Generic IEEE float
//! ops (copysign/signbit/…) live in the shared `emitter::math` module.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

pub fn emit_helper(name: &str, chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) -> bool {
    let chunk = &mut chunks[current];
    match name {
        // math/bits.OnesCount(x) — population count over the 64-bit value.
        "go.bits_ones_count" => {
            chunk.emit_op(Op::I64_TRUNC_F64_S, line);
            chunk.emit_op(Op::I64_POPCNT, line);
            chunk.emit_op(Op::F64_CONVERT_I64_S, line);
        }
        // math/bits.LeadingZeros(x) — count leading zero bits in the 64-bit value.
        "go.bits_leading_zeros" => {
            chunk.emit_op(Op::I64_TRUNC_F64_S, line);
            chunk.emit_op(Op::I64_CLZ, line);
            chunk.emit_op(Op::F64_CONVERT_I64_S, line);
        }
        // math.Remainder(x, y) — IEEE-754 remainder: x - round(x/y)*y with
        // round-half-to-even (WASM f64.nearest). Stack: [x, y].
        "go.math_remainder" => {
            let base = chunk.alloc_scratch(2);
            chunk.emit_op_u16(Op::LOCAL_SET, base + 1, line); // y
            chunk.emit_op_u16(Op::LOCAL_SET, base, line); // x
            chunk.emit_op_u16(Op::LOCAL_GET, base, line); // x
            chunk.emit_op_u16(Op::LOCAL_GET, base, line); // x
            chunk.emit_op_u16(Op::LOCAL_GET, base + 1, line); // y
            chunk.emit_op(Op::F64_DIV, line); // x/y
            chunk.emit_op(Op::F64_NEAREST, line); // round-half-even
            chunk.emit_op_u16(Op::LOCAL_GET, base + 1, line); // y
            chunk.emit_op(Op::F64_MUL, line); // round(x/y)*y
            chunk.emit_op(Op::F64_SUB, line); // x - round(x/y)*y
        }
        // time.Duration unit accessors — the receiver is the duration in
        // nanoseconds (a plain number); divide by the unit.
        "go.dur_minutes" => {
            chunk.emit_f64_const(60_000_000_000.0, line);
            chunk.emit_op(Op::F64_DIV, line);
        }
        "go.dur_seconds" => {
            chunk.emit_f64_const(1_000_000_000.0, line);
            chunk.emit_op(Op::F64_DIV, line);
        }
        "go.dur_hours" => {
            chunk.emit_f64_const(3_600_000_000_000.0, line);
            chunk.emit_op(Op::F64_DIV, line);
        }
        // Nanoseconds() returns the raw count.
        "go.dur_nanoseconds" => {}
        "go.dur_milliseconds" => {
            chunk.emit_f64_const(1_000_000.0, line);
            chunk.emit_op(Op::F64_DIV, line);
            chunk.emit_op(Op::F64_TRUNC, line);
        }
        "go.dur_microseconds" => {
            chunk.emit_f64_const(1_000.0, line);
            chunk.emit_op(Op::F64_DIV, line);
            chunk.emit_op(Op::F64_TRUNC, line);
        }
        _ => return false,
    }
    true
}
