use vybe_compiler::primitives::bigint::{self, ShiftKind};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// `UInt32` logical right shift. Stack: `[value, count] -> [number]`.
///
/// The common BigInt helper owns the fixed-width semantics; this adapter only
/// performs the dotnet boundary coercion from the platform's ordinary numeric
/// model and converts the 32-bit result back to a Number, which is exact across
/// the whole `UInt32` range.
pub fn emit_uint32_ushr(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let count = chunk.alloc_scratch(2);
    let value = count + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, count, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let bigint_ctor = chunk.add_import("ecma:bigint", "BigInt");
    chunk.emit_call(bigint_ctor, 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count, line);
    bigint::emit_wrapped_shift(chunk, 32, ShiftKind::Ushr, line);

    let number = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(number, 1, line);
}
