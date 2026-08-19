//! Fortran bit-counting intrinsics — `popcnt`, `leadz`, `trailz`.
//!
//! Every other bit intrinsic decomposes into operators the AST already has
//! (`iand` is `BitAnd`, `ishft` is a shift, `btest` is a shift and a compare),
//! so the walker folds those and nothing reaches here. Counting bits is the
//! one thing no operator expresses — and it is a single core WASM opcode:
//! `i32.popcnt`, `i32.clz`, `i32.ctz`. `leadz(0)` and `trailz(0)` answering 32
//! is the opcode's own definition, not a special case bolted on.
//!
//! Default INTEGER is kind 4, so these are the 32-bit forms. A kind-8 operand
//! would need `i64.*` and a fold site that knows the declared kind; it does
//! not, so an `integer(kind=8)` argument is counted as 32 bits.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// The operand arrives as whatever produced it — an `F64` for a literal, an
/// `I32` once a bitwise op has touched it. `i32.or` against zero is ECMA
/// ToInt32 (§7.1.6) in the VM's dispatch, which is the same coercion every
/// other Fortran bit intrinsic already goes through on its way to `I32_AND`.
/// The counting opcodes themselves only `as_i32()`, which would truncate
/// rather than wrap.
// popcnt / leadz / trailz USED to be emitted here, each coercing to i32 and
// emitting one opcode. They are now `common:bits.*` — the shared bit family in
// `primitives/bits.rs`, which the `UnaryOp::PopCount` node also emits. Go had
// its own copy of the same three instructions in a different lane; the point of
// centralizing was that neither copy could see the other.

/// `parity(mask)` — true when an ODD number of the array's elements are true.
///
/// Not a bit intrinsic at all; F2008 just introduced it alongside them. It is
/// here rather than folded in the walker because a fold fires on the NAME
/// before anything knows whether the program defines its own `parity` — and one
/// in the suite does. A profile builtin loses to a user definition, which is
/// what an intrinsic is supposed to do.
///
/// Stack on entry: `[mask]`. Stack on exit: `[logical]`.
pub fn emit_fortran_parity(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    // `parity(mask, dim)` reduces along one dimension; that needs the same
    // machinery a `count(mask, dim)` would, and neither exists yet. Answering
    // for the whole array instead would be a wrong number, and returning null
    // is worse — it prints as `null` and the test passes on its exit code. So
    // the unimplemented form throws.
    if argc != 1 {
        for _ in 0..argc {
            chunk.emit_op(Op::DROP, line);
        }
        chunk.emit_string_const("parity(mask, dim) is not implemented", line);
        vybe_compiler::primitives::errors::emit_throw(chunk, line);
        return;
    }

    let mask_slot = chunk.alloc_scratch(1);
    let len_slot = chunk.alloc_scratch(1);
    let index_slot = chunk.alloc_scratch(1);
    let odd_slot = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, mask_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, mask_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, odd_slot, line);

    // BLOCK+LOOP: br_if 1 exits the block (depth 0 = LOOP, depth 1 = BLOCK).
    let block = chunk.emit_block(line);
    let (loop_pos, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, mask_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, odd_slot, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_XOR, line);
    chunk.emit_op_u16(Op::LOCAL_SET, odd_slot, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_pos);
    chunk.emit_end(line);
    chunk.patch_block(block);

    // The answer is a LOGICAL, so it has to leave as a `Value::Bool` rather
    // than as the i32 0/1 the toggle carries — an i32 prints as `0`/`1` where
    // a logical prints as a logical.
    chunk.emit_op_u16(Op::LOCAL_GET, odd_slot, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}
