//! Byte operations whose AST caller declares linear memory backing.

use vybe_runtime::{Chunk, Op};

/// Stack: [destination, source, destination byte stride, unsigned unit count]
/// -> [original destination]. Each unit is a little-endian 16-bit word.
/// Addresses advance with i32 wrapping arithmetic, just like scalar Wasm.
/// Reads and writes remain interleaved: overlap and trap-prefix writes must
/// not be replaced by an eager snapshot or a bulk bounds failure.
pub fn emit_copy_column_16(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let first = c.alloc_scratch(5);
    let (original, source, stride, count, destination) =
        (first, first + 1, first + 2, first + 3, first + 4);
    for slot in (first..first + 4).rev() {
        c.emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    c.emit_op_u16(Op::LOCAL_GET, original, line);
    c.emit_op_u16(Op::LOCAL_SET, destination, line);
    c.emit_op_u16(Op::LOCAL_GET, count, line);
    c.emit_op(Op::I64_TRUNC_F64_S, line);
    c.emit_op(Op::I32_WRAP_I64, line);
    c.emit_op_u16(Op::LOCAL_TEE, count, line);
    let nonempty = c.emit_if(line);
    for slot in [destination, source, stride] {
        c.emit_op_u16(Op::LOCAL_GET, slot, line);
        c.emit_op(Op::I64_TRUNC_F64_S, line);
        c.emit_op(Op::I32_WRAP_I64, line);
        c.emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    let (repeat, _) = c.emit_loop_s(line);
    c.emit_op_u16(Op::LOCAL_GET, destination, line);
    c.emit_op_u16(Op::LOCAL_GET, source, line);
    c.emit_op(Op::I32_LOAD16_U, line);
    c.emit_op(Op::I32_STORE16, line);
    for (cursor, step) in [(source, None), (destination, Some(stride))] {
        c.emit_op_u16(Op::LOCAL_GET, cursor, line);
        if let Some(step) = step {
            c.emit_op_u16(Op::LOCAL_GET, step, line);
        } else {
            c.emit_i32_const(2, line);
        }
        c.emit_op(Op::I32_ADD, line);
        c.emit_op_u16(Op::LOCAL_SET, cursor, line);
    }
    c.emit_op_u16(Op::LOCAL_GET, count, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_SUB, line);
    c.emit_op_u16(Op::LOCAL_TEE, count, line);
    c.emit_br_if(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(nonempty);
    c.emit_op_u16(Op::LOCAL_GET, original, line);
}

/// Stack: [destination, source, destination byte stride, unsigned count]
/// -> [original destination]. Copy in source order, including overlapping ranges.
/// Unlike memory.copy, this deliberately preserves scalar forward-copy semantics.
pub fn emit_copy_column(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let first = c.alloc_scratch(5);
    let (original, source, stride, count, destination) =
        (first, first + 1, first + 2, first + 3, first + 4);
    for slot in (first..first + 4).rev() {
        c.emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    c.emit_op_u16(Op::LOCAL_GET, original, line);
    c.emit_op_u16(Op::LOCAL_SET, destination, line);
    c.emit_op_u16(Op::LOCAL_GET, count, line);
    c.emit_op(Op::I64_TRUNC_F64_S, line);
    c.emit_op(Op::I32_WRAP_I64, line);
    c.emit_op_u16(Op::LOCAL_TEE, count, line);
    let nonempty = c.emit_if(line);
    // C's number-backed arguments need word normalization once, not per byte.
    // A zero-length copy must not inspect or convert either pointer.
    for slot in [destination, source, stride] {
        c.emit_op_u16(Op::LOCAL_GET, slot, line);
        c.emit_op(Op::I64_TRUNC_F64_S, line);
        c.emit_op(Op::I32_WRAP_I64, line);
        c.emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    // Forward copies agree with memory.copy except when the destination starts
    // inside the source at a higher address. Subtraction avoids end overflow.
    c.emit_op_u16(Op::LOCAL_GET, stride, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_EQ, line);
    c.emit_op_u16(Op::LOCAL_GET, destination, line);
    c.emit_op_u16(Op::LOCAL_GET, source, line);
    c.emit_op(Op::I32_LE_U, line);
    c.emit_op_u16(Op::LOCAL_GET, destination, line);
    c.emit_op_u16(Op::LOCAL_GET, source, line);
    c.emit_op(Op::I32_SUB, line);
    c.emit_op_u16(Op::LOCAL_GET, count, line);
    c.emit_op(Op::I32_GE_U, line);
    c.emit_op(Op::I32_OR, line);
    c.emit_op(Op::I32_AND, line);
    // Strided and forward-dependent copies cannot use bulk memory. Avoid
    // checking their contiguous spans; scalar accesses still trap in order.
    let eligible = c.emit_if_i32(line);
    // A bulk trap is atomic; a scalar loop can write a prefix before trapping.
    // Prove both complete spans in i64 (including a full 4-GiB memory) before
    // selecting memory.copy, otherwise retain the original scalar execution.
    // Both spans have the same unsigned length, so only the greater start
    // needs checking. Keep the addition wide: an i32 end would wrap at 4 GiB.
    c.emit_op_u16(Op::LOCAL_GET, destination, line);
    c.emit_op_u16(Op::LOCAL_GET, source, line);
    c.emit_op_u16(Op::LOCAL_GET, destination, line);
    c.emit_op_u16(Op::LOCAL_GET, source, line);
    c.emit_op(Op::I32_GT_U, line);
    c.emit_op(Op::SELECT, line);
    c.emit_op(Op::I64_EXTEND_I32_U, line);
    c.emit_op_u16(Op::LOCAL_GET, count, line);
    c.emit_op(Op::I64_EXTEND_I32_U, line);
    c.emit_op(Op::I64_ADD, line);
    c.emit_op_idx(Op::MEMORY_SIZE, 0u32, line);
    c.emit_op(Op::I64_EXTEND_I32_U, line);
    c.emit_i64_const(65536, line);
    c.emit_op(Op::I64_MUL, line);
    c.emit_op(Op::I64_LE_U, line);
    c.emit_else(line);
    c.emit_i32_const(0, line);
    c.emit_end(line);
    c.patch_block(eligible);
    let contiguous = c.emit_if(line);
    c.emit_op_u16(Op::LOCAL_GET, destination, line);
    c.emit_op_u16(Op::LOCAL_GET, source, line);
    c.emit_op_u16(Op::LOCAL_GET, count, line);
    c.emit_op_idx(Op::MEMORY_COPY, 0u32, line);
    c.emit_leb_u32(0u32, line);
    c.emit_else(line);
    let (repeat, _) = c.emit_loop_s(line);
    c.emit_op_u16(Op::LOCAL_GET, destination, line);
    c.emit_op_u16(Op::LOCAL_GET, source, line);
    c.emit_op(Op::I32_LOAD8_U, line);
    c.emit_op(Op::I32_STORE8, line);
    for (cursor, step) in [(source, None), (destination, Some(stride))] {
        c.emit_op_u16(Op::LOCAL_GET, cursor, line);
        if let Some(step) = step {
            c.emit_op_u16(Op::LOCAL_GET, step, line);
        } else {
            c.emit_i32_const(1, line);
        }
        c.emit_op(Op::I32_ADD, line);
        c.emit_op_u16(Op::LOCAL_SET, cursor, line);
    }
    c.emit_op_u16(Op::LOCAL_GET, count, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_SUB, line);
    c.emit_op_u16(Op::LOCAL_TEE, count, line);
    c.emit_br_if(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(contiguous);
    c.emit_end(line);
    c.patch_block(nonempty);
    c.emit_op_u16(Op::LOCAL_GET, original, line);
}

#[cfg(test)]
mod span_proof_tests {
    use super::*;
    use vybe_runtime::{VM, Value};

    #[test]
    fn maximum_start_proof_matches_both_spans_at_memory32_edges() {
        let edges = [0u32, 1, 2, 65535, 65536, 65537, i32::MAX as u32,
            0x8000_0000, u32::MAX - 1, u32::MAX];
        for destination in edges {
            for source in edges {
                for count in edges {
                    for pages in [0u64, 1, 2, 32768, 65535, 65536] {
                        let limit = pages * 65536;
                        let original = destination as u64 + count as u64 <= limit
                            && source as u64 + count as u64 <= limit;
                        let combined = destination.max(source) as u64 + count as u64 <= limit;
                        assert_eq!(combined, original,
                            "dst={destination} src={source} count={count} pages={pages}");
                    }
                }
            }
        }
    }

    #[test]
    fn emitted_column_copy_preserves_forward_overlap_and_strides() {
        for source in [0usize, 1, 7, 16] {
            for destination in [0usize, 2, 8, 16] {
                for count in [0usize, 1, 4, 8] {
                    for stride in [0usize, 1, 2] {
                        let mut expected: Vec<u8> = (0..64).map(|i| i * 3 + 1).collect();
                        let mut chunk = Chunk::new("<script>");
                        chunk.memory_min_pages = vec![1];
                        for (address, byte) in expected.iter().enumerate() {
                            chunk.emit_i32_const(address as i32, 0);
                            chunk.emit_i32_const(*byte as i32, 0);
                            chunk.emit_op(Op::I32_STORE8, 0);
                        }
                        for argument in [destination, source, stride, count] {
                            chunk.emit_i32_const(argument as i32, 0);
                        }
                        emit_copy_column(std::slice::from_mut(&mut chunk), 0, 0);
                        chunk.emit_op(Op::DROP, 0);
                        for i in 0..count {
                            expected[destination + i * stride] = expected[source + i];
                        }
                        for (address, byte) in expected.iter().enumerate() {
                            chunk.emit_i32_const(address as i32, 0);
                            chunk.emit_op(Op::I32_LOAD8_U, 0);
                            chunk.emit_i32_const(*byte as i32, 0);
                            chunk.emit_op(Op::I32_NE, 0);
                            let mismatch = chunk.emit_if(0);
                            chunk.emit_op(Op::UNREACHABLE, 0);
                            chunk.emit_end(0);
                            chunk.patch_block(mismatch);
                        }
                        chunk.emit_i32_const(1, 0);
                        chunk.emit_op(Op::RETURN, 0);
                        assert_eq!(VM::new().run(vec![chunk]).unwrap(), Value::I32(1));
                    }
                }
            }
        }
    }
}
