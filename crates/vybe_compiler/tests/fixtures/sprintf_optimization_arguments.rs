// Captured argument packing before direct stack-array construction.
use vybe_runtime::{Chunk, Op};
pub fn emit_sprintf(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    // Ensure the helper chunk is in the chunk list.
    let helper_idx = if let Some(idx) = chunks.iter().position(|c| c.name == "__fmt_sprintf") {
        idx
    } else {
        let helper = vybe_compiler::primitives::sprintf::build_sprintf(&mut chunks[0]);
        chunks.push(helper);
        chunks.len() - 1
    };

    // The helper expects (fmt, args_array).  The compiler pushed the args
    // on the stack in order: [fmt, arg0, arg1, ...].
    // We need to pack args 1..N into an array first.

    let nargs = argc as i32; // total args including fmt
    let nrest = nargs - 1; // args after fmt

    // Pop args in reverse, store in temporaries, then rebuild.
    // Simpler: all args are already on the stack in forward order.
    // Strategy:
    //   1. Store all args into locals (right-to-left to keep stack sane).
    //   2. Build one fixed-size array from the stored locals.
    //   3. Call the sprintf helper with (fmt, args_array).

    // Allocate temp locals for the variadic args.
    let first_arg_slot = chunks[current].alloc_scratch(nrest.max(0) as u16);
    let fmt_slot = chunks[current].alloc_scratch(1);

    // Store variadic args (they are on top of stack, in order arg0..argN-1).
    // Stack order: ... fmt arg0 arg1 ... argN-1  (argN-1 on top)
    // Store from top to bottom.
    for k in (0..nrest).rev() {
        let slot = first_arg_slot + k as u16;
        chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    }
    // Now fmt is on top — store it.
    chunks[current].emit_op_u16(Op::LOCAL_SET, fmt_slot, line);

    // Call the helper directly by function reference.
    chunks[current].emit_op_u16(Op::REF_FUNC, helper_idx as u16, line);
    chunks[current].emit(0u8, line); // upvalue count
    chunks[current].emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    for k in 0..nrest {
        chunks[current].emit_op_u16(Op::LOCAL_GET, first_arg_slot + k as u16, line);
    }
    vybe_compiler::primitives::collections::emit_array_new(
        chunks,
        current,
        nrest.max(0) as u16,
        line,
    );
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(&mut chunks[current], 2, line);
}
