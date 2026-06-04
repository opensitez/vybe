//! Shared generator / continuation emission.
//!
//! Source languages spell generators differently (`function*`, Python
//! `yield`, VB `Yield`, C# `yield return`, PHP `Generator`), but this module is
//! the single compiler-side surface for WebAssembly stack-switching generator
//! opcodes.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

/// `suspend $tag`: yield from the current continuation.
///
/// Stack before: `[yield_value]`
/// Stack after resume: `[resume_value]`
pub fn emit_suspend(chunk: &mut Chunk, line: u32) {
    emit_suspend_tagged(chunk, 0, line);
}

/// Tagged `suspend $tag` for language proposals that carry an explicit
/// continuation tag index.
pub fn emit_suspend_tagged(chunk: &mut Chunk, tag: u16, line: u32) {
    chunk.emit_op_u16(Op::SUSPEND, tag, line);
}

/// Tagged `resume $cont`.
pub fn emit_resume_tagged(chunk: &mut Chunk, tag: u16, line: u32) {
    chunk.emit_op_u16(Op::RESUME, tag, line);
}

/// Tagged `resume_throw`.
pub fn emit_resume_throw_tagged(chunk: &mut Chunk, tag: u16, line: u32) {
    chunk.emit_op_u16(Op::RESUME_THROW, tag, line);
}

/// `resume $cont`: resume a continuation with a value.
///
/// Stack before: `[continuation, resume_value]`
/// Stack after: `[yielded_or_returned_value]`
pub fn emit_resume(chunk: &mut Chunk, line: u32) {
    emit_resume_tagged(chunk, 0, line);
}

/// `resume_throw`: resume a continuation by throwing into it.
///
/// Stack before: `[continuation, exception_value]`
pub fn emit_resume_throw(chunk: &mut Chunk, line: u32) {
    emit_resume_throw_tagged(chunk, 0, line);
}

/// Generator iterator advance.
///
/// Stack before: `[continuation]`
/// Stack after: `[value, has_more_i32]`
pub fn emit_next(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::GEN_NEXT, line);
}

/// Drain a generator continuation into a new array using pure WASM opcodes.
///
/// Stack before: `[continuation]`
/// Stack after:  `[array_of_yielded_values]`
///
/// Zero host imports. Zero stdlib. Every instruction is a WASM opcode:
///   - `ARRAY_NEW_FIXED 0`  → empty array
///   - `GEN_NEXT`           → (value, has_more_i32)
///   - `I32_EQZ + BR_IF`   → exit when done
///   - `DUP + ARRAY_LENGTH + ARRAY_SET` → push (auto-extends; VM Object::set)
///
/// Works identically for all languages — same path compile_generator_for_in uses.
/// Both the two-chunk (_into) and chunks/current variants delegate here.
fn emit_drain_loop(cont_slot: u16, result_slot: u16, val_slot: u16, chunk: &mut Chunk, line: u32) {
    // Structure: block { loop { GEN_NEXT; if done { drop val; br 2 (exit block) }; push; br 0 } }
    //
    // Label stack depths from inside the `if` body:
    //   br 0 → exits `if` (falls through)
    //   br 1 → restarts `loop`
    //   br 2 → exits outer `block`  ← use this to break
    //
    // GEN_NEXT pushes [value] then [has_more_i32] (has_more on top).
    // When done: (null, 0). When live: (yielded_value, 1).
    let block_p = chunk.emit_block(line);
    let (loop_p, _) = chunk.emit_loop_s(line);

    chunk.emit_op_u16(Op::LOCAL_GET, cont_slot, line);
    emit_next(chunk, line);
    // Stack: [value, has_more_i32]

    // Stash has_more in val_slot temporarily so we can check it after clearing it from the stack
    chunk.emit_op_u16(Op::LOCAL_SET, val_slot, line); // val_slot = has_more (top)
    chunk.emit_op(Op::DROP, line);
    // Stack: [value]
    chunk.emit_op_u16(Op::LOCAL_GET, val_slot, line); // restore has_more to TOS
    // Stack: [value, has_more]

    chunk.emit_op(Op::I32_EQZ, line);    // 1 if done (has_more==0)
    chunk.emit_if(line);                 // if done {
    chunk.emit_op(Op::DROP, line);       //   drop dangling value (was below has_more)
    chunk.emit_br(2, line);              //   br 2 → exit outer block
    chunk.emit_end(line);                // } end if
    // Stack: [value]  — only reached when has_more=1

    chunk.emit_op_u16(Op::LOCAL_SET, val_slot, line); // val_slot = yielded value
    chunk.emit_op(Op::DROP, line);

    // result[result.length] = val  (ARRAY_SET auto-extends via Object::set)
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);               // i32 length = next index
    chunk.emit_op_u16(Op::LOCAL_GET, val_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);                  // pushes val back
    chunk.emit_op(Op::DROP, line);                       // drop it (stack must be clean for br 0)

    chunk.emit_br(0, line);              // restart loop
    chunk.emit_end(line);
    chunk.patch_loop(loop_p);
    chunk.emit_end(line);
    chunk.patch_block(block_p);
}

/// Two-chunk variant for stdlib.rs `build_*` functions.
/// `imports` is unused (no host imports needed) but kept for API symmetry.
/// Called as a arity-1 function; `code.local_count` must be 1 on entry (arg 0
/// is the continuation, already in local slot 0 by the VM's calling convention).
pub fn emit_drain_into_array_into(_imports: &mut Chunk, code: &mut Chunk, line: u32) {
    // Arg 0 (local slot 0) IS the continuation — no copy needed.
    let cont_slot = 0u16;
    let result_slot = code.local_count;
    code.local_count += 1;
    let val_slot = code.local_count;
    code.local_count += 1;

    code.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line); // empty array
    code.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    code.emit_op(Op::DROP, line);

    emit_drain_loop(cont_slot, result_slot, val_slot, code, line);

    code.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `chunks/current` variant for use inside the `Compiler`.
pub fn emit_drain_into_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let cont_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let result_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let val_slot = chunks[current].local_count;
    chunks[current].local_count += 1;

    chunks[current].emit_op_u16(Op::LOCAL_SET, cont_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line); // empty array
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    emit_drain_loop(cont_slot, result_slot, val_slot, &mut chunks[current], line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}
