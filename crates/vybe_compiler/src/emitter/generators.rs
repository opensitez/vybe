//! Shared generator / continuation emission.
//!
//! Source languages spell generators differently (`function*`, Python
//! `yield`, VB `Yield`, C# `yield return`, PHP `Generator`), but this module is
//! the single compiler-side surface for WebAssembly stack-switching generator
//! opcodes.

use std::sync::Arc;
use vybe_bytecode::chunk::StackSwitchHandler;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use crate::emitter::collections;
use crate::emitter::ops;

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

/// Tag carried by a generator `yield` `suspend` (see `emit_suspend`). The
/// `resume` handler emitted by `emit_next` keys on this tag so a yield is
/// routed to the "more values" path; anything else (a generator that simply
/// returns) falls through `resume` as completion.
const GEN_YIELD_TAG: u16 = 0;

/// Generator iterator advance — spec WASM stack-switching, no custom opcode.
///
/// Stack before: `[continuation]`
/// Stack after:  `[value, has_more_i32]`
///
/// This is the spec-compliant replacement for the former VM-internal
/// `GEN_NEXT` (0xFF) opcode. It drives one step of the generator with the
/// stack-switching `resume` instruction plus an `(on $yield -> handler)`
/// handler vector — exactly the proposal's iteration shape:
///
///   - The generator `suspend`s (tag 0 = yield): the VM routes control to this
///     `resume`'s recorded handler offset with the yielded value on the stack;
///     we push `has_more = 1`.
///   - The generator returns (completes): `resume` falls through with the
///     return value; we push `has_more = 0`.
///
/// Both arms converge through a result-2 `block` so the `[value, has_more]`
/// stack contract the drivers below consume is unchanged.
pub fn emit_next(chunk: &mut Chunk, line: u32) {
    // The continuation must not straddle the BLOCK boundary: the VM records a
    // block's entry stack height, and `resume` consumes the continuation inside
    // the block, which would desync that height. Park it in a fresh local and
    // reload inside the block, leaving the block stack-neutral on entry.
    let cont_slot = chunk.local_count;
    chunk.local_count += 1;
    chunk.emit_op_u16(Op::LOCAL_SET, cont_slot, line); // LOCAL_SET tees the value…
    chunk.emit_op(Op::DROP, line); // …DROP clears it → [cont] becomes []

    let done_key = chunk.add_constant(Value::String(Arc::from("__gen_done")));

    // block (result: value, has_more_i32)
    let block_p = chunk.emit_block_typed(line, 2);

    // Done guard. The spec `resume` instruction TRAPS on a completed
    // continuation, but the JS iterator protocol requires `.next()` after
    // completion to keep returning `{value: undefined, done: true}`. So we
    // stamp `__gen_done` on the continuation when it completes (below) and
    // short-circuit here before ever resuming a finished generator.
    chunk.emit_op_u16(Op::LOCAL_GET, cont_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, done_key, line); // null if never set, else I32(1)
    chunk.emit_op(Op::REF_IS_NULL, line); // i32: 1 if not-yet-done
    chunk.emit_op(Op::I32_EQZ, line); // i32: 1 if already done
    let done_if = chunk.emit_if(line); // if already done {
    chunk.emit_op(Op::NULL, line);
    chunk.emit_op(Op::I32_CONST_0, line); // [null, 0]
    chunk.emit_br(1, line); //   br $block → exit with (undefined, done)
    chunk.emit_end(line); // }
    chunk.patch_block(done_if);

    chunk.emit_op_u16(Op::LOCAL_GET, cont_slot, line); // [cont]
    chunk.emit_op(Op::NULL, line); // resume value → [cont, null]
    let resume_ip = chunk.code.len();
    chunk.emit_op_u16(Op::RESUME, GEN_YIELD_TAG, line);
    // Completion arm (generator returned): stack = [return_value]. Stamp the
    // continuation done so the next call hits the guard above instead of the
    // `resume`-on-completed trap.
    chunk.emit_op_u16(Op::LOCAL_GET, cont_slot, line); // [retval, cont]
    chunk.emit_op(Op::I32_CONST_1, line); // [retval, cont, 1]
    chunk.emit_op_u16(Op::STRUCT_SET, done_key, line); // [retval, 1]
    chunk.emit_op(Op::DROP, line); // [retval]
    chunk.emit_op(Op::I32_CONST_0, line); // has_more = 0 → [retval, 0]
    chunk.emit_br(0, line); // br $block → converge at END, skipping the yield arm
    // Yield arm: the VM jumps here (ip = handler_ip) with [yielded_value]
    let handler_ip = chunk.code.len();
    chunk.emit_op(Op::I32_CONST_1, line); // has_more = 1 → [yieldval, 1]
    chunk.emit_end(line); // end $block → [value, has_more]
    chunk.patch_block(block_p);

    chunk.stack_switch_handlers.insert(
        resume_ip,
        vec![StackSwitchHandler {
            kind: 0, // on-tag-to-label
            tag_index: GEN_YIELD_TAG as u32,
            label_index: handler_ip as u32,
        }],
    );
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

    chunk.emit_op(Op::I32_EQZ, line); // 1 if done (has_more==0)
    chunk.emit_if(line); // if done {
    chunk.emit_op(Op::DROP, line); //   drop dangling value (was below has_more)
    chunk.emit_br(2, line); //   br 2 → exit outer block
    chunk.emit_end(line); // } end if
    // Stack: [value]  — only reached when has_more=1

    chunk.emit_op_u16(Op::LOCAL_SET, val_slot, line); // val_slot = yielded value
    chunk.emit_op(Op::DROP, line);

    // result[result.length] = val  (ARRAY_SET auto-extends via Object::set)
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line); // i32 length = next index
    chunk.emit_op_u16(Op::LOCAL_GET, val_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line); // pushes val back
    chunk.emit_op(Op::DROP, line); // drop it (stack must be clean for br 0)

    chunk.emit_br(0, line); // restart loop
    chunk.emit_end(line);
    chunk.patch_loop(loop_p);
    chunk.emit_end(line);
    chunk.patch_block(block_p);
}

/// Two-chunk variant for runtime helper builder functions.
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

/// Take up to `limit` yielded values from a generator continuation into a new array.
///
/// Stack before: `[continuation, limit]`
/// Stack after:  `[array_of_yielded_values]`
///
/// This is the bounded counterpart to `emit_drain_into_array`: it is safe for
/// infinite generators and keeps source-language iterator helpers on the same
/// WASM continuation opcodes (`GEN_NEXT`) instead of asking host functions to
/// resume continuations.
pub fn emit_take_into_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let limit_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let cont_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let result_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let count_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let value_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let has_more_slot = chunks[current].local_count;
    chunks[current].local_count += 1;

    chunks[current].emit_op_u16(Op::LOCAL_SET, limit_slot, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cont_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op(Op::I32_CONST_0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    let block_p = chunks[current].emit_block(line);
    let (loop_p, _) = chunks[current].emit_loop_s(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, limit_slot, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cont_slot, line);
    emit_next(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, has_more_slot, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, has_more_slot, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunks[current].emit_op(Op::I32_CONST_1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_p);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_p);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Flat-map an array with a mapper that returns a generator continuation.
///
/// Stack before: `[source_array, mapper_fn]`
/// Stack after:  `[array_of_all_yielded_values]`
///
/// This keeps generator-returning iterator helpers in bytecode/continuation
/// space. Hosts can invoke callbacks, but they cannot resume a WASM
/// continuation, so the returned continuation is drained here via `GEN_NEXT`.
pub fn emit_flat_map_generator_mapper_into_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let mapper_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let source_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let result_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let i_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let len_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let item_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let cont_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let value_slot = chunks[current].local_count;
    chunks[current].local_count += 1;
    let has_more_slot = chunks[current].local_count;
    chunks[current].local_count += 1;

    chunks[current].emit_op_u16(Op::LOCAL_SET, mapper_slot, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op(Op::I32_CONST_0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, mapper_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cont_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    let inner_block = chunks[current].emit_block(line);
    let (inner_loop, _) = chunks[current].emit_loop_s(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cont_slot, line);
    emit_next(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, has_more_slot, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, has_more_slot, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(inner_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(inner_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op(Op::I32_CONST_1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}
