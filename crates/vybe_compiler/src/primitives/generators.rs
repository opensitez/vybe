//! Shared generator / continuation emission.
//!
//! Source languages spell generators differently (`function*`, Python
//! `yield`, VB `Yield`, C# `yield return`, PHP `Generator`), but this module is
//! the single compiler-side surface for WebAssembly stack-switching generator
//! opcodes.

use crate::primitives::instructions::{core_wasm, host, recipes};
use std::sync::Arc;
use vybe_runtime::chunk::StackSwitchHandler;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use crate::primitives::collections;
use crate::primitives::ops;
// For the buffered-iterator façade at the end of this module, which is an
// `impl Compiler` rather than free emitters.
use crate::primitives::Compiler;
use vybe_ast::Expression;

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
    chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, cont_slot, line);

    let done_key = chunk.add_constant(Value::String(Arc::from("__gen_done")));

    // block (result: value, has_more_i32)
    let block_p = chunk.emit_block_typed(line, 2);

    // Done guard. The spec `resume` instruction TRAPS on a completed
    // continuation, but the JS iterator protocol requires `.next()` after
    // completion to keep returning `{value: undefined, done: true}`. So we
    // stamp `__gen_done` on the continuation when it completes (below) and
    // short-circuit here before ever resuming a finished generator.
    chunk.emit_op_u16(Op::LOCAL_GET, cont_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, done_key, line); // null if never set, else I32(1)
    chunk.emit_op(Op::REF_IS_NULL, line); // i32: 1 if not-yet-done
    chunk.emit_op(Op::I32_EQZ, line); // i32: 1 if already done
    let done_if = chunk.emit_if(line); // if already done {
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    core_wasm::i32_const(chunk, line, 0); // [null, 0]
    chunk.emit_br(1, line); //   br $block → exit with (undefined, done)
    chunk.emit_end(line); // }
    chunk.patch_block(done_if);

    chunk.emit_op_u16(Op::LOCAL_GET, cont_slot, line); // [cont]
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line); // resume value → [cont, null]
    let resume_ip = chunk.code.len();
    chunk.emit_op_u16(Op::RESUME, GEN_YIELD_TAG, line);
    // Completion arm (generator returned): stack = [return_value]. Stamp the
    // continuation done so the next call hits the guard above instead of the
    // `resume`-on-completed trap.
    chunk.emit_op_u16(Op::LOCAL_GET, cont_slot, line); // [retval, cont]
    core_wasm::i32_const(chunk, line, 1); // [retval, cont, 1]
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, done_key, line); // [retval, 1]
    core_wasm::i32_const(chunk, line, 0); // has_more = 0 → [retval, 0]
    chunk.emit_br(0, line); // br $block → converge at END, skipping the yield arm
    // Yield arm: the VM jumps here (ip = handler_ip) with [yielded_value]
    let handler_ip = chunk.code.len();
    core_wasm::i32_const(chunk, line, 1); // has_more = 1 → [yieldval, 1]
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

    // result[result.length] = val  (ARRAY_SET auto-extends via Object::set)
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line); // i32 length = next index
    chunk.emit_op_u16(Op::LOCAL_GET, val_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line); // pushes val back

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
    code.alloc_scratch(1);
    let val_slot = code.local_count;
    code.alloc_scratch(1);

    code.emit_array_new_fixed(0, 0, line); // empty array
    code.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    emit_drain_loop(cont_slot, result_slot, val_slot, code, line);

    code.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `chunks/current` variant for use inside the `Compiler`.
pub fn emit_drain_into_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let cont_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let val_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, cont_slot, line);

    chunks[current].emit_array_new_fixed(0, 0, line); // empty array
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

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
    chunks[current].alloc_scratch(1);
    let cont_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let count_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let has_more_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, limit_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cont_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);

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
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

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
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);

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
    chunks[current].alloc_scratch(1);
    let source_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let i_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let item_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let cont_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);
    let has_more_slot = chunks[current].local_count;
    chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, mapper_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

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

    chunks[current].emit_op_u16(Op::LOCAL_GET, mapper_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cont_slot, line);

    let inner_block = chunks[current].emit_block(line);
    let (inner_loop, _) = chunks[current].emit_loop_s(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cont_slot, line);
    emit_next(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, has_more_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

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
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Generator entry control dispatch — checks if the control parameter is a
/// generator control packet (`{__vybe_generator_control: true, op: "throw"|"return", value}`).
///
/// Emitted once at the start of every generator body. `emit_return` callback
/// lets the Compiler inject finally handling for the return path.
pub fn emit_entry_control(
    chunk: &mut Chunk,
    control_slot: u16,
    line: u32,
    emit_return: &mut dyn FnMut(&mut Chunk, u32),
) {
    chunk.emit_op_u16(Op::LOCAL_GET, control_slot, line);
    recipes::is_object(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, control_slot, line);
    let marker_key = chunk.add_constant(Value::String(Arc::from("__vybe_generator_control")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, marker_key, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, control_slot, line);
    let op_key = chunk.add_constant(Value::String(Arc::from("op")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, op_key, line);
    chunk.emit_string_const("throw", line);
    host::emit(chunk, "wasm:js-string", "equals", 2, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, control_slot, line);
    let value_key = chunk.add_constant(Value::String(Arc::from("value")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, value_key, line);
    crate::primitives::errors::emit_throw(chunk, line);

    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, control_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, op_key, line);
    chunk.emit_string_const("return", line);
    host::emit(chunk, "wasm:js-string", "equals", 2, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, control_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, value_key, line);
    emit_return(chunk, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// Generator resume value dispatch — at each yield resume point, checks if
/// the resume value is a control packet and handles throw/return.
///
/// Returns (resume_slot, result_slot) for the caller.
pub fn emit_resume_dispatch(
    chunk: &mut Chunk,
    line: u32,
    emit_return: &mut dyn FnMut(&mut Chunk, u32),
) -> (u16, u16) {
    let resume_slot = chunk.local_count;
    chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, resume_slot, line);

    let result_slot = chunk.local_count;
    chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, resume_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, resume_slot, line);
    recipes::is_object(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, resume_slot, line);
    let marker_key = chunk.add_constant(Value::String(Arc::from("__vybe_generator_control")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, marker_key, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, resume_slot, line);
    let op_key = chunk.add_constant(Value::String(Arc::from("op")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, op_key, line);
    chunk.emit_string_const("throw", line);
    host::emit(chunk, "wasm:js-string", "equals", 2, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, resume_slot, line);
    let value_key = chunk.add_constant(Value::String(Arc::from("value")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, value_key, line);
    crate::primitives::errors::emit_throw(chunk, line);

    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, resume_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, op_key, line);
    chunk.emit_string_const("return", line);
    host::emit(chunk, "wasm:js-string", "equals", 2, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, resume_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, value_key, line);
    emit_return(chunk, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);

    (resume_slot, result_slot)
}

/// Drain any iterable into an array using the ECMA-262 iterator protocol.
///
/// Implements §7.4.2 GetIterator + §7.4.4 IteratorNext + §7.4.5 IteratorStep:
///   1. `iter_fn = obj[Symbol.iterator]`  (STRUCT_GET "Symbol(@@iterator)")
///   2. `iter = iter_fn(obj)`             (CALL_REF 1 — passes receiver)
///   3. `next_fn = iter.next`             (STRUCT_GET "next")
///   4. loop: `step = next_fn(iter)` → check `step.done` → push `step.value`
///
/// TypeRegistry resolves `Symbol(@@iterator)` for built-in types:
///   Array → ecma:array.values (§23.1.3.38)
///   Map   → ecma:map.entries  (§24.1.3.12)
///   Set   → ecma:set.values   (§24.2.3.11)
///   String→ ecma:string.iterator (§22.1.5.1)
/// Custom classes define `[Symbol.iterator]()` as an own property.
///
/// Stack before: [iterable]
/// Stack after:  [array_of_values]
/// §7.4.2 GetIterator(obj, ASYNC) for `for await`: if the object in
/// `iter_slot` is not already an iterator (no `next`), resolve its
/// `[Symbol.asyncIterator]` method (walker fallback key `asyncIterator`)
/// and CALL it with the object as receiver, storing the returned async
/// iterator back into `iter_slot`. An async-generator method returns a
/// generator continuation, so the for-of runtime-generator gate then
/// drives it lazily via WASM stack-switching. Leaves `iter_slot`
/// untouched when no async-iterator method exists (later gates handle
/// sync iterables / array-likes).
pub fn emit_resolve_async_iterator(
    chunks: &mut [Chunk],
    current: usize,
    iter_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let fn_slot = chunk.alloc_scratch(1);

    // Already an iterator (`next` present) → leave as-is (§25.1.2 fast path).
    let next_key = chunk.add_constant(vybe_runtime::Value::String(Arc::from("next")));
    chunk.emit_op_u16(Op::LOCAL_GET, iter_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, next_key, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);

    // fn = obj[Symbol(@@asyncIterator)] ?? obj.asyncIterator
    let async_sym = chunk.add_constant(vybe_runtime::Value::String(Arc::from(
        "Symbol(@@asyncIterator)",
    )));
    chunk.emit_op_u16(Op::LOCAL_GET, iter_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, async_sym, line);
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    let async_key = chunk.add_constant(vybe_runtime::Value::String(Arc::from("asyncIterator")));
    chunk.emit_op_u16(Op::LOCAL_GET, iter_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, async_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunk.emit_op(Op::END, line);

    // fn non-null → call it; ADOPT the result only when it is a generator
    // continuation (async-generator method), which the runtime-generator
    // gate drives lazily with per-step await. A plain `{ async next() }`
    // iterator object stays on the original iterable — the host iterForOf
    // fallback drives that shape (replacing it here would strand it).
    chunk.emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, iter_slot, line);
    crate::primitives::globals::emit_write(chunk, "__js_this", line);
    chunk.emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, iter_slot, line);
    chunk.emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    let result_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    let is_gen_idx = chunk.add_import("ecma:value", "isGenerator");
    chunk.emit_call(is_gen_idx, 1, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, iter_slot, line);
    chunk.emit_op(Op::END, line);
    chunk.emit_op(Op::END, line);

    chunk.emit_op(Op::END, line);
}

pub fn emit_drain_custom_iterable(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_drain_iterable_inner(chunks, current, line, false);
}

pub fn emit_drain_async_iterable(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_drain_iterable_inner(chunks, current, line, true);
}

fn emit_drain_iterable_inner(chunks: &mut [Chunk], current: usize, line: u32, async_iter: bool) {
    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    let iter_fn_slot = chunk.alloc_scratch(1);
    let iter_slot = chunk.alloc_scratch(1);
    let next_fn_slot = chunk.alloc_scratch(1);
    let step_slot = chunk.alloc_scratch(1);
    let result_slot = chunk.alloc_scratch(1);
    let done_slot = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    // result = []
    chunk.emit_array_new_fixed(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    // Outer block — BR 0 exits with result_slot on stack
    let outer_block = chunk.emit_block(line);

    // §25.1.2: If obj already has "next", it IS the iterator
    // (%IteratorPrototype%[@@iterator]() returns this).
    // Skip the [Symbol.iterator] call to avoid infinite chains
    // (e.g. ArrayIterator has ObjectKind::Array → TypeRegistry
    // would resolve values() as its @@iterator, creating a loop).
    let next_key = chunk.add_constant(vybe_runtime::Value::String(Arc::from("next")));
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, next_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, next_fn_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, next_fn_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);

    // No "next" → get iterator method
    // §7.4.2 GetIterator: for async, try @@asyncIterator first
    if async_iter {
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        let async_sym = chunk.add_constant(vybe_runtime::Value::String(Arc::from(
            "Symbol(@@asyncIterator)",
        )));
        chunk.emit_struct_field_op(Op::STRUCT_GET, 0, async_sym, line);
        chunk.emit_op_u16(Op::LOCAL_SET, iter_fn_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, iter_fn_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        let async_key = chunk.add_constant(vybe_runtime::Value::String(Arc::from("asyncIterator")));
        chunk.emit_struct_field_op(Op::STRUCT_GET, 0, async_key, line);
        chunk.emit_op_u16(Op::LOCAL_SET, iter_fn_slot, line);
        chunk.emit_end(line);

        chunk.emit_op_u16(Op::LOCAL_GET, iter_fn_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if(line);
    }
    // Sync: try @@iterator
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let sym_key = chunk.add_constant(vybe_runtime::Value::String(Arc::from("Symbol(@@iterator)")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, sym_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, iter_fn_slot, line);

    // Fallback: walker-normalized "iterator" key for custom classes
    chunk.emit_op_u16(Op::LOCAL_GET, iter_fn_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let iter_key = chunk.add_constant(vybe_runtime::Value::String(Arc::from("iterator")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, iter_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, iter_fn_slot, line);
    chunk.emit_end(line);
    if async_iter {
        chunk.emit_end(line); // close the async-first if
    }

    // No iterator → skip to end of this IF; the iterForOf fallback
    // after the IF handles array-likes and other non-iterables.
    chunk.emit_op_u16(Op::LOCAL_GET, iter_fn_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_br_if(0, line); // BR to end of IF(no next), not outer_block

    // iter = iter_fn(obj) — §7.4.2 step 4: Call(method, obj)
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    crate::primitives::globals::emit_write(chunk, "__js_this", line);
    chunk.emit_op_u16(Op::LOCAL_GET, iter_fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, iter_slot, line);

    // iter null → exit
    chunk.emit_op_u16(Op::LOCAL_GET, iter_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_br_if(1, line); // BR outer_block

    // Get next from the new iterator
    chunk.emit_op_u16(Op::LOCAL_GET, iter_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, next_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, next_fn_slot, line);

    chunk.emit_else(line);

    // obj already has "next" — it IS the iterator (§25.1.2)
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, iter_slot, line);

    chunk.emit_end(line);

    // no next → pure WASM array-like fallback (§23.1.2.1 step 5)
    // Check for "length" property and iterate by numeric index.
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, next_fn_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);

    let length_key = chunk.add_constant(vybe_runtime::Value::String(Arc::from("length")));
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, length_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, done_slot, line); // reuse done_slot for length
    chunk.emit_op_u16(Op::LOCAL_GET, done_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_br_if(1, line); // no length → exit outer_block (this-if=0, outer_block=1)

    // Loop from 0 to length
    let idx_slot = step_slot; // reuse step_slot for index counter
    chunk.emit_array_new_fixed(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    let al_block = chunk.emit_block(line);
    let (al_loop, _) = chunk.emit_loop_s(line);
    // if idx >= length → break
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, done_slot, line);
    chunk.emit_op(Op::F64_GE, line);
    chunk.emit_br_if(1, line); // exit al_block
    // result.push(obj[idx]) — ARRAY_GET handles string coercion
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    crate::primitives::collections::emit_push(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    // idx++
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunk.emit_br(0, line); // continue loop
    chunk.emit_end(line);
    chunk.patch_loop(al_loop);
    chunk.emit_end(line);
    chunk.patch_block(al_block);
    chunk.emit_br(1, line); // exit outer_block with populated result
    // Nesting: outer_block(1) → this-if(0). br(1) = outer_block.

    chunk.emit_end(line); // end if(no next)

    // §7.4.5 IteratorStep loop
    let block_p = chunk.emit_block(line);
    let (loop_p, _) = chunk.emit_loop_s(line);

    // __js_this = iter; step = next_fn(iter)
    chunk.emit_op_u16(Op::LOCAL_GET, iter_slot, line);
    crate::primitives::globals::emit_write(chunk, "__js_this", line);
    chunk.emit_op_u16(Op::LOCAL_GET, next_fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, iter_slot, line);
    chunk.emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    if async_iter {
        // §25.5.2.1: Await IteratorResult from async next()
        super::functions::emit_await(chunk, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, step_slot, line);

    // step null → break
    chunk.emit_op_u16(Op::LOCAL_GET, step_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_br_if(1, line);

    // done = step.done; if truthy → break
    chunk.emit_op_u16(Op::LOCAL_GET, step_slot, line);
    let done_key = chunk.add_constant(vybe_runtime::Value::String(Arc::from("done")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, done_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, done_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, done_slot, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);

    // result.push(step.value)
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, step_slot, line);
    let value_key = chunk.add_constant(vybe_runtime::Value::String(Arc::from("value")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, value_key, line);
    crate::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_p);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_p);

    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

// ── Generator driver chunks (`it.next` protocol methods) ──────────────
//
// Whole-chunk builders for the `__stdlib_generator_next` /
// `__stdlib_async_generator_next` drivers the VM attaches to every
// continuation (`attach_continuation_protocols`). They compose the
// recipe emitters above; `runtime_helpers.rs` only registers them.

/// Sync driver — drive one step, return the raw `{value, done}`
/// IteratorResult (§27.5.1.2).
pub fn build_generator_next(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_generator_next");
    c.arity = 0;
    c.local_count = 2; // value(0) + has_more(1)
    let value_local = 0u16;
    let has_more_local = 1u16;
    let value_key = c.add_constant(Value::String(Arc::from("value")));
    let done_key = c.add_constant(Value::String(Arc::from("done")));

    crate::primitives::globals::emit_read(&mut c, "__js_this", 0);
    emit_next(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, has_more_local, 0);
    c.emit_op_u16(Op::LOCAL_SET, value_local, 0);

    c.emit_struct_new(0, 0, 0);
    c.emit_dup(0);
    c.emit_op_u16(Op::LOCAL_GET, value_local, 0);
    c.emit_struct_field_op(Op::STRUCT_SET, 0, value_key, 0);
    c.emit_dup(0);
    c.emit_op_u16(Op::LOCAL_GET, has_more_local, 0);
    ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_struct_field_op(Op::STRUCT_SET, 0, done_key, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

/// §27.6.1.2 AsyncGenerator.prototype.next(v) — drive one step and
/// return a PROMISE of the IteratorResult. Mirrors the compiler's
/// inline sync fast path (primitives/calls.rs `.next()`), plus the async
/// deltas:
///   - `next()`  → spec `resume` + `(on yield)` drive (`emit_next`);
///   - `next(v)` → RESUME with v (the suspended `yield` evaluates to
///     v), spec `done` via `ecma:value.isGeneratorDone`;
///   - a prior `.return()` stamped `__vybe_gen_returned` → short-
///     circuit to `{undefined, true}` (§27.5.1.2 step 2 mirror);
///   - Await(value) before delivery — async bodies hand back promises
///     (e.g. the wrapped return value); a rejection rejects the result;
///   - completion with no explicit return leaves null → deliver
///     undefined (§27.5.3.5); `done` is a real Boolean;
///   - any throw out of the body REJECTS (`ecma:promise.reject`)
///     instead of throwing synchronously into the caller.
pub fn build_async_generator_next(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_async_generator_next");
    c.arity = 1; // optional resume value; missing arg pads as Undefined
    c.local_count = 4; // v(0) + value(1) + done_i32(2) + err(3)
    let v_local = 0u16;
    let value_local = 1u16;
    let done_local = 2u16;
    let err_local = 3u16;
    let value_key = c.add_constant(Value::String(Arc::from("value")));
    let done_key = c.add_constant(Value::String(Arc::from("done")));
    let returned_key = c.add_constant(Value::String(Arc::from("__vybe_gen_returned")));
    let started_key = c.add_constant(Value::String(Arc::from("__vybe_gen_started")));
    // Per-chunk import convention (same as the sibling helpers): the
    // CALL_IMPORT operand indexes THIS chunk's own import table, which
    // the runtime resolves by (module, name).
    let resolve_idx = c.add_import("ecma:promise", "resolve");
    let reject_idx = c.add_import("ecma:promise", "reject");
    let is_done_idx = c.add_import("ecma:value", "isGeneratorDone");

    let try_patch = crate::primitives::errors::emit_try_start(&mut c, 0);

    crate::primitives::globals::emit_read(&mut c, "__js_this", 0);
    c.emit_struct_field_op(Op::STRUCT_GET, 0, returned_key, 0); // null if never stamped
    ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_if(0);
    // `.return()` already closed the generator: {undefined, true}
    // without resuming (resume on a returned cont would run the body).
    core_wasm::undefined(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, value_local, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op_u16(Op::LOCAL_SET, done_local, 0);
    c.emit_else(0);

    // next() vs next(v): REF_IS_NULL is true for both Null and the
    // Undefined the VM pads missing args with — and §27.6.1.2 treats
    // next() and next(undefined) identically.
    c.emit_op_u16(Op::LOCAL_GET, v_local, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_if(0);
    // next() — drive one step; emit_next pushes [value, has_more].
    crate::primitives::globals::emit_read(&mut c, "__js_this", 0);
    emit_next(&mut c, 0);
    ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    ops::emit_dyn_not_into(imports, &mut c, 0); // i32 done = !has_more
    c.emit_op_u16(Op::LOCAL_SET, done_local, 0);
    c.emit_op_u16(Op::LOCAL_SET, value_local, 0);
    c.emit_else(0);
    // next(v) — RESUME with v; the suspended `yield` evaluates to v.
    crate::primitives::globals::emit_read(&mut c, "__js_this", 0);
    c.emit_op_u16(Op::LOCAL_GET, v_local, 0);
    emit_resume(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, value_local, 0);
    crate::primitives::globals::emit_read(&mut c, "__js_this", 0);
    c.emit_call(is_done_idx, 1, 0);
    ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, done_local, 0);
    c.emit_end(0);

    // Stamp started — `.throw()` on an unstarted generator keys on it.
    crate::primitives::globals::emit_read(&mut c, "__js_this", 0);
    core_wasm::bool_const(&mut c, 0, true);
    c.emit_struct_field_op(Op::STRUCT_SET, 0, started_key, 0);

    // §27.6.1.2: Await(value). Async bodies hand back promises (the
    // completion value arrives promise-wrapped); deliver the settled
    // value. Non-promises pass through; a rejection throws into the
    // catch below and rejects the result promise.

    // §27.6.1.2: Await(value). Async bodies hand back promises (the
    // completion value arrives promise-wrapped); deliver the settled
    // value. Non-promises pass through; a rejection throws into the
    // catch below and rejects the result promise.
    c.emit_op_u16(Op::LOCAL_GET, value_local, 0);
    crate::primitives::functions::emit_await(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, value_local, 0);

    // §27.5.3.5: completion with no explicit return value leaves null
    // on the stack — the spec `value` is undefined.
    c.emit_op_u16(Op::LOCAL_GET, done_local, 0);
    ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_if(0);
    c.emit_op_u16(Op::LOCAL_GET, value_local, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_if(0);
    core_wasm::undefined(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, value_local, 0);
    c.emit_end(0);
    c.emit_end(0);

    c.emit_end(0); // end returned-short-circuit if/else
    crate::primitives::errors::emit_try_end(&mut c, 0);

    c.emit_struct_new(0, 0, 0);
    c.emit_dup(0);
    c.emit_op_u16(Op::LOCAL_GET, value_local, 0);
    c.emit_struct_field_op(Op::STRUCT_SET, 0, value_key, 0);
    c.emit_dup(0);
    c.emit_op_u16(Op::LOCAL_GET, done_local, 0);
    // §27.6.1.2: `done` is a Boolean, not the raw i32 flag.
    ops::emit_i32_to_bool(&mut c, 0);
    c.emit_struct_field_op(Op::STRUCT_SET, 0, done_key, 0);
    c.emit_call(resolve_idx, 1, 0);
    c.emit_op(Op::RETURN, 0);

    crate::primitives::errors::patch_catch(&mut c, try_patch);
    c.emit_op_u16(Op::LOCAL_SET, err_local, 0);
    c.emit_op_u16(Op::LOCAL_GET, err_local, 0);
    c.emit_call(reject_idx, 1, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── Buffered iterator façade over a generator ───────────────────────────
//
// A language whose iterator surface is STATEFUL — `$gen->current()` re-read
// any number of times without advancing, `valid()` answering before the first
// `next()` — cannot sit directly on the ECMA generator protocol, where the
// only way to observe a value is to consume it. The gap is bridged by caching
// the last `IteratorResult` on the generator object itself, so the stateful
// methods read the cache and only `next`/`send`/`throw` actually advance.
//
// Gated by `profile.buffered_iterator_methods`. This lives beside the
// generator primitives it drives (`emit_next`, `emit_resume`,
// `emit_resume_throw`) rather than in a language-named module — the caching is
// a property of the ITERATOR PROTOCOL mismatch, not of any one language.
#[derive(Clone, Copy)]
enum BufferedGeneratorStepMode {
    Valid,
    Current,
    Value,
}

impl Compiler {
    fn emit_buffered_generator_set_bool_property(&mut self, obj_slot: u16, key: u16, value: bool) {
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_const(Value::Bool(value));
        self.emit_struct_field_op(Op::STRUCT_SET, 0, key);
    }

    fn emit_buffered_generator_mark_started(&mut self, obj_slot: u16, started_key: u16) {
        self.emit_buffered_generator_set_bool_property(obj_slot, started_key, true);
    }

    fn emit_buffered_generator_store_yielded_state(
        &mut self,
        obj_slot: u16,
        value_slot: u16,
        done_key: u16,
        current_key: u16,
    ) {
        self.emit_buffered_generator_set_bool_property(obj_slot, done_key, false);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_generator_yield_value(value_slot);
        self.emit_struct_field_op(Op::STRUCT_SET, 0, current_key);
    }

    fn emit_buffered_generator_store_completed_state(
        &mut self,
        obj_slot: u16,
        value_slot: u16,
        done_key: u16,
        current_key: u16,
        return_key: u16,
    ) {
        self.emit_buffered_generator_set_bool_property(obj_slot, done_key, true);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_struct_field_op(Op::STRUCT_SET, 0, return_key);
        self.emit_buffered_generator_set_bool_property(obj_slot, current_key, false);
    }

    fn emit_buffered_generator_set_step_result(
        &mut self,
        value_slot: u16,
        result_slot: u16,
        mode: BufferedGeneratorStepMode,
        yielded: bool,
    ) {
        match (mode, yielded) {
            (BufferedGeneratorStepMode::Valid, true) => self.emit_const(Value::Bool(true)),
            (BufferedGeneratorStepMode::Valid, false) => self.emit_const(Value::Bool(false)),
            (BufferedGeneratorStepMode::Current, false) => self.emit_const(Value::Bool(false)),
            (BufferedGeneratorStepMode::Current, true)
            | (BufferedGeneratorStepMode::Value, true) => {
                self.emit_generator_yield_value(value_slot);
            }
            (BufferedGeneratorStepMode::Value, false) => {
                self.emit_u16(Op::LOCAL_GET, value_slot);
            }
        }
        self.emit_u16(Op::LOCAL_SET, result_slot);
    }

    fn emit_buffered_generator_apply_next_result(
        &mut self,
        obj_slot: u16,
        value_slot: u16,
        has_more_slot: u16,
        result_slot: u16,
        mode: BufferedGeneratorStepMode,
        done_key: u16,
        current_key: u16,
        return_key: u16,
    ) {
        self.emit_u16(Op::LOCAL_GET, has_more_slot);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        };
        let line = self.line;
        self.chunk().emit_if(line);

        self.emit_buffered_generator_store_yielded_state(
            obj_slot,
            value_slot,
            done_key,
            current_key,
        );
        self.emit_buffered_generator_set_step_result(value_slot, result_slot, mode, true);

        self.chunk().emit_else(line);

        self.emit_buffered_generator_store_completed_state(
            obj_slot,
            value_slot,
            done_key,
            current_key,
            return_key,
        );
        self.emit_buffered_generator_set_step_result(value_slot, result_slot, mode, false);

        self.chunk().emit_end(line);
    }

    fn emit_buffered_generator_apply_resume_result(
        &mut self,
        obj_slot: u16,
        value_slot: u16,
        result_slot: u16,
        done_key: u16,
        current_key: u16,
        return_key: u16,
    ) {
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        let is_done_idx = self.import("ecma:value", "isGeneratorDone");
        self.emit_host_call(is_done_idx, 1);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        };
        let line = self.line;
        self.chunk().emit_if(line);

        self.emit_buffered_generator_store_completed_state(
            obj_slot,
            value_slot,
            done_key,
            current_key,
            return_key,
        );
        self.emit_buffered_generator_set_step_result(
            value_slot,
            result_slot,
            BufferedGeneratorStepMode::Value,
            false,
        );

        self.chunk().emit_else(line);

        self.emit_buffered_generator_store_yielded_state(
            obj_slot,
            value_slot,
            done_key,
            current_key,
        );
        self.emit_buffered_generator_set_step_result(
            value_slot,
            result_slot,
            BufferedGeneratorStepMode::Value,
            true,
        );

        self.chunk().emit_end(line);
    }

    fn emit_buffered_generator_start_with_next(
        &mut self,
        obj_slot: u16,
        result_slot: u16,
        mode: BufferedGeneratorStepMode,
        started_key: u16,
        done_key: u16,
        current_key: u16,
        return_key: u16,
    ) {
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        let line = self.line;
        crate::primitives::generators::emit_next(self.chunk(), line);
        let has_more_slot = self.define_local("__php_gen_has_more");
        self.emit_u16(Op::LOCAL_SET, has_more_slot);
        let value_slot = self.define_local("__php_gen_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);

        self.emit_buffered_generator_mark_started(obj_slot, started_key);
        self.emit_buffered_generator_apply_next_result(
            obj_slot,
            value_slot,
            has_more_slot,
            result_slot,
            mode,
            done_key,
            current_key,
            return_key,
        );
    }

    pub(crate) fn maybe_define_buffered_generator_key_index_slot(
        &mut self,
        key: Option<&str>,
    ) -> Option<u16> {
        if self.profile.buffered_iterator_methods && key.is_some() {
            let slot = self.define_local("__php_gen_loop_index");
            self.emit_const(Value::F64(0.0));
            self.emit_u16(Op::LOCAL_SET, slot);
            Some(slot)
        } else {
            None
        }
    }

    pub(crate) fn emit_buffered_generator_foreach_state(
        &mut self,
        cont_slot: u16,
        has_more_slot: u16,
        value_slot: u16,
    ) {
        let started_key = self.str_const("__php_gen_started");
        let current_key = self.str_const("__php_gen_current");
        let done_key = self.str_const("__php_gen_done");
        let return_key = self.str_const("__php_gen_return");

        self.emit_u16(Op::LOCAL_GET, cont_slot);
        self.emit_const(Value::Bool(true));
        self.emit_struct_field_op(Op::STRUCT_SET, 0, started_key);


        self.emit_u16(Op::LOCAL_GET, has_more_slot);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        };
        let line = self.line;
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, cont_slot);
        self.emit_const(Value::Bool(false));
        self.emit_struct_field_op(Op::STRUCT_SET, 0, done_key);
        self.emit_u16(Op::LOCAL_GET, cont_slot);
        self.emit_generator_yield_value(value_slot);
        self.emit_struct_field_op(Op::STRUCT_SET, 0, current_key);
        self.chunk().emit_else(line);

        self.emit_u16(Op::LOCAL_GET, cont_slot);
        self.emit_const(Value::Bool(true));
        self.emit_struct_field_op(Op::STRUCT_SET, 0, done_key);
        self.emit_u16(Op::LOCAL_GET, cont_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_struct_field_op(Op::STRUCT_SET, 0, return_key);
        self.emit_u16(Op::LOCAL_GET, cont_slot);
        self.emit_const(Value::Bool(false));
        self.emit_struct_field_op(Op::STRUCT_SET, 0, current_key);
        self.chunk().emit_br(2, line);

        self.chunk().emit_end(line);
    }

    pub(crate) fn emit_buffered_generator_key_binding(
        &mut self,
        key_slot: u16,
        value_slot: u16,
        key_index_slot: Option<u16>,
    ) {
        self.emit_generator_yield_key_or_fallback(value_slot, key_index_slot);
        self.emit_u16(Op::LOCAL_SET, key_slot);
    }

    pub(crate) fn emit_buffered_generator_value_binding(&mut self, var_slot: u16, value_slot: u16) {
        self.emit_generator_yield_value(value_slot);
        self.emit_u16(Op::LOCAL_SET, var_slot);
    }

    pub(crate) fn emit_buffered_generator_method_dispatch(
        &mut self,
        obj_tmp: u16,
        field_name: &str,
        arg_exprs: &[&Expression],
    ) -> Result<Option<usize>, String> {
        let is_buffered_generator_method = (field_name == "current" && arg_exprs.is_empty())
            || (field_name == "send" && arg_exprs.len() == 1)
            || (field_name == "next" && arg_exprs.is_empty())
            || (field_name == "throw" && arg_exprs.len() == 1)
            || (field_name == "valid" && arg_exprs.is_empty())
            || (field_name == "getReturn" && arg_exprs.is_empty())
            || (field_name == "rewind" && arg_exprs.is_empty());

        if !is_buffered_generator_method {
            return Ok(None);
        }

        let started_key = self.str_const("__php_gen_started");
        let current_key = self.str_const("__php_gen_current");
        let done_key = self.str_const("__php_gen_done");
        let return_key = self.str_const("__php_gen_return");
        let result_slot = self.define_local("__php_gen_method_result");

        self.emit_u16(Op::LOCAL_GET, obj_tmp);
        let is_gen_idx = self.import("ecma:value", "isGenerator");
        self.emit_host_call(is_gen_idx, 1);
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        match field_name {
            "getReturn" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_struct_field_op(Op::STRUCT_GET, 0, return_key);
                self.emit_u16(Op::LOCAL_SET, result_slot);
            }
            "valid" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_struct_field_op(Op::STRUCT_GET, 0, started_key);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                let line = self.line;
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_struct_field_op(Op::STRUCT_GET, 0, done_key);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_not(self.chunk(), line);
                };
                self.emit_u16(Op::LOCAL_SET, result_slot);

                self.chunk().emit_else(line);
                self.emit_buffered_generator_start_with_next(
                    obj_tmp,
                    result_slot,
                    BufferedGeneratorStepMode::Valid,
                    started_key,
                    done_key,
                    current_key,
                    return_key,
                );
                self.chunk().emit_end(line);
            }
            "current" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_struct_field_op(Op::STRUCT_GET, 0, started_key);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                let line = self.line;
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_struct_field_op(Op::STRUCT_GET, 0, done_key);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.chunk().emit_if(line);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::LOCAL_SET, result_slot);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_struct_field_op(Op::STRUCT_GET, 0, current_key);
                self.emit_u16(Op::LOCAL_SET, result_slot);

                self.chunk().emit_end(line);
                self.chunk().emit_else(line);
                self.emit_buffered_generator_start_with_next(
                    obj_tmp,
                    result_slot,
                    BufferedGeneratorStepMode::Current,
                    started_key,
                    done_key,
                    current_key,
                    return_key,
                );
                self.chunk().emit_end(line);
            }
            "send" | "next" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_struct_field_op(Op::STRUCT_GET, 0, started_key);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                let line = self.line;
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_struct_field_op(Op::STRUCT_GET, 0, done_key);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.chunk().emit_if(line);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::LOCAL_SET, result_slot);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                if field_name == "send" {
                    self.compile_expr(&arg_exprs[0])?;
                } else {
                    self.emit_null();
                }
                let line = self.line;
                crate::primitives::generators::emit_resume(self.chunk(), line);
                let value_slot = self.define_local("__php_gen_resume_value");
                self.emit_u16(Op::LOCAL_SET, value_slot);
                self.emit_buffered_generator_apply_resume_result(
                    obj_tmp,
                    value_slot,
                    result_slot,
                    done_key,
                    current_key,
                    return_key,
                );
                self.chunk().emit_end(line);

                self.chunk().emit_else(line);
                self.emit_buffered_generator_start_with_next(
                    obj_tmp,
                    result_slot,
                    BufferedGeneratorStepMode::Value,
                    started_key,
                    done_key,
                    current_key,
                    return_key,
                );
                self.chunk().emit_end(line);
                // Mark as moved for rewind() check
                let moved_key = self.str_const("__php_gen_moved");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(true));
                self.emit_struct_field_op(Op::STRUCT_SET, 0, moved_key);
            }
            "throw" => {
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_struct_field_op(Op::STRUCT_GET, 0, started_key);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                let line = self.line;
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_struct_field_op(Op::STRUCT_GET, 0, done_key);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.chunk().emit_if(line);
                self.emit_const(Value::Bool(false));
                self.emit_u16(Op::LOCAL_SET, result_slot);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.compile_expr(&arg_exprs[0])?;
                let line = self.line;
                crate::primitives::generators::emit_resume_throw(self.chunk(), line);
                let value_slot = self.define_local("__php_gen_throw_value");
                self.emit_u16(Op::LOCAL_SET, value_slot);
                self.emit_buffered_generator_apply_resume_result(
                    obj_tmp,
                    value_slot,
                    result_slot,
                    done_key,
                    current_key,
                    return_key,
                );
                self.chunk().emit_end(line);

                self.chunk().emit_else(line);
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                let line = self.line;
                crate::primitives::generators::emit_next(self.chunk(), line);
                let has_more_slot = self.define_local("__php_gen_throw_has_more");
                self.emit_u16(Op::LOCAL_SET, has_more_slot);
                let start_value_slot = self.define_local("__php_gen_throw_start_value");
                self.emit_u16(Op::LOCAL_SET, start_value_slot);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_const(Value::Bool(true));
                self.emit_struct_field_op(Op::STRUCT_SET, 0, started_key);


                self.emit_u16(Op::LOCAL_GET, has_more_slot);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                self.chunk().emit_if(line);

                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.compile_expr(&arg_exprs[0])?;
                let line = self.line;
                crate::primitives::generators::emit_resume_throw(self.chunk(), line);
                let start_resume_slot = self.define_local("__php_gen_throw_resume_value");
                self.emit_u16(Op::LOCAL_SET, start_resume_slot);
                self.emit_buffered_generator_apply_resume_result(
                    obj_tmp,
                    start_resume_slot,
                    result_slot,
                    done_key,
                    current_key,
                    return_key,
                );

                self.chunk().emit_else(line);
                self.emit_buffered_generator_store_completed_state(
                    obj_tmp,
                    start_value_slot,
                    done_key,
                    current_key,
                    return_key,
                );
                self.emit_buffered_generator_set_step_result(
                    start_value_slot,
                    result_slot,
                    BufferedGeneratorStepMode::Value,
                    false,
                );

                self.chunk().emit_end(line);
                self.chunk().emit_end(line);
            }
            "rewind" => {
                // PHP Generator::rewind() throws if the generator has
                // been advanced past the initial yield (via next/send).
                let moved_key = self.str_const("__php_gen_moved");
                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                self.emit_struct_field_op(Op::STRUCT_GET, 0, moved_key);
                {
                    let line = self.line;
                    crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
                };
                let line = self.line;
                self.chunk().emit_if(line);
                self.emit_const(Value::String(Arc::from(
                    "Cannot rewind a generator that was already run",
                )));
                crate::primitives::errors::emit_throw(&mut self.chunks[self.current], line);
                self.chunk().emit_end(line);
                self.emit_null();
                self.emit_u16(Op::LOCAL_SET, result_slot);
            }
            _ => unreachable!(),
        }

        let line = self.line;
        self.chunk().emit_else(line);
        Ok(Some(result_slot as usize))
    }

    pub(crate) fn emit_buffered_generator_close_ident_if_needed(&mut self, name: &str) {
        if !self.profile.buffered_iterator_methods {
            return;
        }

        self.emit_var_get(name);
        let gen_slot = self.define_local("__buffered_generator_overwrite");
        self.emit_u16(Op::LOCAL_SET, gen_slot);

        self.emit_u16(Op::LOCAL_GET, gen_slot);
        let is_generator = self.import("ecma:value", "isGenerator");
        self.emit_host_call(is_generator, 1);
        let line = self.line;
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, gen_slot);
        let started_key = self.str_const("__php_gen_started");
        self.emit_struct_field_op(Op::STRUCT_GET, 0, started_key);
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, gen_slot);
        let done_key = self.str_const("__php_gen_done");
        self.emit_struct_field_op(Op::STRUCT_GET, 0, done_key);
        crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        self.emit(Op::I32_EQZ);
        self.chunk().emit_if(line);

        self.emit_u16(Op::LOCAL_GET, gen_slot);
        self.emit_null();
        self.emit_generator_control_packet_from_stack("return");
        let line = self.line;
        crate::primitives::generators::emit_resume(self.chunk(), line);
        self.emit(Op::DROP);

        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
        self.chunk().emit_end(line);
    }
    pub(crate) fn finish_buffered_generator_method_dispatch(&mut self, result_slot: usize) {
        let line = self.line;
        self.emit_u16(Op::LOCAL_SET, result_slot as u16);
        self.chunk().emit_end(line);
        self.emit_u16(Op::LOCAL_GET, result_slot as u16);
    }
}
