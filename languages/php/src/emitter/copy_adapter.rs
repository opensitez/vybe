//! PHP copy/identity adapters.
//!
//! PHP arrays are value-copied on assignment, while objects and callables stay
//! reference-like. The syntax-level walker still marks only the RHS sites that
//! can alias; this adapter provides the PHP leaf behind that marker without a
//! PHP copy-on-assign and strict-equality adapter.

use vybe_compiler::primitives::functions::create_function_chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn call_import(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, name);
    chunk.emit_call(idx, argc, line);
}

fn call_ref(chunk: &mut Chunk, argc: u8, line: u32) {
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, argc, line);
}

fn ref_func(chunk: &mut Chunk, func_idx: usize, line: u32) {
    chunk.emit_op_u16(Op::REF_FUNC, func_idx as u16, line);
    chunk.emit(0, line);
}

fn bump_i32(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    lset(chunk, slot, line);
}

fn helper_loop_start(chunk: &mut Chunk, line: u32) -> vybe_compiler::primitives::loops::LoopState {
    let block_patch = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    vybe_compiler::primitives::loops::LoopState {
        block_patch,
        loop_patch,
        body_block_patch: None,
    }
}

fn helper_loop_end(
    chunk: &mut Chunk,
    state: vybe_compiler::primitives::loops::LoopState,
    line: u32,
) {
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(state.loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(state.block_patch);
}

/// Stack: [] -> [i32 bool]. PHP arrays exclude objects such as ArrayObject,
/// even when their implementation uses map storage.
pub(crate) fn emit_php_arrayish_slot(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    call_import(chunk, "php:array", "isArray", 1, line);
    // The existing PHP array predicate returns I32. Its generic native host
    // boundary boxes the value; recover that exact checked I32 for consumers.
    call_import(chunk, "wasm:js-number", "toI32", 1, line);
}

fn build_strict_eq_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let helper_idx = chunks.len();
    let mut c = create_function_chunk("__php_strict_eq_impl", 2);
    c.local_count = c.local_count.max(2);
    c.param_count = 2;
    c.takes_receiver = false;
    c.direct_entry_has_captures = Some(false);
    c.module_receiver_abi =
        vybe_compiler::primitives::class_context::module_receiver_abi(chunks);
    // Capture only the empty helper's declaration before its sizeable body is
    // emitted. Recursive comparisons have this exact capture-free signature.
    let recursive_signature = c.clone();
    let a_slot = 0;
    let b_slot = 1;
    let a_is_slot = alloc_local(&mut c);
    let b_is_slot = alloc_local(&mut c);
    let entries_a_slot = alloc_local(&mut c);
    let entries_b_slot = alloc_local(&mut c);
    let len_slot = alloc_local(&mut c);
    let i_slot = alloc_local(&mut c);
    let pair_a_slot = alloc_local(&mut c);
    let pair_b_slot = alloc_local(&mut c);
    let result_slot = alloc_local(&mut c);

    emit_php_arrayish_slot(&mut c, a_slot, line);
    lset(&mut c, a_is_slot, line);
    emit_php_arrayish_slot(&mut c, b_slot, line);
    lset(&mut c, b_is_slot, line);

    lget(&mut c, a_is_slot, line);
    lget(&mut c, b_is_slot, line);
    c.emit_op(Op::I32_OR, line);
    c.emit_if_value(line);

    lget(&mut c, a_is_slot, line);
    lget(&mut c, b_is_slot, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    c.emit_bool_const(false, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);

    lget(&mut c, a_slot, line);
    call_import(&mut c, "ecma:object", "entries", 1, line);
    lset(&mut c, entries_a_slot, line);
    lget(&mut c, b_slot, line);
    call_import(&mut c, "ecma:object", "entries", 1, line);
    lset(&mut c, entries_b_slot, line);

    lget(&mut c, entries_a_slot, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    lset(&mut c, len_slot, line);
    lget(&mut c, len_slot, line);
    lget(&mut c, entries_b_slot, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    c.emit_op(Op::I32_NE, line);
    c.emit_if(line);
    c.emit_bool_const(false, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);

    // Keep the loop's control flag in the same i32 representation as its
    // comparisons. Only the public/recursive return ABI is a boxed boolean.
    c.emit_i32_const(1, line);
    lset(&mut c, result_slot, line);
    c.emit_i32_const(0, line);
    lset(&mut c, i_slot, line);

    let loop_state = helper_loop_start(&mut c, line);
    lget(&mut c, i_slot, line);
    lget(&mut c, len_slot, line);
    c.emit_op(Op::I32_LT_S, line);
    lget(&mut c, result_slot, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_br_if(1, line);

    lget(&mut c, entries_a_slot, line);
    lget(&mut c, i_slot, line);
    c.emit_op(Op::ARRAY_GET, line);
    lset(&mut c, pair_a_slot, line);
    lget(&mut c, entries_b_slot, line);
    lget(&mut c, i_slot, line);
    c.emit_op(Op::ARRAY_GET, line);
    lset(&mut c, pair_b_slot, line);

    lget(&mut c, pair_a_slot, line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::ARRAY_GET, line);
    lget(&mut c, pair_b_slot, line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::ARRAY_GET, line);
    vybe_compiler::primitives::ops::emit_js_strict_eq(&mut c, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    c.emit_i32_const(0, line);
    lset(&mut c, result_slot, line);
    c.emit_end(line);

    lget(&mut c, result_slot, line);
    c.emit_if(line);
    lget(&mut c, pair_a_slot, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::ARRAY_GET, line);
    lget(&mut c, pair_b_slot, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::ARRAY_GET, line);
    c.emit_direct_defined_call(
        u32::try_from(helper_idx).expect("strict equality helper index exceeds u32"),
        &recursive_signature, false, line,
    ).expect("strict equality helper requires a fixed capture-free signature");
    // This recursive helper always returns a boolean, not a PHP truthy value.
    call_import(&mut c, "wasm:js-boolean", "cast", 1, line);
    lset(&mut c, result_slot, line);
    c.emit_end(line);

    bump_i32(&mut c, i_slot, line);
    helper_loop_end(&mut c, loop_state, line);

    lget(&mut c, result_slot, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut c, line);
    c.emit_op(Op::RETURN, line);

    c.emit_else(line);

    lget(&mut c, a_slot, line);
    lget(&mut c, b_slot, line);
    vybe_compiler::primitives::ops::emit_js_strict_eq(&mut c, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut c, line);
    c.emit_op(Op::RETURN, line);

    c.emit_end(line);

    chunks.push(c);
    helper_idx
}

pub fn emit_php_copy_on_assign(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    call_import(&mut chunks[current], "php:array", "copy", 1, line);
}

pub fn emit_php_strict_eq(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let helper_idx = build_strict_eq_helper(chunks, line);
    if vybe_compiler::primitives::functions::try_emit_direct_helper_call(
        chunks, current, helper_idx, 2, line,
    ) {
        return;
    }
    let b_slot = alloc_local(&mut chunks[current]);
    let a_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], b_slot, line);
    lset(&mut chunks[current], a_slot, line);
    ref_func(&mut chunks[current], helper_idx, line);
    lget(&mut chunks[current], a_slot, line);
    lget(&mut chunks[current], b_slot, line);
    call_ref(&mut chunks[current], 2, line);
}

#[cfg(test)]
mod compact_strict_equality_tests {
    use super::*;

    #[test]
    fn strict_equality_outer_and_recursive_edges_are_direct() {
        let mut chunks = vec![Chunk::new("strict-equality-caller")];
        emit_php_strict_eq(&mut chunks, 0, 2, 1);
        assert_eq!(chunks[0].code.len(), 7);
        assert_eq!(chunks[0].local_count, 0);
        assert_eq!(chunks[0].direct_defined_calls.len(), 1);
        assert_eq!(chunks[1].direct_defined_calls.len(), 1);
        assert_eq!(chunks[1].direct_defined_calls.values().next().unwrap().target, 1);
    }

}

fn strict_eq_context_key(chunks: &[Chunk], _: u32) -> Vec<u8> {
    vec![match vybe_compiler::primitives::class_context::module_receiver_abi(chunks) {
        vybe_runtime::chunk::ReceiverAbi::Ambient => 0,
        vybe_runtime::chunk::ReceiverAbi::Parameter => 1,
    }]
}
fn unused_strict_eq_factory() -> Chunk { unreachable!("strict equality uses its unit-owned contextual builder") }
fn unused_strict_eq_caller(_: &mut Chunk, _: usize, _: u32) { unreachable!("strict equality uses its checked direct invocation") }
pub(crate) fn strict_eq_definition() -> vybe_compiler::primitives::helper_templates::Definition {
    use vybe_compiler::primitives::helper_templates::{Contract, Definition};
    Definition {
        contract: Contract::ContextualTypedDirectCallAppendOnlyUnit {
            key: strict_eq_context_key, build: build_strict_eq_helper,
            result: vybe_runtime::registry::emitter_results::ResultRepresentation::BoxedBoolean,
        },
        operation: "php.strict_eq", namespace: "php", version: 1,
        parser: crate::parse, producer: super::dispatch::dispatch, arity: 2,
        helper_name: "__php_strict_eq_impl",
        factory: unused_strict_eq_factory, caller: unused_strict_eq_caller,
    }
}
