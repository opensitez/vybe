//! Shared Base64 codec core.
//!
//! Language and platform adapters own their public surface shapes
//! (`java.util.Base64`, `.NET Convert`, PHP functions, Dart codecs, ...). The
//! actual binary-string Base64 transform is one common primitive over the ECMA
//! host imports, so adapters can share the same core instead of each spelling
//! reaching directly for `btoa`/`atob`.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use crate::primitives::instructions::host;
use crate::primitives::{collections, loops, strings};

/// Stack: `[binary_string] -> [base64_string]`.
pub fn emit_encode_binary_string(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:string", "btoa", 1, line);
}

/// Stack: `[base64_string] -> [binary_string]`.
pub fn emit_decode_binary_string(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:string", "atob", 1, line);
}

/// Stack: `[byte_array] -> [binary_string]`.
pub fn emit_byte_array_to_binary_string(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_byte_array_slot_to_binary_string(chunks, current, None, None, None, line);
}

/// Locals: `bytes_slot[, start_slot, end_slot] -> [binary_string]`.
pub fn emit_byte_array_slot_to_binary_string(
    chunks: &mut [Chunk],
    current: usize,
    bytes_slot: Option<u16>,
    start_slot: Option<u16>,
    end_slot: Option<u16>,
    line: u32,
) {
    let owned_bytes_slot;
    let bytes_slot = match bytes_slot {
        Some(slot) => slot,
        None => {
            owned_bytes_slot = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, owned_bytes_slot, line);
            owned_bytes_slot
        }
    };
    if !crate::primitives::polyfills::is_compiling_runtime_helper() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
        if start_slot.is_none() && end_slot.is_none() {
            crate::primitives::bundle::emit_stacked_helper_call(
                &mut chunks[current],
                "__vybe_bytes_to_binary_string_all",
                1,
                line,
            );
            return;
        }
        match start_slot {
            Some(slot) => chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line),
            None => chunks[current].emit_i32_const(0, line),
        }
        match end_slot {
            Some(slot) => chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line),
            None => {
                chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
                collections::emit_len(chunks, current, line);
            }
        }
        crate::primitives::bundle::emit_stacked_helper_call(
            &mut chunks[current],
            "__vybe_bytes_to_binary_string",
            3,
            line,
        );
        return;
    }
    let start = chunks[current].alloc_scratch(1);
    let end = chunks[current].alloc_scratch(1);
    match start_slot {
        Some(slot) => chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line),
        None => chunks[current].emit_i32_const(0, line),
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
    match end_slot {
        Some(slot) => chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line),
        None => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
            collections::emit_len(chunks, current, line);
        }
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, end, line);
    // Keep the short-input loop free of fragment bookkeeping. Both paths
    // retain the original signed-i32 indexing/conversion contract.
    chunks[current].emit_op_u16(Op::LOCAL_GET, end, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_i32_const(8192, line);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_if_value(line);
    emit_small_byte_array_string(
        chunks,
        current,
        Some(bytes_slot),
        Some(start),
        Some(end),
        line,
    );
    chunks[current].emit_else(line);
    emit_fragmented_byte_array_string(chunks, current, bytes_slot, Some(start), Some(end), line);
    chunks[current].emit_end(line);
}

fn emit_fragmented_byte_array_string(
    chunks: &mut [Chunk],
    current: usize,
    bytes_slot: u16,
    start_slot: Option<u16>,
    end_slot: Option<u16>,
    line: u32,
) {
    let out_slot = chunks[current].alloc_scratch(1);
    let i_slot = chunks[current].alloc_scratch(1);
    let end_local = chunks[current].alloc_scratch(1);
    let fragment = chunks[current].alloc_scratch(1);
    let fragment_len = chunks[current].alloc_scratch(1);

    // Every fragment is fromCharCode's string result. Retain input reads and
    // conversion order, but copy the assembled string only once at the join.
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fragment, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fragment_len, line);
    match start_slot {
        Some(slot) => chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line),
        None => chunks[current].emit_i32_const(0, line),
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    match end_slot {
        Some(slot) => chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line),
        None => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
            collections::emit_len(chunks, current, line);
        }
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_local, line);

    let state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_local, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, fragment, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    collections::emit_get(chunks, current, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "fromCharCode",
        1,
        line,
    );
    strings::emit_concat(&mut chunks[current], 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fragment, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fragment_len, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_TEE, fragment_len, line);
    chunks[current].emit_i32_const(256, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fragment, line);
    host::emit(&mut chunks[current], "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fragment, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fragment_len, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fragment, line);
    host::emit(&mut chunks[current], "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_string_const("", line);
    host::emit(&mut chunks[current], "ecma:array", "join", 2, line);
}

fn emit_small_byte_array_string(
    chunks: &mut [Chunk],
    current: usize,
    bytes_slot: Option<u16>,
    start_slot: Option<u16>,
    end_slot: Option<u16>,
    line: u32,
) {
    let owned_bytes_slot;
    let bytes_slot = match bytes_slot {
        Some(slot) => slot,
        None => {
            owned_bytes_slot = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, owned_bytes_slot, line);
            owned_bytes_slot
        }
    };
    let out_slot = chunks[current].alloc_scratch(1);
    let i_slot = chunks[current].alloc_scratch(1);
    let end_local = chunks[current].alloc_scratch(1);

    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    match start_slot {
        Some(slot) => chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line),
        None => chunks[current].emit_i32_const(0, line),
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    match end_slot {
        Some(slot) => chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line),
        None => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
            collections::emit_len(chunks, current, line);
        }
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_local, line);

    let state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_local, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    collections::emit_get(chunks, current, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "fromCharCode",
        1,
        line,
    );
    strings::emit_concat(&mut chunks[current], 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

/// Stack: `[binary_string] -> [byte_array]`.
pub fn emit_binary_string_to_byte_array(chunks: &mut [Chunk], current: usize, line: u32) {
    if !crate::primitives::polyfills::is_compiling_runtime_helper() {
        crate::primitives::bundle::emit_stacked_helper_call(
            &mut chunks[current],
            "__vybe_binary_string_to_bytes",
            1,
            line,
        );
        return;
    }
    let text_slot = chunks[current].alloc_scratch(1);
    let out_slot = chunks[current].alloc_scratch(1);
    let i_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    strings::emit_length(&mut chunks[current], line);
    host::emit(&mut chunks[current], "wasm:js-number", "toI32", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    let state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "charCodeAt",
        2,
        line,
    );
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

pub(crate) fn build_shared_helper(name: &str) -> Chunk {
    let argc = if name == "__stdlib_bytes_to_binary_string" {
        3
    } else {
        1
    };
    let mut chunk = Chunk::new(name);
    chunk.arity = argc;
    chunk.local_count = argc as u16;
    let mut chunks = vec![chunk];
    match name {
        "__stdlib_bytes_to_binary_string_all" => {
            emit_byte_array_slot_to_binary_string(&mut chunks, 0, Some(0), None, None, 0)
        }
        "__stdlib_bytes_to_binary_string" => {
            emit_byte_array_slot_to_binary_string(&mut chunks, 0, Some(0), Some(1), Some(2), 0)
        }
        "__stdlib_binary_string_to_bytes" => {
            chunks[0].emit_op_u16(Op::LOCAL_GET, 0, 0);
            emit_binary_string_to_byte_array(&mut chunks, 0, 0);
        }
        _ => unreachable!("unknown binary string helper"),
    }
    chunks[0].emit_op(Op::RETURN, 0);
    chunks.pop().unwrap()
}
