use vybe_compiler::primitives::{
    class_slots,
    instructions::core_wasm,
    tuples::{emit_is_tuple, emit_tag},
};
use vybe_runtime::{Chunk, opcode::Op};
pub fn emit_propagate_tag(chunks: &mut [Chunk], current: usize, line: u32) {
    let src = chunks[current].alloc_scratch(1);
    set(chunks, current, src, line); // [result]; source stashed

    // if isArray(source) && isTuple(source) && isArray(result) { tag(result) }
    get(chunks, current, src, line);
    call_is_array(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(chunks, current, src, line);
    emit_is_tuple(chunks, current, line);
    chunks[current].emit_if(line);
    core_wasm::dup(&mut chunks[current], line); // [result, result]
    call_is_array(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_tag(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn get(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn call_is_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(idx, 1, line);
}

fn emit_positional_item_fields(chunks: &mut [Chunk], current: usize, n: usize, line: u32) {
    for i in 0..n {
        emit_positional_item_field(chunks, current, i, &format!("Item{}", i + 1), line);
        emit_positional_item_field(chunks, current, i, &format!("item{}", i + 1), line);
    }
}

fn emit_positional_item_field(
    chunks: &mut [Chunk],
    current: usize,
    index: usize,
    name: &str,
    line: u32,
) {
    let c = &mut chunks[current];
    c.emit_dup(line);
    c.emit_dup(line);
    core_wasm::i32_const(c, line, index as i32);
    c.emit_op(Op::ARRAY_GET, line);
    let k = class_slots::resolve_interned(
        c,
        &class_slots::ClassSlot::internal(name),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_set(
        c,
        class_slots::ObjSource::Stack,
        &k,
        class_slots::ValueSource::Stack,
        line,
    );
}

/// Structural `repr` transform shared across languages: rewrite an
/// already-formatted list string `"[a, b]"` into the tuple form `"(a, b)"`,
/// adding the single-element trailing comma `"(x,)"`. Element formatting stays
/// the front-end's concern (it produced the list string); only the brackets and
/// the lone-element comma are universal. Stack: `[list_string] -> [tuple_string]`.
pub fn emit_list_string_to_tuple(chunk: &mut Chunk, line: u32) {
    let s = chunk.alloc_scratch(4);
    let len = s + 1;
    let inner = s + 2;
    let res = s + 3;
    chunk.emit_op_u16(Op::LOCAL_SET, s, line);

    // len = s.length
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    let strlen = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(strlen, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len, line);

    // inner = slice(s, 1, len - 1)  (ecma:array.slice dispatches string→substring)
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op_u16(Op::LOCAL_GET, len, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_SUB, line);
    let slice = chunk.add_import("ecma:array", "slice");
    chunk.emit_call(slice, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, inner, line);

    // res = "(" + inner
    core_wasm::string_const(chunk, line, "(");
    chunk.emit_op_u16(Op::LOCAL_GET, inner, line);
    vybe_compiler::primitives::strings::emit_str_concat(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, res, line);

    // single element? inner.length > 0 && inner.indexOf(", ") < 0
    chunk.emit_op_u16(Op::LOCAL_GET, inner, line);
    let ilen = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(ilen, 1, line);
    core_wasm::i32_const(chunk, line, 0);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, inner, line);
    core_wasm::string_const(chunk, line, ", ");
    let index_of = chunk.add_import("ecma:string", "indexOf");
    chunk.emit_call(index_of, 2, line);
    core_wasm::i32_const(chunk, line, 0);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, res, line);
    core_wasm::string_const(chunk, line, ",");
    vybe_compiler::primitives::strings::emit_str_concat(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, res, line);
    chunk.emit_end(line);

    // res + ")"
    chunk.emit_op_u16(Op::LOCAL_GET, res, line);
    core_wasm::string_const(chunk, line, ")");
    vybe_compiler::primitives::strings::emit_str_concat(chunk, line);
}
