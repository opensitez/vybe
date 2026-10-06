// Frozen compact hybrid experiment; generic production remains typed.
use vybe_compiler::primitives::strings::{
    emit_length as emit_str_length, emit_substring as emit_str_substring,
};
use vybe_runtime::{Chunk, Op};
fn emit_str_char_code_at(c: &mut Chunk, line: u32) {
    let f = c.add_import("wasm:js-string", "charCodeAt");
    c.emit_call(f, 2, line);
}
// Exact immediate encoding of the String/I32 constants used by this frozen
// builder, matching expressions::emit_const_index without exposing compiler API.
fn emit_const_index(chunk: &mut Chunk, index: u32, line: u32) {
    match chunk.constants[index as usize].clone() {
        vybe_runtime::Value::String(value) => chunk.emit_string_const(&value, line),
        vybe_runtime::Value::I32(value) => chunk.emit_i32_const(value, line),
        other => panic!("unexpected format_map fixture constant: {other:?}"),
    }
}
use vybe_compiler::primitives::strings::emit_str_concat;
pub fn build_format_map(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_format_map");
    c.arity = 2;
    c.local_count = 7; // s(0), d(1), out(2), i(3), len(4), end(5), key(6)
    let s = 0u16;
    let d = 1u16;
    let out = 2u16;
    let i = 3u16;
    let len = 4u16;
    let end = 5u16;
    let key = 6u16;
    let parts = c.alloc_scratch(3);
    let fragment = parts + 1;
    let mode = parts + 2; // 0: typed; 1: inspect first placeholder; 2: fragments
    let create_parts = c.add_import("vybe:js-array", "newWithLength");
    let append_part = c.add_import("ecma:array", "push");
    let join_parts = c.add_import("ecma:array", "join");

    let to_str = c.add_import("ecma:string", "String");

    // out = ""
    let empty = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("")));
    emit_const_index(&mut c, empty, 0);
    c.emit_op_u16(Op::LOCAL_SET, out, 0);

    // i = 0; len = wasm:js-string.length(s)
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    emit_str_length(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_i32_const(128, 0);
    c.emit_op(Op::I32_GE_S, 0);
    c.emit_op_u16(Op::LOCAL_SET, mode, 0);

    let open_brace = c.add_constant(vybe_runtime::Value::I32(b'{' as i32));
    let close_brace = c.add_constant(vybe_runtime::Value::I32(b'}' as i32));

    let outer_block = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    // Private cursors and js-string.length are i32 values.
    // if i >= len: break
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::I32_LT_S, 0);
    c.emit_op(vybe_runtime::opcode::Op::I32_EQZ, 0);
    c.emit_br_if(1, 0);

    // ch = wasm:js-string.charCodeAt(s, i)
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    emit_str_char_code_at(&mut c, 0);

    // Branch on '{' / '}' / literal
    let ch_slot = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, ch_slot, 0);

    // -- '{' branch --
    let open_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, ch_slot, 0);
    emit_const_index(&mut c, open_brace, 0);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);

    // Find closing '}': end = i+1; while end < len && s[end] != '}': end++
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, end, 0);

    let scan_block = c.emit_block(0);
    let (scan_loop, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::I32_LT_S, 0);
    c.emit_op(vybe_runtime::opcode::Op::I32_EQZ, 0);
    c.emit_br_if(1, 0);
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    emit_str_char_code_at(&mut c, 0);
    emit_const_index(&mut c, close_brace, 0);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_br_if(1, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, end, 0);
    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(scan_loop);
    c.emit_end(0);
    c.patch_block(scan_block);

    // key = s.substring(i+1, end); out += String(d[key]); i = end + 1
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);

    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_GET, d, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_call(to_str, 1, 0);
    // Both operands stay on the stack on the ordinary typed path.
    c.emit_op_u16(Op::LOCAL_GET, mode, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_if_params(0, 2, 1);
    emit_str_concat(&mut c, 0);
    c.emit_else(0);
    // Only the candidate/active path spills the two proven-string operands.
    c.emit_op_u16(Op::LOCAL_SET, fragment, 0);
    c.emit_op_u16(Op::LOCAL_SET, out, 0);
    c.emit_op_u16(Op::LOCAL_GET, mode, 0);
    c.emit_i32_const(1, 0);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_if(0);
    c.emit_i32_const(0, 0);
    c.emit_op_u16(Op::LOCAL_SET, mode, 0);
    c.emit_op_u16(Op::LOCAL_GET, fragment, 0);
    emit_str_length(&mut c, 0);
    c.emit_i32_const(65536, 0);
    c.emit_op(Op::I32_GE_S, 0);
    c.emit_if(0);
    c.emit_i32_const(0, 0);
    c.emit_call(create_parts, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, parts, 0);
    c.emit_i32_const(2, 0);
    c.emit_op_u16(Op::LOCAL_SET, mode, 0);
    c.emit_end(0);
    c.emit_end(0);

    c.emit_op_u16(Op::LOCAL_GET, mode, 0);
    c.emit_i32_const(2, 0);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_if_value(0);
    for value in [out, fragment] {
        c.emit_op_u16(Op::LOCAL_GET, parts, 0);
        c.emit_op_u16(Op::LOCAL_GET, value, 0);
        c.emit_call(append_part, 2, 0);
        c.emit_op(Op::DROP, 0);
    }
    c.emit_string_const("", 0);
    c.emit_else(0);
    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_GET, fragment, 0);
    emit_str_concat(&mut c, 0);
    c.emit_end(0);
    c.emit_end(0);
    c.emit_op_u16(Op::LOCAL_SET, out, 0);

    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_br(1, 0); // continue outer loop
    c.emit_end(0);
    c.patch_block(open_block);

    // -- literal char path: out += s.substring(i, i+1); i++
    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    emit_str_substring(&mut c, 0);
    emit_str_concat(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, out, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(outer_block);

    c.emit_op_u16(Op::LOCAL_GET, mode, 0);
    c.emit_i32_const(2, 0);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_if_value(0);
    c.emit_op_u16(Op::LOCAL_GET, parts, 0);
    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_call(append_part, 2, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, parts, 0);
    c.emit_string_const("", 0);
    c.emit_call(join_parts, 2, 0);
    c.emit_else(0);
    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_end(0);
    c.emit_op(Op::RETURN, 0);
    c
}
