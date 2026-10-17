use vybe_compiler::primitives::strings::{emit_str_char_code_at, emit_str_concat};
use vybe_runtime::{Chunk, Op};
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

    let to_str = c.add_import("ecma:string", "String");

    // out = ""
    let empty = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("")));
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, empty, 0);
    c.emit_op_u16(Op::LOCAL_SET, out, 0);

    // i = 0; len = wasm:js-string.length(s)
    vybe_compiler::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    emit_str_length(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    let open_brace = c.add_constant(vybe_runtime::Value::I32(b'{' as i32));

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
    vybe_compiler::primitives::expressions::emit_const_index(&mut c, open_brace, 0);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0);

    emit_format_map_closing_brace(&mut c, s, i, len, end);

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
    // Both the accumulated output and String(d[key]) are strings.
    emit_str_concat(&mut c, 0);
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

    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
fn emit_format_map_closing_brace(c: &mut Chunk, s: u16, i: u16, len: u16, end: u16) {
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_string_const("}", 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_i32_const(1, 0);
    c.emit_op(Op::I32_ADD, 0);
    let find = c.add_import("ecma:string", "indexOf");
    c.emit_call(find, 3, 0);
    c.emit_op_u16(Op::LOCAL_TEE, end, 0);
    c.emit_f64_const(0.0, 0);
    c.emit_op(Op::F64_LT, 0);
    c.emit_if(0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op_u16(Op::LOCAL_SET, end, 0);
    c.emit_end(0);
}
fn emit_str_length(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(idx, 1, line);
}
fn emit_str_substring(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "substring");
    chunk.emit_call(idx, 3, line);
}
