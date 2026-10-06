use std::sync::Arc;
use vybe_runtime::{Chunk, Op, Value};
fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}
fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
        Value::BigInt(v) => chunk.emit_i64_const(v.to_i64_wrapping(), line),
        Value::String(s) => chunk.emit_string_const(&s, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),

        _ => {
            unreachable!("push_const: unexpected value type");
        }
    }
}
fn push_str(chunk: &mut Chunk, v: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(v)), line);
}
fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}
fn coerce_to_str(chunk: &mut Chunk, line: u32) {
    vybe_language_php::emitter::string_adapter::emit_echo_stringify(
        std::slice::from_mut(chunk),
        0,
        1,
        line,
    );
}
pub fn emit_strip_tags(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // Scan the string: copy text through; for each `<...>` tag, extract its
    // name (letters/digits after `<` and an optional `/`) and keep the whole
    // tag only when `<name>` appears in the (lower-cased) allow-list. PHP
    // keeps tag *text content* — only the tags themselves are removed.
    let chunk = &mut chunks[current];
    let allowed_slot = alloc_local(chunk);
    let s_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let j_slot = alloc_local(chunk);
    let k_slot = alloc_local(chunk);
    let name_slot = alloc_local(chunk);
    let cc_slot = alloc_local(chunk);

    // Stack (bottom→top): s, [allowed]. Pop allowed (lower-cased) then s.
    if argc >= 2 {
        coerce_to_str(chunk, line);
        {
            let idx = chunk.add_import("ecma:string", "toLowerCase");
            chunk.emit_call(idx, 1, line);
        }
        lset(chunk, allowed_slot, line);
    } else {
        push_str(chunk, "", line);
        lset(chunk, allowed_slot, line);
    }
    coerce_to_str(chunk, line);
    lset(chunk, s_slot, line);
    push_str(chunk, "", line);
    lset(chunk, out_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, n_slot, line);

    let _ = chunk;
    let outer = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // if s.charCodeAt(i) == '<' (60)
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    push_const(chunk, Value::F64(60.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    // j = s.indexOf('>', i)
    lget(chunk, s_slot, line);
    push_str(chunk, ">", line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "indexOf");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, j_slot, line);

    // if j < 0: append the rest and finish; else process the tag.
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // out += s.substring(i, n); i = n
    lget(chunk, out_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    lget(chunk, n_slot, line);
    lset(chunk, i_slot, line);
    chunk.emit_else(line);

    // k = i + 1; if s.charCodeAt(k) == '/' (47): k += 1
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, k_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, k_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    push_const(chunk, Value::F64(47.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, k_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, k_slot, line);
    chunk.emit_end(line);

    // name = "" ; read [A-Za-z0-9] from k
    push_str(chunk, "", line);
    lset(chunk, name_slot, line);
    let _ = chunk;
    let inner = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, k_slot, line);
    lget(chunk, n_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, s_slot, line);
    lget(chunk, k_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, cc_slot, line);
    // is_alnum = (48..=57) | (65..=90) | (97..=122)
    // digit
    lget(chunk, cc_slot, line);
    push_const(chunk, Value::F64(48.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    lget(chunk, cc_slot, line);
    push_const(chunk, Value::F64(57.0), line);
    vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    // upper
    lget(chunk, cc_slot, line);
    push_const(chunk, Value::F64(65.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    lget(chunk, cc_slot, line);
    push_const(chunk, Value::F64(90.0), line);
    vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_OR, line);
    // lower
    lget(chunk, cc_slot, line);
    push_const(chunk, Value::F64(97.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    lget(chunk, cc_slot, line);
    push_const(chunk, Value::F64(122.0), line);
    vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if_value(line);
    // alnum: name += s.charAt(k); k += 1
    lget(chunk, name_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, k_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, name_slot, line);
    lget(chunk, k_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, k_slot, line);
    chunk.emit_else(line);
    // non-alnum: stop the name scan
    lget(chunk, n_slot, line);
    lset(chunk, k_slot, line);
    chunk.emit_end(line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, inner, line);
    let chunk = &mut chunks[current];

    // keep = allowed.indexOf("<" + name.toLowerCase() + ">") >= 0
    lget(chunk, allowed_slot, line);
    push_str(chunk, "<", line);
    lget(chunk, name_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "toLowerCase");
        chunk.emit_call(idx, 1, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    push_str(chunk, ">", line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    {
        let idx = chunk.add_import("ecma:string", "indexOf");
        chunk.emit_call(idx, 2, line);
    }
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // keep the whole tag: out += s.substring(i, j+1)
    lget(chunk, out_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    {
        let idx = chunk.add_import("ecma:string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    chunk.emit_end(line);
    // i = j + 1
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_end(line); // end j<0 else

    chunk.emit_else(line); // char is not '<'
    // out += s.charAt(i); i += 1
    lget(chunk, out_slot, line);
    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_end(line); // end if '<'
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, outer, line);
    let chunk = &mut chunks[current];
    lget(chunk, out_slot, line);
}
