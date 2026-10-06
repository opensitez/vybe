// Exact factory frozen2026-10-05 before proposed predicate rewrite.
use vybe_runtime::{Chunk, Op};
pub fn build_is_numeric(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_isnumeric");
    c.arity = 1;
    c.local_count = 2; // val(0), result(1)
    let val = 0u16;
    let result = 1u16;

    // VB `IsNumeric(v)`:
    //   typeof(v) ∈ {"number", "i32", "i64"}                → true
    //   typeof(v) == "string" && !isNaN(parseFloat(v))      → true
    //   otherwise                                            → false
    //
    // Block-and-br_if cascade so each positive case short-circuits and
    // the next check is skipped.
    let num_str = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("number")));
    let i32_str = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("i32")));
    let i64_str = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("i64")));
    let str_str = c.add_constant(vybe_runtime::Value::String(std::sync::Arc::from("string")));

    let done = c.emit_block(0);

    // typeof(v) == "number"
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    frozen_typeof(&mut c, 0);
    frozen_const_index(&mut c, num_str, 0);
    frozen_equals(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_br_if(0, 0);

    // typeof(v) == "i32"
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    frozen_typeof(&mut c, 0);
    frozen_const_index(&mut c, i32_str, 0);
    frozen_equals(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_br_if(0, 0);

    // typeof(v) == "i64"
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    frozen_typeof(&mut c, 0);
    frozen_const_index(&mut c, i64_str, 0);
    frozen_equals(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_br_if(0, 0);

    // typeof(v) == "string" — try parseFloat, accept iff !isNaN
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    frozen_typeof(&mut c, 0);
    frozen_const_index(&mut c, str_str, 0);
    frozen_equals(&mut c, 0);
    // [is_string]
    frozen_not(imports, &mut c, 0);
    c.emit_br_if(0, 0); // not a string → done with result still false

    // result = !isNaN(parseFloat(v))  ≡  parsed == parsed
    let pf_idx = c.add_import("ecma:number", "parseFloat");
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_call(pf_idx, 1, 0);
    c.emit_dup(0);
    c.emit_op(Op::F64_EQ, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    c.emit_end(0);
    c.patch_block(done);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn frozen_typeof(c: &mut Chunk, line: u32) {
    let i = c.add_import("ecma:value", "typeof");
    c.emit_call(i, 1, line);
}
fn frozen_equals(c: &mut Chunk, line: u32) {
    let i = c.add_import("wasm:js-string", "equals");
    c.emit_call(i, 2, line);
}
fn frozen_const_index(c: &mut Chunk, index: u32, line: u32) {
    let vybe_runtime::Value::String(s) = c.constants[index as usize].clone() else {
        panic!("frozen string constant")
    };
    c.emit_string_const(&s, line);
}
fn frozen_not(_: &mut Chunk, c: &mut Chunk, line: u32) {
    let i = c.add_import("ecma:boolean", "toBoolean");
    let j = c.add_import("wasm:js-boolean", "cast");
    c.emit_call(i, 1, line);
    c.emit_call(j, 1, line);
    c.emit_op(Op::I32_EQZ, line);
}
