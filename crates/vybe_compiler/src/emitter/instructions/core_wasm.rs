use vybe_bytecode::{Chunk, Op};
use vybe_bytecode::Value;

pub fn dup(c: &mut Chunk, line: u32) {
    c.emit_dup(line);
}

pub fn i32_const(c: &mut Chunk, line: u32, v: i32) {
    c.emit_i32_const(v, line);
}

pub fn i64_const(c: &mut Chunk, line: u32, v: i64) {
    c.emit_i64_const(v, line);
}

pub fn f64_const(c: &mut Chunk, line: u32, v: f64) {
    c.emit_f64_const(v, line);
}

pub fn bool_const(c: &mut Chunk, line: u32, v: bool) {
    c.emit_bool_const(v, line);
}

pub fn undefined(c: &mut Chunk, line: u32) {
    let idx = c.add_constant(Value::String(std::sync::Arc::from("undefined")));
    c.emit_op_u16(Op::GLOBAL_GET, idx, line);
}

pub fn null(c: &mut Chunk, line: u32) {
    c.emit_op(Op::NULL, line);
}

pub fn string_const(c: &mut Chunk, line: u32, s: &str) {
    c.emit_string_const(s, line);
}

pub fn emit_value(c: &mut Chunk, line: u32, val: &Value) {
    match val {
        Value::I32(v) => c.emit_i32_const(*v, line),
        Value::I64(v) => c.emit_i64_const(*v, line),
        Value::F64(v) => c.emit_f64_const(*v, line),
        Value::Bool(v) => c.emit_bool_const(*v, line),
        Value::Null => c.emit_op(Op::NULL, line),
        Value::Undefined => undefined(c, line),
        Value::String(s) => c.emit_string_const(s, line),
        Value::BigInt(v) => c.emit_i64_const(*v, line),
        Value::V128(v) => {
            c.emit_op(Op::V128_CONST, line);
            for b in v { c.emit(*b, line); }
        }
        other => panic!("emit_value: no WASM-compliant encoding for {:?}", other),
    }
}
