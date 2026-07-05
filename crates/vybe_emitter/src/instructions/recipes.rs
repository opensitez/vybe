use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use super::host::{EmitRegistry, FunctionRegistry};

pub fn register_all(_fns: &FunctionRegistry, emits: &mut EmitRegistry) {
    emits.register("common:is_object", is_object_fn);
    emits.register("common:is_func", is_func_fn);
}

fn is_object_fn(_fns: &FunctionRegistry, c: &mut Chunk, line: u32) {
    // WASM GC ref.test with type "object" — single opcode, null-safe
    let type_idx = c.add_constant(Value::String(Arc::from("object")));
    c.emit_op_u16(Op::REF_TEST, type_idx, line);
}

fn is_func_fn(_fns: &FunctionRegistry, c: &mut Chunk, line: u32) {
    let type_idx = c.add_constant(Value::String(Arc::from("func")));
    c.emit_op_u16(Op::REF_TEST, type_idx, line);
}

pub fn is_object(c: &mut Chunk, line: u32) {
    let type_idx = c.add_constant(Value::String(Arc::from("object")));
    c.emit_op_u16(Op::REF_TEST, type_idx, line);
}

pub fn is_func(c: &mut Chunk, line: u32) {
    let type_idx = c.add_constant(Value::String(Arc::from("func")));
    c.emit_op_u16(Op::REF_TEST, type_idx, line);
}
