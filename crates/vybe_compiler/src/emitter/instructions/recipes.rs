use vybe_bytecode::opcode::Op;
use vybe_bytecode::Chunk;

use super::host::{EmitRegistry, FunctionRegistry};

pub fn register_all(_fns: &FunctionRegistry, emits: &mut EmitRegistry) {
    emits.register("common:is_object", is_object_fn);
    emits.register("common:is_func", is_func_fn);
    emits.register("common:string_reverse", string_reverse_fn);
}

fn is_object_fn(fns: &FunctionRegistry, c: &mut Chunk, line: u32) {
    // typeof(null) === "object" per spec, but null is NOT an object.
    // Guard: dup → ref.is_null → if null { drop; false } else { typeof == "object" }
    c.emit_dup(line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if(line);
    c.emit_op(Op::DROP, line);
    super::core_wasm::bool_const(c, line, false);
    c.emit_else(line);
    fns.emit(c, "ecma:value", "typeof", 1, line);
    c.emit_string_const("object", line);
    fns.emit(c, "wasm:js-string", "equals", 2, line);
    c.emit_end(line);
}

fn is_func_fn(fns: &FunctionRegistry, c: &mut Chunk, line: u32) {
    fns.emit(c, "ecma:value", "typeof", 1, line);
    c.emit_string_const("function", line);
    fns.emit(c, "wasm:js-string", "equals", 2, line);
}

fn string_reverse_fn(fns: &FunctionRegistry, c: &mut Chunk, line: u32) {
    c.emit_string_const("", line);
    fns.emit(c, "ecma:string", "split", 2, line);
    fns.emit(c, "ecma:array", "reverse", 1, line);
    c.emit_string_const("", line);
    fns.emit(c, "ecma:array", "join", 2, line);
}

pub fn is_object(c: &mut Chunk, line: u32) {
    let ctx = super::host::CapabilityContext::get();
    is_object_fn(&ctx.functions, c, line);
}

pub fn is_func(c: &mut Chunk, line: u32) {
    let ctx = super::host::CapabilityContext::get();
    is_func_fn(&ctx.functions, c, line);
}

pub fn string_reverse(c: &mut Chunk, line: u32) {
    let ctx = super::host::CapabilityContext::get();
    string_reverse_fn(&ctx.functions, c, line);
}
