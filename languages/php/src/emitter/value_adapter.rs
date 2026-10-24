//! PHP value predicates as guest code over `ecma:*` and `wasm:*` builtins.
use vybe_compiler::primitives::{globals, ops};
use vybe_runtime::{Chunk, opcode::Op};

/// The global holding the compiled numeric-string pattern.
const NUMERIC_PATTERN_GLOBAL: &str = "__php_numeric_string_pattern";

/// php's numeric string: optional surrounding whitespace (space, \t, \n, \r,
/// \v, \f), an optional sign, digits with an optional fraction (at least one
/// digit overall) and an optional exponent with digits — nothing else.
const NUMERIC_PATTERN: &str =
    r"^[ \t\n\r\v\f]*[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?[ \t\n\r\v\f]*$";

/// Push the compiled numeric-string RegExp, compiling it on first use.
fn emit_numeric_pattern(chunk: &mut Chunk, line: u32) {
    let pattern = chunk.alloc_scratch(1);
    globals::emit_read(chunk, NUMERIC_PATTERN_GLOBAL, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, pattern, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, pattern, line);
    let undefined = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(undefined, 1, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    chunk.emit_string_const(NUMERIC_PATTERN, line);
    chunk.emit_string_const("", line);
    let compile = chunk.add_import("ecma:regexp", "new");
    chunk.emit_call(compile, 2, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, pattern, line);
    globals::emit_write(chunk, NUMERIC_PATTERN_GLOBAL, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, pattern, line);
}

/// Stack: `[string] -> [i32]`. Is the string a php numeric string?
pub fn emit_is_numeric_string_condition(chunk: &mut Chunk, line: u32) {
    let text = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, text, line);
    emit_numeric_pattern(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text, line);
    let test = chunk.add_import("ecma:regexp", "test");
    chunk.emit_call(test, 2, line);
    let cast = chunk.add_import("wasm:js-boolean", "cast");
    chunk.emit_call(cast, 1, line);
}

/// Stack: `[value] -> [Boolean]`. php `is_numeric`: a number, or a numeric
/// string.
pub fn emit_is_numeric(chunk: &mut Chunk, line: u32) {
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_TEE, value, line);
    let number = chunk.add_import("wasm:js-number", "test");
    chunk.emit_call(number, 1, line);
    chunk.emit_if_i32(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let string = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(string, 1, line);
    chunk.emit_if_i32(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    emit_is_numeric_string_condition(chunk, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    ops::emit_i32_to_bool(chunk, line);
}

/// Push an i32: do the values in `a` and `b` have the same php truthiness?
pub fn emit_same_truthiness_condition(chunk: &mut Chunk, a: u16, b: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQ, line);
}
