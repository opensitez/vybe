use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

const PATTERN_KEY: &str = "__pattern";
const VALUE_KEY: &str = "value";
const COUNT_KEY: &str = "count";

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    let idx = chunk.add_constant(val);
    chunk.emit_op_u16(Op::CONST, idx, line);
}

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    let slot = chunk.local_count;
    chunk.local_count = slot + 1;
    slot
}

pub fn emit_regex_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let pattern_key = chunk.add_constant(Value::String(Arc::from(PATTERN_KEY)));
    match argc {
        0 => {
            chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
            chunk.emit_op(Op::DUP, line);
            push_const(chunk, Value::String(Arc::from("")), line);
            chunk.emit_op_u16(Op::STRUCT_SET, pattern_key, line);
            chunk.emit_op(Op::DROP, line);
        }
        _ => {
            for _ in 1..argc {
                chunk.emit_op(Op::DROP, line);
            }
            let pattern_slot = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, pattern_slot, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
            chunk.emit_op(Op::DUP, line);
            chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
            chunk.emit_op_u16(Op::STRUCT_SET, pattern_key, line);
            chunk.emit_op(Op::DROP, line);
        }
    }
}

pub fn emit_regex_is_match(chunks: &mut [Chunk], current: usize, line: u32) {
    let test_idx = chunks[0].add_import("ecma:regexp", "test");
    let chunk = &mut chunks[current];
    let input_slot = reserve_slot(chunk);
    let self_slot = reserve_slot(chunk);
    let pattern_key = chunk.add_constant(Value::String(Arc::from(PATTERN_KEY)));

    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, pattern_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, test_idx, line);
    chunk.emit(2, line);
}

pub fn emit_regex_replace(chunks: &mut [Chunk], current: usize, line: u32) {
    let replace_idx = chunks[0].add_import("ecma:regexp", "replaceAll");
    let chunk = &mut chunks[current];
    let replacement_slot = reserve_slot(chunk);
    let input_slot = reserve_slot(chunk);
    let self_slot = reserve_slot(chunk);
    let pattern_key = chunk.add_constant(Value::String(Arc::from(PATTERN_KEY)));

    chunk.emit_op_u16(Op::LOCAL_SET, replacement_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, pattern_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, replacement_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, replace_idx, line);
    chunk.emit(3, line);
}

pub fn emit_regex_split(chunks: &mut [Chunk], current: usize, line: u32) {
    let split_idx = chunks[0].add_import("ecma:regexp", "split");
    let chunk = &mut chunks[current];
    let input_slot = reserve_slot(chunk);
    let self_slot = reserve_slot(chunk);
    let pattern_key = chunk.add_constant(Value::String(Arc::from(PATTERN_KEY)));

    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, pattern_key, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, split_idx, line);
    chunk.emit(2, line);
}

pub fn emit_regex_match(chunks: &mut [Chunk], current: usize, line: u32) {
    let exec_idx = chunks[0].add_import("ecma:regexp", "exec");
    let chunk = &mut chunks[current];
    let input_slot = reserve_slot(chunk);
    let self_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);
    let obj_slot = reserve_slot(chunk);
    let pattern_key = chunk.add_constant(Value::String(Arc::from(PATTERN_KEY)));
    let value_key = chunk.add_constant(Value::String(Arc::from(VALUE_KEY)));

    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, pattern_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, exec_idx, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::STRUCT_SET, value_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

pub fn emit_regex_matches(chunks: &mut [Chunk], current: usize, line: u32) {
    let match_all_idx = chunks[0].add_import("ecma:regexp", "matchAll");
    let chunk = &mut chunks[current];
    let input_slot = reserve_slot(chunk);
    let self_slot = reserve_slot(chunk);
    let count_slot = reserve_slot(chunk);
    let obj_slot = reserve_slot(chunk);
    let pattern_key = chunk.add_constant(Value::String(Arc::from(PATTERN_KEY)));
    let count_key = chunk.add_constant(Value::String(Arc::from(COUNT_KEY)));

    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, self_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, pattern_key, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, match_all_idx, line);
    chunk.emit(2, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, count_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}
