use vybe_bytecode::Chunk;

use super::support::emit_null_result;

pub fn emit_release(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_null_result(chunks, current, argc, line);
}

pub fn emit_return_record(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_null_result(chunks, current, argc, line);
}

pub fn emit_sort(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_null_result(chunks, current, argc, line);
}

pub fn emit_merge(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_null_result(chunks, current, argc, line);
}