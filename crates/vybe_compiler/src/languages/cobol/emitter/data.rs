use vybe_bytecode::Chunk;

use super::support::emit_null_result;

pub fn emit_copy(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_null_result(chunks, current, argc, line);
}

pub fn emit_validate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_null_result(chunks, current, argc, line);
}

pub fn emit_typedef(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_null_result(chunks, current, argc, line);
}

pub fn emit_move_corresponding(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_null_result(chunks, current, argc, line);
}