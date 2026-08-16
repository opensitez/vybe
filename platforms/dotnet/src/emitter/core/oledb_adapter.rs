use vybe_runtime::Chunk;

use crate::emitter::core::sqlclient_adapter;

pub fn emit_oledb_connection_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    sqlclient_adapter::emit_connection_new(chunks, current, argc, line);
}

pub fn emit_oledb_command_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    sqlclient_adapter::emit_command_new(chunks, current, argc, line);
}
