//! Go `encoding/*` adapter helpers that need byte-level loops.

use vybe_runtime::Chunk;

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "go.hex.EncodeToString" if argc == 1 => emit_hex_encode_to_string(chunks, current, line),
        "go.hex.DecodeStringBytes" if argc == 1 => {
            emit_hex_decode_string_bytes(chunks, current, line)
        }
        _ => return false,
    }
    true
}

fn emit_hex_encode_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::base64::emit_byte_array_to_binary_string(chunks, current, line);
    vybe_compiler::primitives::string_encoding::emit_bin2hex(chunks, current, 1, line);
}

fn emit_hex_decode_string_bytes(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::string_encoding::emit_hex2bin(chunks, current, 1, line);
    vybe_compiler::primitives::base64::emit_binary_string_to_byte_array(chunks, current, line);
}
