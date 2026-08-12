//! JVM Base64 adapter glue.
//!
//! The Java-family surfaces (`java.util.Base64`, Kotlin calls into `java.*`,
//! etc.) should not spell ECMA host imports directly. They route through this
//! platform layer, which delegates the actual binary-string codec to the
//! shared compiler primitive.

use vybe_runtime::Chunk;

pub fn emit_encode_binary_string(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::base64::emit_encode_binary_string(chunks, current, line);
}

pub fn emit_decode_binary_string(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::base64::emit_decode_binary_string(chunks, current, line);
}
