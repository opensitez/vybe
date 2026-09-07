//! Go-specific common dispatch.

use vybe_runtime::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    if crate::emitter::core_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    if crate::emitter::fmt_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    if crate::emitter::json_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    if crate::emitter::regexp_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    if crate::emitter::errors_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    if crate::emitter::path_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    if crate::emitter::time_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    if crate::emitter::container_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    if crate::emitter::maps_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    if crate::emitter::sort_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    if crate::emitter::strings_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    if crate::emitter::runtime_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    if crate::emitter::reflection_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    if crate::emitter::binary_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    if crate::emitter::encoding_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    if crate::emitter::url_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    if crate::emitter::netip_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    crate::emitter::math_adapter::emit_helper(name, chunks, current, argc, line)
}
