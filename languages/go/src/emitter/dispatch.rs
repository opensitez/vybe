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

/// Stable definition addresses for this exact Go producer; retain existing
/// sort recipes when adding encoding recipes. The published slice is immutable.
pub fn helper_definitions() -> &'static [vybe_compiler::primitives::helper_templates::Definition] {
    use vybe_compiler::primitives::helper_templates::Definition;
    static DEFINITIONS: std::sync::OnceLock<Vec<Definition>>=std::sync::OnceLock::new();
    DEFINITIONS.get_or_init(|| {
        super::sort_adapter::HELPER_DEFINITIONS.iter()
            .chain(super::encoding_adapter::HELPER_DEFINITIONS.iter())
            .chain(super::maps_adapter::HELPER_DEFINITIONS.iter())
            .map(|d|Definition {
                contract:d.contract,operation:d.operation,namespace:d.namespace,version:d.version,
                parser:d.parser,producer:d.producer,arity:d.arity,helper_name:d.helper_name,
                factory:d.factory,caller:d.caller,
            }).collect()
    }).as_slice()
}
