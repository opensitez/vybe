use vybe_bytecode::Chunk;

/// Lua-specific `common:lua.*` hooks only when walker normalization cannot
/// reach the shared emitter (`loops`, `expressions`, `strings`, `collections`,
/// `io`, …). Control flow and builtins use the same compiler + profile paths as JS.
pub fn dispatch(
    _name: &str,
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) -> bool {
    false
}
