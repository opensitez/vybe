use vybe_runtime::Chunk;

pub fn dispatch(
    _name: &str,
    _chunks: &mut Vec<Chunk>,
    _current: usize,
    _argc: u8,
    _line: u32,
) -> bool {
    false
}
