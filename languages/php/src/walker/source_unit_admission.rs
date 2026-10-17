//! Conservative admission for the optional source-unit receipt path.
//! This does not recognize PHP expressions, resolve a path, or evaluate a load.

pub(super) fn may_contain_source_load(source: &str) -> bool {
    let bytes = source.as_bytes();
    bytes.iter().enumerate().any(|(offset, byte)| {
        let stem: &[u8] = match byte {
            b'i' | b'I' => b"include",
            b'r' | b'R' => b"require",
            _ => return false,
        };
        bytes.get(offset..offset + stem.len())
            .is_some_and(|candidate| candidate.eq_ignore_ascii_case(stem))
    })
}
