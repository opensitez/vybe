//! Bounded immutable-string UTF-16 lengths for runtime property reads.
use std::sync::{Arc, Weak};
pub(crate) const MAX_SOURCE_BYTES: usize = 65_536;

/// stringref measure_utf8/measure_wtf8 use -1 above the proposal's limit.
#[inline]
pub(crate) fn measure_utf8_bytes(bytes: usize) -> i32 {
    i32::try_from(bytes).unwrap_or(-1)
}

#[cfg(test)]
mod stringref_measure_tests {
    #[test]
    fn encoding_limit_uses_the_proposal_sentinel() {
        use super::measure_utf8_bytes as measure;
        assert_eq!(measure(0), 0);
        assert_eq!(measure(i32::MAX as usize - 1), i32::MAX - 1);
        assert_eq!(measure(i32::MAX as usize), i32::MAX);
        assert_eq!(measure(i32::MAX as usize + 1), -1);
        assert_eq!(measure(u32::MAX as usize), -1);
        assert_eq!(measure(usize::MAX), -1);
    }
}

/// Count UTF-16 units in valid UTF-8 without allocating a decoded buffer.
#[inline]
pub fn utf16_units(source: &str) -> usize {
    if source.len() >= 128 {
        return utf16_units_wide(source);
    }
    if source.len() < 8 {
        return source.encode_utf16().count();
    }
    let mut chunks = source.as_bytes().chunks_exact(8);
    let mut units = 0;
    for chunk in &mut chunks {
        let word = u64::from_le_bytes(chunk.try_into().unwrap());
        let high = 0x8080_8080_8080_8080u64;
        // Valid UTF-8 has one leading byte per scalar. Continuation bytes
        // have high bits 10; four-byte leaders have high bits 1111 and
        // contribute a second UTF-16 unit. Only each byte's high bit is
        // counted, so shifts cannot mix neighboring byte classifications.
        // Chunk boundaries may split scalars without affecting the count.
        let continuation = word & !(word << 1) & high;
        let astral = word & (word << 1) & (word << 2) & (word << 3) & high;
        units += 8 - continuation.count_ones() as usize + astral.count_ones() as usize;
    }
    units + chunks.remainder().iter().map(|&byte|
        usize::from(byte & 0xc0 != 0x80) + usize::from(byte >= 0xf0)
    ).sum::<usize>()
}

// Keep the uncached bulk scan out of the common short/cached code path.
#[inline(never)]
fn utf16_units_wide(source: &str) -> usize {
    if source.is_ascii() {
        return source.len();
    }
    let mut chunks = source.as_bytes().chunks_exact(16);
    let mut units = 0;
    for chunk in &mut chunks {
        let word = u128::from_le_bytes(chunk.try_into().unwrap());
        let high = 0x8080_8080_8080_8080_8080_8080_8080_8080u128;
        let continuation = word & !(word << 1) & high;
        let astral = word & (word << 1) & (word << 2) & (word << 3) & high;
        units += 16 - continuation.count_ones() as usize + astral.count_ones() as usize;
    }
    let mut chunks = chunks.remainder().chunks_exact(8);
    for chunk in &mut chunks {
        let word = u64::from_le_bytes(chunk.try_into().unwrap());
        let high = 0x8080_8080_8080_8080u64;
        let continuation = word & !(word << 1) & high;
        let astral = word & (word << 1) & (word << 2) & (word << 3) & high;
        units += 8 - continuation.count_ones() as usize + astral.count_ones() as usize;
    }
    units + chunks.remainder().iter().map(|&byte|
        usize::from(byte & 0xc0 != 0x80) + usize::from(byte >= 0xf0)
    ).sum::<usize>()
}
#[derive(Default)]
struct Entry {
    source: Option<Weak<str>>,
    units: usize,
}
#[derive(Default)]
pub(crate) struct StringLengths {
    entries: [Entry; 2],
    recent: usize,
}
impl StringLengths {
    pub(crate) fn count(&mut self, source: &Arc<str>) -> usize {
        if source.len() > MAX_SOURCE_BYTES {
            return utf16_units_wide(source);
        }
        if let Some(slot) = self.entries.iter().position(|entry|
            entry.source.as_ref().is_some_and(|old|
                std::ptr::eq(old.as_ptr(), Arc::as_ptr(source)))) {
            self.recent = slot;
            return self.entries[slot].units;
        }
        let units = utf16_units(source);
        let slot = self.entries.iter().position(|entry|
            entry.source.as_ref().is_none_or(|old| old.strong_count() == 0))
            .unwrap_or(1 - self.recent);
        self.entries[slot] = Entry { source: Some(Arc::downgrade(source)), units };
        self.recent = slot;
        units
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn wide_count_matches_every_scalar_at_every_vector_alignment() {
        for cp in 0..=0x10ffff {
            let Some(ch) = char::from_u32(cp) else { continue };
            let mut encoded = [0; 4];
            let scalar = ch.encode_utf8(&mut encoded);
            for prefix in 0..16 {
                let mut bytes = [b'x'; 288];
                bytes[prefix..prefix + scalar.len()].copy_from_slice(scalar.as_bytes());
                let len = prefix + scalar.len() + 256;
                let source = std::str::from_utf8(&bytes[..len]).unwrap();
                assert_eq!(utf16_units_wide(source), prefix + ch.len_utf16() + 256);
            }
        }
    }
    #[test]
    fn bulk_count_matches_every_scalar_at_every_word_alignment() {
        for cp in 0..=0x10ffff {
            let Some(ch) = char::from_u32(cp) else { continue };
            let mut encoded = [0; 4];
            let scalar = ch.encode_utf8(&mut encoded);
            for prefix in 0..8 {
                let mut bytes = [b'x'; 20];
                bytes[prefix..prefix + scalar.len()].copy_from_slice(scalar.as_bytes());
                let len = prefix + scalar.len() + 8;
                let source = std::str::from_utf8(&bytes[..len]).unwrap();
                assert_eq!(utf16_units(source), prefix + ch.len_utf16() + 8);
            }
        }
    }

    #[test]
    fn bulk_count_preserves_tails_and_mixed_sequences() {
        let source = "A\0Ωé中😀𐐀".repeat(1024);
        for end in 0..=source.len() {
            if source.is_char_boundary(end) {
                let prefix = &source[..end];
                assert_eq!(utf16_units(prefix), prefix.encode_utf16().count());
            }
        }
    }

    #[test]
    fn lengths_preserve_unicode_bounds_eviction_and_weak_roots() {
        let mut cache = StringLengths::default();
        for cp in 0..=0x10ffff {
            let Some(ch) = char::from_u32(cp) else { continue };
            let source: Arc<str> = Arc::from(ch.to_string());
            assert_eq!(cache.count(&source), ch.len_utf16());
            assert_eq!(cache.count(&source), ch.len_utf16());
        }
        let roots: Vec<Arc<str>> = ["", "A\0Ω😀", "Ωé中", "😀A𐐀"].into_iter().map(Arc::from).collect();
        for _ in 0..3 { for root in &roots { assert_eq!(cache.count(root), root.encode_utf16().count()); } }
        let root: Arc<str> = Arc::from("Ω".repeat(32768));
        assert_eq!(cache.count(&root), 32768);
        let weak = Arc::downgrade(&root);
        assert_eq!(Arc::strong_count(&root), 1);
        drop(root);
        assert!(weak.upgrade().is_none());
        let oversized: Arc<str> = Arc::from("Ω".repeat(32769));
        assert_eq!(cache.count(&oversized), 32769);
        assert!(!cache.entries.iter().any(|entry| entry.source.as_ref().is_some_and(|old|
            std::ptr::eq(old.as_ptr(), Arc::as_ptr(&oversized)))));
        assert!(cache.entries.iter().filter(|entry| entry.source.is_some()).count() <= 2);
    }
}
