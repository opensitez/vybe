//! Persistent slot-name snapshots for include/link metadata.
//!
//! Appending to a published snapshot copies block pointers and at most one
//! bounded tail of string references, never the complete string payload set.
//! These are host metadata, not guest objects or Wasm instructions.
use std::sync::{Arc, OnceLock};
use std::ops::{Deref, Index, RangeFull};

const BLOCK_SIZE: usize = 128;
type Block = Vec<Arc<String>>;

#[derive(Default)]
pub struct SegmentedGlobalNames {
    blocks: Arc<Vec<Arc<Block>>>,
    len: usize,
    flat: OnceLock<Vec<String>>,
}

impl Clone for SegmentedGlobalNames {
    fn clone(&self) -> Self {
        Self { blocks: self.blocks.clone(), len: self.len, flat: OnceLock::new() }
    }
}

impl std::fmt::Debug for SegmentedGlobalNames {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_list().entries(self.iter()).finish()
    }
}

impl PartialEq for SegmentedGlobalNames {
    fn eq(&self, other: &Self) -> bool {
        self.len == other.len && self.iter().eq(other.iter())
    }
}
impl Eq for SegmentedGlobalNames {}
impl PartialEq<Vec<String>> for SegmentedGlobalNames {
    fn eq(&self, other: &Vec<String>) -> bool { self.iter().eq(other.iter()) }
}
impl PartialEq<SegmentedGlobalNames> for Vec<String> {
    fn eq(&self, other: &SegmentedGlobalNames) -> bool { other == self }
}

impl Deref for SegmentedGlobalNames {
    type Target = [String];
    fn deref(&self) -> &Self::Target { self.as_slice() }
}
impl Index<usize> for SegmentedGlobalNames {
    type Output = String;
    fn index(&self, index: usize) -> &String { self.get(index).expect("global name index out of bounds") }
}
impl Index<RangeFull> for SegmentedGlobalNames {
    type Output = [String];
    fn index(&self, _: RangeFull) -> &[String] { self.as_slice() }
}

pub struct NamesIter<'a> {
    names: &'a SegmentedGlobalNames,
    front: usize,
    back: usize,
}
impl<'a> Iterator for NamesIter<'a> {
    type Item = &'a String;
    fn next(&mut self) -> Option<Self::Item> {
        if self.front == self.back { return None; }
        let value = self.names.get(self.front);
        self.front += 1;
        value
    }
    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.back - self.front;
        (remaining, Some(remaining))
    }
}
impl DoubleEndedIterator for NamesIter<'_> {
    fn next_back(&mut self) -> Option<Self::Item> {
        if self.front == self.back { return None; }
        self.back -= 1;
        self.names.get(self.back)
    }
}
impl ExactSizeIterator for NamesIter<'_> {}
impl<'a> IntoIterator for &'a SegmentedGlobalNames {
    type Item = &'a String;
    type IntoIter = NamesIter<'a>;
    fn into_iter(self) -> Self::IntoIter { self.iter() }
}

impl SegmentedGlobalNames {
    pub fn len(&self) -> usize { self.len }
    pub fn is_empty(&self) -> bool { self.len == 0 }

    pub fn get(&self, slot: usize) -> Option<&String> {
        if slot >= self.len { return None; }
        Some(self.blocks[slot / BLOCK_SIZE][slot % BLOCK_SIZE].as_ref())
    }

    pub fn iter(&self) -> NamesIter<'_> {
        NamesIter { names: self, front: 0, back: self.len }
    }

    /// Compatibility for cold consumers that require an actual slice. Indexed
    /// and iterator access never invokes this lazy flattening path.
    pub fn as_slice(&self) -> &[String] {
        self.flat.get_or_init(|| self.iter().cloned().collect())
    }

    pub fn push(&mut self, name: impl Into<String>) {
        self.push_shared(Arc::new(name.into()));
    }

    fn push_shared(&mut self, name: Arc<String>) {
        self.flat.take();
        let blocks = Arc::make_mut(&mut self.blocks);
        if self.len % BLOCK_SIZE == 0 {
            blocks.push(Arc::new(Vec::with_capacity(BLOCK_SIZE)));
        }
        Arc::make_mut(blocks.last_mut().unwrap()).push(name);
        self.len += 1;
    }

    /// Replace an existing slot without changing any published snapshot.
    pub fn set(&mut self, slot: usize, name: impl Into<String>) {
        assert!(slot < self.len, "global name slot out of bounds");
        self.flat.take();
        let blocks = Arc::make_mut(&mut self.blocks);
        Arc::make_mut(&mut blocks[slot / BLOCK_SIZE])[slot % BLOCK_SIZE] = Arc::new(name.into());
    }

    /// Preserve unnamed global slots as empty names, including trailing holes.
    pub fn resize(&mut self, len: usize) {
        if len <= self.len {
            self.truncate(len);
            return;
        }
        static EMPTY: OnceLock<Arc<String>> = OnceLock::new();
        let empty = EMPTY.get_or_init(|| Arc::new(String::new()));
        while self.len < len { self.push_shared(empty.clone()); }
    }

    pub fn truncate(&mut self, len: usize) {
        if len >= self.len { return; }
        self.flat.take();
        let blocks = Arc::make_mut(&mut self.blocks);
        blocks.truncate(len.div_ceil(BLOCK_SIZE));
        let tail = len % BLOCK_SIZE;
        if tail != 0 { Arc::make_mut(blocks.last_mut().unwrap()).truncate(tail); }
        self.len = len;
    }

    /// Exact textual prefix check; shared complete blocks require no string
    /// comparisons. Aliases/remaps are never accepted by length alone.
    pub fn is_prefix_of(&self, other: &Self) -> bool {
        if self.len > other.len { return false; }
        for (index, block) in self.blocks.iter().enumerate() {
            let other_block = &other.blocks[index];
            if Arc::ptr_eq(block, other_block) { continue; }
            if block.len() > other_block.len()
                || !block.iter().zip(other_block.iter()).all(|(a, b)| a == b)
            { return false; }
        }
        true
    }
}

impl FromIterator<String> for SegmentedGlobalNames {
    fn from_iter<I: IntoIterator<Item = String>>(names: I) -> Self {
        let mut output = Self::default();
        for name in names { output.push(name); }
        output
    }
}

impl From<Vec<String>> for SegmentedGlobalNames {
    fn from(names: Vec<String>) -> Self { names.into_iter().collect() }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn appends_share_full_blocks_and_all_existing_string_payloads() {
        let mut table: SegmentedGlobalNames = (0..300).map(|i| format!("global_{i}")).collect();
        let before = table.clone();
        table.push("new");
        assert!(Arc::ptr_eq(&before.blocks[0], &table.blocks[0]));
        assert!(Arc::ptr_eq(&before.blocks[1], &table.blocks[1]));
        assert!(!Arc::ptr_eq(&before.blocks[2], &table.blocks[2]));
        for slot in 0..before.len() {
            assert!(Arc::ptr_eq(
                &before.blocks[slot / BLOCK_SIZE][slot % BLOCK_SIZE],
                &table.blocks[slot / BLOCK_SIZE][slot % BLOCK_SIZE],
            ));
        }
        assert_eq!(before.get(300), None);
        assert_eq!(table.get(300).map(String::as_str), Some("new"));
        assert!(before.is_prefix_of(&table));
        assert!(!table.is_prefix_of(&before));
    }

    #[test]
    fn holes_remaps_shrink_and_regrowth_preserve_exact_snapshots() {
        let mut table = SegmentedGlobalNames::default();
        table.resize(260);
        table.set(0, "namespace.CaseSensitive");
        table.set(129, "unicode_\u{e9}");
        let before = table.clone();
        table.set(129, "remapped");
        assert!(!before.is_prefix_of(&table));
        assert_eq!(before.get(129).map(String::as_str), Some("unicode_\u{e9}"));
        table.truncate(129);
        assert!(table.is_prefix_of(&before));
        table.resize(261);
        assert_eq!(table.get(129).map(String::as_str), Some(""));
        assert_eq!(table.get(260).map(String::as_str), Some(""));
        assert_eq!(table.iter().count(), table.len());
        table.truncate(0);
        assert!(table.is_empty());
        assert_eq!(table.get(0), None);
        assert!(table.is_prefix_of(&before));
    }

    #[test]
    fn independently_built_equal_names_compare_by_content() {
        let first: SegmentedGlobalNames = ["a".into(), "b".into()].into_iter().collect();
        let second: SegmentedGlobalNames = ["a".into(), "b".into(), "c".into()].into_iter().collect();
        assert!(first.is_prefix_of(&second));
        let mut changed = second.clone();
        changed.set(0, "A");
        assert!(!first.is_prefix_of(&changed));
    }
}
