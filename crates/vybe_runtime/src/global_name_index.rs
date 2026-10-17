//! Mutation-observed global names and their derived operand-name table.
//!
//! Every mutable map borrow invalidates the derived table before access is
//! granted. This deliberately permits remapping, deletion, aliases, and holes;
//! no append-only assumption or hash-map prefix scan is needed on cache hits.
use std::collections::HashMap;
use std::ops::{Deref, DerefMut};
use std::sync::{Arc, Mutex};
use crate::segmented_global_names::SegmentedGlobalNames;

#[derive(Clone, Debug)]
struct SegmentedCache {
    slot_count: usize,
    names: SegmentedGlobalNames,
    append_safe: bool,
}

#[derive(Debug, Default)]
pub struct GlobalNameIndex {
    names: HashMap<String, u32>,
    operand_table: Mutex<Option<(usize, Arc<Vec<String>>)>>,
    segmented: Mutex<Option<SegmentedCache>>,
}

impl Clone for GlobalNameIndex {
    fn clone(&self) -> Self {
        Self {
            names: self.names.clone(),
            operand_table: Mutex::new(self.operand_table.lock().unwrap().clone()),
            segmented: Mutex::new(self.segmented.lock().unwrap().clone()),
        }
    }
}

impl GlobalNameIndex {
    /// Observe an exact insertion. Only a new, unique trailing slot may extend
    /// a published table without consulting the authoritative rebuild policy.
    pub fn insert(&mut self, name: String, slot: u32) -> Option<u32> {
        self.operand_table.get_mut().unwrap().take();
        let cache = self.segmented.get_mut().unwrap();
        match self.names.entry(name) {
            std::collections::hash_map::Entry::Occupied(mut entry) => {
                *cache = None;
                Some(entry.insert(slot))
            }
            std::collections::hash_map::Entry::Vacant(entry) => {
                if let Some(current) = cache {
                    if current.append_safe && slot as usize == current.slot_count {
                        current.names.push(entry.key().clone());
                        current.slot_count += 1;
                    } else {
                        *cache = None;
                    }
                }
                entry.insert(slot);
                None
            }
        }
    }

    /// A stable snapshot with an exact rebuild fallback for arbitrary map
    /// edits or independent slot-count changes. The builder must return the
    /// authoritative slot-name table, including holes and alias resolution.
    pub fn cached_segmented_table(
        &self,
        slot_count: usize,
        build: impl FnOnce() -> Vec<String>,
    ) -> SegmentedGlobalNames {
        let mut cache = self.segmented.lock().unwrap();
        if let Some(current) = &*cache {
            if current.slot_count == slot_count { return current.names.clone(); }
        }
        let names: SegmentedGlobalNames = build().into_iter().collect();
        // Existing aliases can change winner when HashMap rehashes on insert.
        // Such maps must rebuild rather than extending a previous decision.
        let mut occupied = vec![false; slot_count];
        let append_safe = names.len() == slot_count && self.names.values().all(|slot| {
            let Some(seen) = occupied.get_mut(*slot as usize) else { return false; };
            let unique = !*seen;
            *seen = true;
            unique
        });
        *cache = Some(SegmentedCache { slot_count, names: names.clone(), append_safe });
        names
    }

    /// The caller retains authority over hole/alias spelling and slot order.
    /// Only a stable map's derived result is cached. No guest values are held.
    pub fn cached_operand_table(
        &self,
        slot_count: usize,
        build: impl FnOnce() -> Vec<String>,
    ) -> Arc<Vec<String>> {
        let mut cache = self.operand_table.lock().unwrap();
        if let Some((count, table)) = &*cache {
            if *count == slot_count { return table.clone(); }
        }
        let table = Arc::new(build());
        *cache = Some((slot_count, table.clone()));
        table
    }
}

impl From<HashMap<String, u32>> for GlobalNameIndex {
    fn from(names: HashMap<String, u32>) -> Self {
        Self { names, operand_table: Mutex::new(None), segmented: Mutex::new(None) }
    }
}

impl Deref for GlobalNameIndex {
    type Target = HashMap<String, u32>;
    fn deref(&self) -> &Self::Target { &self.names }
}

impl DerefMut for GlobalNameIndex {
    fn deref_mut(&mut self) -> &mut Self::Target {
        self.operand_table.get_mut().unwrap().take();
        self.segmented.get_mut().unwrap().take();
        &mut self.names
    }
}

impl<'a> IntoIterator for &'a GlobalNameIndex {
    type Item = (&'a String, &'a u32);
    type IntoIter = std::collections::hash_map::Iter<'a, String, u32>;
    fn into_iter(self) -> Self::IntoIter { self.names.iter() }
}

impl<'a> IntoIterator for &'a mut GlobalNameIndex {
    type Item = (&'a String, &'a mut u32);
    type IntoIter = std::collections::hash_map::IterMut<'a, String, u32>;
    fn into_iter(self) -> Self::IntoIter {
        self.operand_table.get_mut().unwrap().take();
        self.segmented.get_mut().unwrap().take();
        self.names.iter_mut()
    }
}

impl IntoIterator for GlobalNameIndex {
    type Item = (String, u32);
    type IntoIter = std::collections::hash_map::IntoIter<String, u32>;
    fn into_iter(self) -> Self::IntoIter { self.names.into_iter() }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn table(index: &GlobalNameIndex) -> Arc<Vec<String>> {
        let length = index.values().copied().max().map_or(0, |slot| slot as usize + 1);
        index.cached_operand_table(length, || {
            let mut output = vec![String::new(); length];
            // Test oracle only: make alias resolution deterministic. Production
            // callers supply their existing authoritative construction policy.
            let mut names: Vec<_> = index.iter().collect();
            names.sort_by_key(|(name, _)| *name);
            for (name, slot) in names { output[*slot as usize] = name.clone(); }
            output
        })
    }

    #[test]
    fn unchanged_map_reuses_the_same_table_without_rebuilding() {
        let index = GlobalNameIndex::default();
        let first = table(&index);
        let again = index.cached_operand_table(0, || panic!("cache hit must not build"));
        assert!(Arc::ptr_eq(&first, &again));
    }

    #[test]
    fn every_public_mutation_invalidates_before_the_next_snapshot() {
        let mut index = GlobalNameIndex::default();
        index.insert("a".into(), 0);
        let first = table(&index);
        index.insert("a".into(), 3);
        let remapped = table(&index);
        assert!(!Arc::ptr_eq(&first, &remapped));
        assert_eq!(&*remapped, &["", "", "", "a"]);
        index.insert("z".into(), 3);
        assert_eq!(table(&index)[3], "z");
        index.remove("z");
        assert_eq!(table(&index)[3], "a");
        *index.get_mut("a").unwrap() = 1;
        assert_eq!(&*table(&index), &["", "a"]);
        index.clear();
        assert!(table(&index).is_empty());
        assert_eq!(&*first, &["a"], "old snapshots remain immutable");
    }

    #[test]
    fn cloned_maps_share_the_view_until_one_copy_mutates() {
        let mut original = GlobalNameIndex::default();
        original.insert("a".into(), 0);
        let first = table(&original);
        let mut fork = original.clone();
        assert!(Arc::ptr_eq(&first, &table(&fork)));
        for (_, slot) in &mut fork { *slot = 2; }
        assert_eq!(&*table(&fork), &["", "", "a"]);
        assert!(Arc::ptr_eq(&first, &table(&original)));
    }

    #[test]
    fn independent_slot_growth_and_shrink_rebuild_the_bounded_cache() {
        let index = GlobalNameIndex::default();
        let initial = index.cached_operand_table(1, || vec!["".into()]);
        let grown = index.cached_operand_table(3, || vec![String::new(); 3]);
        assert_eq!(grown.len(), 3);
        assert!(!Arc::ptr_eq(&initial, &grown));
        let reused = index.cached_operand_table(3, || panic!("same slot count must reuse"));
        assert!(Arc::ptr_eq(&grown, &reused));
        let shrunk = index.cached_operand_table(0, Vec::new);
        assert!(shrunk.is_empty());
        assert_eq!(initial.len(), 1, "published views remain immutable");
    }

    fn segmented(index: &GlobalNameIndex, count: usize) -> SegmentedGlobalNames {
        index.cached_segmented_table(count, || {
            let mut names = vec![String::new(); count];
            let mut entries: Vec<_> = index.iter().collect();
            entries.sort_by_key(|(name, _)| *name);
            for (name, slot) in entries { names[*slot as usize] = name.clone(); }
            names
        })
    }

    #[test]
    fn unique_trailing_inserts_extend_without_running_the_rebuild() {
        let mut index = GlobalNameIndex::default();
        for slot in 0..300 { index.insert(format!("name_{slot}"), slot); }
        let original = segmented(&index, 300);
        for slot in 300..500 {
            index.insert(format!("name_{slot}"), slot);
            let latest = index.cached_segmented_table(slot as usize + 1,
                || panic!("trailing insert must not rebuild all names"));
            assert_eq!(latest.get(slot as usize).map(String::as_str), Some(format!("name_{slot}").as_str()));
            assert!(original.is_prefix_of(&latest));
        }
        assert_eq!(original.len(), 300);
    }

    #[test]
    fn aliases_remaps_and_mutable_borrows_use_the_exact_fallback() {
        let mut index = GlobalNameIndex::default();
        index.insert("a".into(), 0);
        let original = segmented(&index, 1);
        index.insert("z".into(), 0);
        assert_eq!(segmented(&index, 1).get(0).map(String::as_str), Some("z"));
        index.insert("trailing".into(), 1);
        let latest = segmented(&index, 2);
        assert_eq!(latest.get(0).map(String::as_str), Some("z"));
        index.remove("z");
        assert_eq!(segmented(&index, 2).get(0).map(String::as_str), Some("a"));
        *index.get_mut("a").unwrap() = 2;
        let remapped = segmented(&index, 3);
        assert_eq!(remapped.get(0).map(String::as_str), Some(""));
        assert_eq!(remapped.get(2).map(String::as_str), Some("a"));
        assert_eq!(original.get(0).map(String::as_str), Some("a"));
    }
}
