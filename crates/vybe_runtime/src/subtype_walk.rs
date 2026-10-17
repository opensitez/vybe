//! Scratch storage for the native RTT graph walk, not a type-ID cache.
use super::TypeRegistry;
use std::collections::HashSet;

struct Pending {
    inline: [usize; 8],
    len: usize,
    spill: Vec<usize>,
}

impl Pending {
    fn new() -> Self { Self { inline: [0; 8], len: 0, spill: Vec::new() } }

    #[inline]
    fn push(&mut self, id: usize) {
        if self.len < self.inline.len() && self.spill.is_empty() {
            self.inline[self.len] = id;
            self.len += 1;
        } else { self.spill.push(id); }
    }

    #[inline]
    fn extend(&mut self, ids: &[usize]) {
        if ids.is_empty() { return; }
        let count = if self.spill.is_empty() {
            ids.len().min(self.inline.len() - self.len)
        } else { 0 };
        self.inline[self.len..self.len + count].copy_from_slice(&ids[..count]);
        self.len += count;
        self.spill.extend_from_slice(&ids[count..]);
    }

    #[inline]
    fn pop(&mut self) -> Option<usize> {
        if let Some(id) = self.spill.pop() { return Some(id); }
        if self.len == 0 { return None; }
        self.len -= 1;
        Some(self.inline[self.len])
    }

    fn into_vec(mut self) -> Vec<usize> {
        // Low indices are the bottom of the LIFO stack. Append the spill
        // after the inline prefix, preserving precisely the pending order.
        // Reuse already-owned spare storage, but do not force a realloc:
        // measurements favor the existing allocation path when capacity is
        // insufficient. usize entries have no reference/drop side effects.
        let old_len = self.spill.len();
        if self.spill.capacity() - old_len >= self.len {
            self.spill.resize(old_len + self.len, 0);
            self.spill.copy_within(..old_len, self.len);
            self.spill[..self.len].copy_from_slice(&self.inline[..self.len]);
            return self.spill;
        }
        let mut pending = Vec::with_capacity(self.len + self.spill.len());
        pending.extend_from_slice(&self.inline[..self.len]);
        pending.extend(self.spill);
        pending
    }
}

struct Visited {
    inline: [usize; 8],
    len: usize,
    spill: Option<HashSet<usize>>,
}

impl Visited {
    fn new(root: usize) -> Self {
        let mut inline = [0; 8];
        inline[0] = root;
        Self { inline, len: 1, spill: None }
    }

    #[inline]
    fn insert(&mut self, id: usize) -> bool {
        if let Some(seen) = &mut self.spill { return seen.insert(id); }
        if self.inline[..self.len].contains(&id) { return false; }
        if self.len < self.inline.len() {
            self.inline[self.len] = id;
            self.len += 1;
            return true;
        }
        let mut seen = HashSet::with_capacity(16);
        seen.extend(self.inline);
        let fresh = seen.insert(id);
        self.spill = Some(seen);
        fresh
    }
}

// Keep worklist/visited storage out of the common parent-only call frame.
#[inline(never)]
pub(super) fn graph_is_subtype(reg: &TypeRegistry, start: usize, target: usize) -> bool {
    let root = &reg.types[start];
    if root.implements.len() > 8 {
        let mut pending = root.implements.clone();
        if let Some(parent) = root.parent { pending.push(parent); }
        return vector_graph_continue(reg, pending, Visited::new(start), target);
    }
    let mut pending = Pending::new();
    pending.extend(&root.implements);
    if let Some(parent) = root.parent { pending.push(parent); }
    let mut seen = Visited::new(start);
    while let Some(id) = pending.pop() {
        if id == target { return true; }
        if !seen.insert(id) { continue; }
        if let Some(ty) = reg.types.get(id) {
            pending.extend(&ty.implements);
            if let Some(parent) = ty.parent { pending.push(parent); }
            if !pending.spill.is_empty() {
                return vector_graph_continue(reg, pending.into_vec(), seen, target);
            }
        }
    }
    false
}

#[inline(never)]
fn vector_graph_continue(reg: &TypeRegistry, mut pending: Vec<usize>, mut seen: Visited,
    target: usize) -> bool
{
    while let Some(id) = pending.pop() {
        if id == target { return true; }
        if !seen.insert(id) { continue; }
        if let Some(ty) = reg.types.get(id) {
            pending.extend(ty.implements.iter().copied());
            if let Some(parent) = ty.parent { pending.push(parent); }
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;
    use super::super::TypeDef;

    #[test]
    fn pending_preserves_lifo_across_spill_and_refill() {
        let mut actual = Pending::new();
        let mut expected = Vec::new();
        for width in [0usize, 1, 7, 8, 9, 32, 1024] {
            let ids: Vec<usize> = (0..width).collect();
            actual.extend(&ids);
            expected.extend_from_slice(&ids);
            actual.push(usize::MAX);
            expected.push(usize::MAX);
            for _ in 0..width / 2 { assert_eq!(actual.pop(), expected.pop()); }
            actual.extend(&[91, 92, 93]);
            expected.extend_from_slice(&[91, 92, 93]);
            while !expected.is_empty() { assert_eq!(actual.pop(), expected.pop()); }
            assert_eq!(actual.pop(), None);
        }
        let mut small = Pending::new();
        small.extend(&[1, 2, 3, 4, 5, 6, 7, 8]);
        assert_eq!(small.spill.capacity(), 0);
        for id in (1..=8).rev() { assert_eq!(small.pop(), Some(id)); }
        assert_eq!(small.pop(), None);
    }

    #[test]
    fn frontier_promotion_preserves_pending_order_and_remaining_prefix() {
        for width in [0usize, 1, 7, 8, 9, 32, 1024] {
            for pops in [0usize, 1, 7, 8, 9] {
                let mut actual = Pending::new();
                let mut expected: Vec<usize> = (0..width).collect();
                actual.extend(&expected);
                for _ in 0..pops { assert_eq!(actual.pop(), expected.pop()); }
                actual.extend(&[usize::MAX, 91, 92]);
                expected.extend([usize::MAX, 91, 92]);
                assert_eq!(actual.into_vec(), expected);
            }
        }
    }

    #[test]
    fn frontier_promotion_reuses_spare_storage_without_changing_order() {
        for prefix_len in 0..=8 {
            let mut pending = Pending::new();
            pending.extend(&(0..prefix_len).collect::<Vec<_>>());
            pending.spill = Vec::with_capacity(32);
            pending.spill.extend(8..20);
            let allocation = pending.spill.as_ptr();
            let expected: Vec<_> = (0..prefix_len).chain(8..20).collect();
            let promoted = pending.into_vec();
            assert_eq!(promoted, expected);
            assert_eq!(promoted.as_ptr(), allocation);
            assert_eq!(promoted.capacity(), 32);
        }
    }

    #[test]
    fn late_wide_frontiers_preserve_cycles_hits_and_misses() {
        for width in [7usize, 8, 9, 32, 1024] {
            let mut reg = TypeRegistry::new();
            for id in 1..=width + 1 { reg.register(TypeDef::new(&format!("late_{id}"))); }
            reg.types[0].implements.push(1);
            reg.types[1].implements.extend(2..=width + 1);
            for id in 2..=width + 1 { reg.types[id].parent = Some(0); }
            for target in (0..=width + 2).chain(std::iter::once(usize::MAX)) {
                assert_eq!(graph_is_subtype(&reg, 0, target), vector_frontier(&reg, 0, target),
                    "width={width} target={target}");
            }
        }
    }

    fn vector_frontier(reg: &TypeRegistry, start: usize, target: usize) -> bool {
        let root = &reg.types[start];
        let mut pending = root.implements.clone();
        if let Some(parent) = root.parent { pending.push(parent); }
        let mut seen = Visited::new(start);
        while let Some(id) = pending.pop() {
            if id == target { return true; }
            if !seen.insert(id) { continue; }
            if let Some(ty) = reg.types.get(id) {
                pending.extend(ty.implements.iter().copied());
                if let Some(parent) = ty.parent { pending.push(parent); }
            }
        }
        false
    }

    #[test]
    #[ignore = "paired graph frontier allocation benchmark"]
    fn pending_frontier_benchmark() {
        use std::{hint::black_box, time::Instant};
        for width in [1usize, 2, 4, 8, 9, 32, 1024] {
            for depth in [2usize, 8, 32] {
                let mut reg = TypeRegistry::new();
                for id in 1..=width + depth {
                    reg.register(TypeDef::new(&format!("frontier_{id}")));
                }
                reg.types[0].implements.extend(1..=width);
                reg.types[1].parent = Some(width + 1);
                for id in width + 1..width + depth { reg.types[id].parent = Some(id + 1); }
                let rounds = (16384 / (width + depth)).clamp(32, 1024);
                let measure = |kernel: fn(&TypeRegistry, usize, usize) -> bool| {
                    let start = Instant::now();
                    for _ in 0..rounds {
                        assert!(!black_box(kernel(black_box(&reg), 0, black_box(usize::MAX))));
                    }
                    start.elapsed().as_secs_f64() * 1e9 / rounds as f64
                };
                let mut old = Vec::new();
                let mut new = Vec::new();
                for trial in 0..15 {
                    if trial % 2 == 0 { old.push(measure(vector_frontier)); new.push(measure(graph_is_subtype)); }
                    else { new.push(measure(graph_is_subtype)); old.push(measure(vector_frontier)); }
                }
                old.sort_by(f64::total_cmp);
                new.sort_by(f64::total_cmp);
                println!("frontier width={width} depth={depth} vector_ns={:.2} inline_ns={:.2} ratio={:.2}",
                    old[7], new[7], old[7] / new[7]);
            }
        }
    }

    #[test]
    fn visited_spills_only_after_eight_distinct_ids_and_preserves_duplicates() {
        let mut actual = Visited::new(usize::MAX);
        let mut expected = HashSet::from([usize::MAX]);
        for id in 0..7 {
            assert_eq!(actual.insert(id), expected.insert(id));
            assert!(actual.spill.is_none());
            assert!(!actual.insert(id));
        }
        assert!(!actual.insert(usize::MAX));
        assert!(actual.spill.is_none());
        for id in 0..32 {
            assert_eq!(actual.insert(id), expected.insert(id));
            assert!(!actual.insert(id));
        }
        assert!(actual.spill.is_some());
        for id in 0..32 { assert!(!actual.insert(id)); }
        assert!(!actual.insert(usize::MAX));
    }
}
