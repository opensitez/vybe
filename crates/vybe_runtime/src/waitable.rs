//! CM3 waitable sets per `Concurrency.md §Waitables and Waitable Sets`.
//! A waitable set aggregates streams, futures, and subtasks so a single
//! `waitable-set.wait` or `waitable-set.poll` can block/check across all.

use crate::event_loop::{EventLoop, FuturePhase};
use std::collections::HashMap;

/// EventCode values written to linear memory on waitable completion.
/// Matches the spec table in `Concurrency.md`.
#[repr(u32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EventCode {
    None = 0,
    Subtask = 1,
    StreamRead = 2,
    StreamWrite = 3,
    FutureRead = 4,
    FutureWrite = 5,
    TaskCancelled = 6,
}

/// A single waitable — stream, future, or subtask.
#[derive(Debug, Clone)]
pub enum Waitable {
    /// Readable end of a stream<T> (stream_id from EventLoop).
    Stream(u64),
    StreamWrite(u64),
    /// Readable end of a future<T> (future_id from EventLoop).
    Future(u64),
    FutureWrite(u64),
    /// A subtask (backed by a future_id from EventLoop).
    Subtask(u64),
}

/// A waitable set — heterogeneous pool of waitables.
/// Created by `waitable-set.new`, populated by `waitable.join`,
/// queried by `waitable-set.wait` / `waitable-set.poll`.
#[derive(Debug, Clone, Default)]
pub struct WaitableSet {
    pub id: u32,
    pub members: Vec<Waitable>,
    pub member_handles: Vec<u32>,
}

impl WaitableSet {
    pub fn new(id: u32) -> Self {
        WaitableSet {
            id,
            members: Vec::new(),
            member_handles: Vec::new(),
        }
    }

    /// Join a waitable into this set.
    pub fn join(&mut self, w: Waitable) {
        self.members.push(w);
        self.member_handles.push(0);
    }

    pub fn join_handle(&mut self, handle: u32, waitable: Waitable) {
        self.members.push(waitable);
        self.member_handles.push(handle);
    }

    /// Poll all members: returns the first ready event (code, handle, payload).
    /// Returns None if no member is ready.
    pub fn poll_ready(&self, el: &EventLoop) -> Option<(EventCode, u32, u32)> {
        for (index, w) in self.members.iter().enumerate() {
            let handle = self.member_handles.get(index).copied().unwrap_or(0);
            match w {
                Waitable::Stream(id) => {
                    if let Some(payload) = el.stream_zero_completion(*id, true) {
                        return Some((EventCode::StreamRead, handle, payload));
                    }
                }
                Waitable::StreamWrite(id) => {
                    if let Some(payload) = el.stream_zero_completion(*id, false) {
                        return Some((EventCode::StreamWrite, handle, payload));
                    }
                }
                Waitable::Future(id) => {
                    if let Some(rec) = el.future_states.get(id) {
                        if rec.phase != FuturePhase::Pending {
                            return Some((EventCode::FutureRead, handle, 0));
                        }
                    }
                }
                Waitable::FutureWrite(id) => {
                    if el.future_states.get(id).is_some_and(|rec| rec.phase == FuturePhase::Pending) {
                        return Some((EventCode::FutureWrite, handle, 0));
                    }
                }
                Waitable::Subtask(id) => {
                    if let Some(rec) = el.future_states.get(id) {
                        if rec.phase != FuturePhase::Pending {
                            return Some((EventCode::Subtask, handle, 0));
                        }
                    }
                }
            }
        }
        None
    }
}

/// VM-level registry of all live waitable sets.
#[derive(Debug, Clone, Default)]
pub struct WaitableRegistry {
    sets: HashMap<u32, WaitableSet>,
    next_id: u32,
    membership_index: Option<HashMap<u32, Vec<u32>>>,
}

impl WaitableRegistry {
    pub fn new() -> Self {
        WaitableRegistry {
            sets: HashMap::new(),
            next_id: 1,
            membership_index: None,
        }
    }

    /// Create a new waitable set and return its u32 handle.
    pub fn create(&mut self) -> u32 {
        let id = self.next_id;
        self.next_id += 1;
        self.sets.insert(id, WaitableSet::new(id));
        id
    }

    pub fn get(&self, id: u32) -> Option<&WaitableSet> {
        self.sets.get(&id)
    }

    pub fn get_mut(&mut self, id: u32) -> Option<&mut WaitableSet> {
        // Public mutation can replace or reorder either membership vector.
        self.membership_index = None;
        self.sets.get_mut(&id)
    }

    pub(crate) fn join_handle(&mut self, id: u32, handle: u32, waitable: Waitable) -> bool {
        let Some(set) = self.sets.get_mut(&id) else { return false; };
        set.join_handle(handle, waitable);
        if handle != 0 {
            if let Some(index) = &mut self.membership_index {
                let sets = index.entry(handle).or_default();
                if !sets.contains(&id) { sets.push(id); }
            }
        }
        true
    }

    pub fn contains_handle(&self, handle: u32) -> bool {
        if handle != 0 {
            if let Some(index) = &self.membership_index { return index.contains_key(&handle); }
        }
        self.sets.values().any(|set| set.member_handles.contains(&handle))
    }

    pub(crate) fn detach_handle(&mut self, handle: u32) {
        if handle == 0 { return; }
        if self.membership_index.is_none() {
            let mut index: HashMap<u32, Vec<u32>> = HashMap::new();
            for (&id, set) in &self.sets {
                for &member in &set.member_handles {
                    if member == 0 { continue; }
                    let sets = index.entry(member).or_default();
                    if sets.last() != Some(&id) { sets.push(id); }
                }
            }
            self.membership_index = Some(index);
        }
        let Some(sets) = self.membership_index.as_mut().unwrap().remove(&handle) else { return; };
        for id in sets {
            let Some(set) = self.sets.get_mut(&id) else { continue; };
            let mut index = 0;
            set.members.retain(|_| {
                let keep = set.member_handles.get(index).copied().unwrap_or(0) != handle;
                index += 1;
                keep
            });
            set.member_handles.retain(|member| *member != handle);
        }
        let index = self.membership_index.as_mut().unwrap();
        if index.is_empty() && index.capacity() > 64 {
            index.shrink_to(64);
        }
    }

    pub fn remove(&mut self, id: u32) -> Option<WaitableSet> {
        let set = self.sets.remove(&id)?;
        if let Some(index) = &mut self.membership_index {
            for handle in &set.member_handles {
                if let std::collections::hash_map::Entry::Occupied(mut entry) = index.entry(*handle) {
                    entry.get_mut().retain(|member| *member != id);
                    if entry.get().is_empty() { entry.remove(); }
                }
            }
        }
        if self.sets.is_empty() { self.membership_index = None; }
        Some(set)
    }
}

#[cfg(test)]
mod tests {
    use super::{EventCode, Waitable, WaitableRegistry, WaitableSet};
    use crate::event_loop::EventLoop;

    #[test]
    fn indexed_membership_observes_mutation_removal_clone_and_reuse() {
        let mut registry = WaitableRegistry::new();
        let first = registry.create();
        let second = registry.create();
        registry.get_mut(first).unwrap().join_handle(17, Waitable::Stream(1));
        registry.get_mut(second).unwrap().join_handle(17, Waitable::StreamWrite(1));
        registry.get_mut(second).unwrap().join_handle(23, Waitable::Future(2));
        registry.detach_handle(999);
        assert!(registry.membership_index.is_some());
        let mut cloned = registry.clone();
        cloned.detach_handle(17);
        assert!(!cloned.contains_handle(17));
        assert!(registry.contains_handle(17));
        assert!(cloned.contains_handle(23));
        registry.get_mut(first).unwrap().member_handles[0] = 42;
        assert!(registry.membership_index.is_none());
        registry.detach_handle(17);
        assert!(!registry.contains_handle(17));
        assert!(registry.contains_handle(42));
        assert_eq!(registry.get(second).unwrap().member_handles, [23]);
        registry.remove(first).unwrap();
        assert!(!registry.contains_handle(42));
        assert!(registry.join_handle(second, 17, Waitable::Stream(3)));
        assert!(registry.contains_handle(17));
        assert!(!registry.join_handle(u32::MAX, 17, Waitable::Stream(4)));
        assert!(registry.contains_handle(17));
        registry.detach_handle(17);
        assert_eq!(registry.get(second).unwrap().member_handles, [23]);
        registry.remove(second).unwrap();
        assert!(registry.membership_index.is_none());
    }

    #[test]
    fn empty_membership_index_releases_peak_capacity_and_remains_usable() {
        let mut registry = WaitableRegistry::new();
        for handle in 1..=1_000 {
            let set = registry.create();
            assert!(registry.join_handle(set, handle, Waitable::Stream(handle as u64)));
        }
        for handle in 1..=1_000 { registry.detach_handle(handle); }
        let index = registry.membership_index.as_ref().unwrap();
        assert!(index.is_empty());
        assert!(index.capacity() <= 128);
        for _ in 0..100 {
            assert!(registry.join_handle(1, 17, Waitable::Stream(1)));
            assert!(registry.contains_handle(17));
            registry.detach_handle(17);
            assert!(!registry.contains_handle(17));
        }
    }

    #[test]
    fn detach_handle_preserves_member_alignment_order_and_unrelated_sets() {
        let mut registry = WaitableRegistry::new();
        let first = registry.create();
        let second = registry.create();
        for set_id in [first, second] {
            let set = registry.get_mut(set_id).unwrap();
            set.join(Waitable::Stream(10));
            set.join_handle(17, Waitable::Stream(11));
            set.join_handle(23, Waitable::StreamWrite(12));
            set.join_handle(17, Waitable::StreamWrite(13));
            set.join_handle(42, Waitable::Future(14));
        }
        registry.detach_handle(0);
        assert_eq!(registry.get(first).unwrap().members.len(), 5);
        registry.detach_handle(17);
        assert!(!registry.contains_handle(17));
        for set_id in [first, second] {
            let set = registry.get(set_id).unwrap();
            assert_eq!(set.member_handles, [0, 23, 42]);
            assert!(matches!(set.members.as_slice(), [Waitable::Stream(10),
                Waitable::StreamWrite(12), Waitable::Future(14)]));
        }
        registry.detach_handle(999);
        assert_eq!(registry.get(first).unwrap().member_handles, [0, 23, 42]);
    }

    #[test]
    fn zero_write_waitable_uses_handle_and_completion_payload() {
        let mut events = EventLoop::new();
        let id = events.create_stream();
        let mut set = WaitableSet::new(1);
        set.join_handle(17, Waitable::StreamWrite(id));
        assert_eq!(set.poll_ready(&events), None);
        assert_eq!(events.stream_zero_write(id), None);
        assert_eq!(set.poll_ready(&events), None);
        assert_eq!(events.stream_zero_read(id), None);
        assert_eq!(set.poll_ready(&events), Some((EventCode::StreamWrite, 17, 0)));
        assert_eq!(events.take_stream_zero_completion(id, false), Some(0));
        assert_eq!(set.poll_ready(&events), None);
    }

    #[test]
    fn mixed_joins_keep_event_handles_aligned() {
        let mut events = EventLoop::new();
        let idle = events.create_stream();
        let ready = events.create_stream();
        let mut set = WaitableSet::new(1);
        set.join(Waitable::Stream(idle));
        set.join_handle(23, Waitable::Stream(ready));
        events.stream_push_bytes(ready, vec![1]);
        assert_eq!(set.poll_ready(&events), None);
        events.complete_guest_stream_read(ready, 16);
        assert_eq!(set.poll_ready(&events), Some((EventCode::StreamRead, 23, 16)));
    }
}
