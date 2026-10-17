//! CM3 runtime handle table per `CanonicalABI.md §HandleTable`.
//! Maps i32 indices to typed component resources for a single component instance.
//! Distinct from `TypeRegistry` (compile-time type metadata).

use crate::cm_task::CMTask;
use crate::value::Value;
use std::collections::HashMap;

/// Where one end of a `stream`/`future` is in the copy lifecycle —
/// `CanonicalABI.md §stream_copy`'s `CopyState`.
///
/// This is per-END state, not per-stream: a readable and a writable end of the
/// same stream advance independently. `stream_copy` traps unless the end it is
/// handed is `Idle`, which is what forbids a second concurrent read on one end
/// (*"In the future, the 'trap if not IDLE' condition could be relaxed to allow
/// multiple pipelined reads or writes"*). `Done` is entered only on `Dropped` —
/// once an end learns the other side is gone, anything but `drop-*` traps.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CopyState {
    #[default]
    Idle,
    Copying,
    Done,
}

/// One end of a `stream`/`future`, with the lifecycle state `stream_copy`
/// inspects. `id` indexes the EventLoop registry.
///
/// `in_waitable_set` is the second trap condition: a SYNCHRONOUS
/// `stream.{read,write}` on an end that is already being awaited
/// asynchronously traps, because the event would have two claimants.
#[derive(Debug, Clone, Copy, Default)]
pub struct StreamEnd {
    pub id: u64,
    pub type_id: Option<u32>,
    pub state: CopyState,
    pub in_waitable_set: bool,
}

impl StreamEnd {
    pub fn new(id: u64) -> Self {
        StreamEnd {
            id,
            type_id: None,
            state: CopyState::Idle,
            in_waitable_set: false,
        }
    }

    pub fn typed(id: u64, type_id: Option<u32>) -> Self {
        Self { type_id, ..Self::new(id) }
    }
}

/// Typed resource entry stored in the handle table.
#[derive(Debug, Clone)]
pub enum HandleEntry {
    /// Readable end of a stream<T>: the stream_id in the EventLoop registry.
    ReadableStreamEnd(StreamEnd),
    /// Writable end of a stream<T>.
    WritableStreamEnd(StreamEnd),
    /// Readable end of a future<T>: the future_id in the EventLoop registry.
    ReadableFutureEnd(StreamEnd),
    /// Writable end of a future<T>.
    WritableFutureEnd(StreamEnd),
    /// A pending async subtask (the future_id that will resolve when the subtask completes).
    Subtask { future_id: u64, state: SubtaskState },
    /// An owned component resource (type_id + value).
    OwnedResource {
        type_id: u32,
        value: Value,
        num_lends: u32,
    },
    /// 📝 `class ErrorContext: debug_message: String` — `CanonicalABI.md:5144`.
    ///
    /// The spec's handle table is
    /// `Table[ResourceHandle | Waitable | WaitableSet | ErrorContext]`, so this
    /// belongs here and not in a table of its own.
    ErrorContext { debug_message: String },
    /// A borrowed resource lent from another task (scope_task is the lending CMTask id).
    BorrowedResource {
        type_id: u32,
        value: Value,
        scope_task: u32,
        num_lends: u32,
    },
}

impl HandleEntry {
    /// What this entry IS, for a trap that has to say why a handle was wrong.
    ///
    /// `trap_if(not isinstance(h, X))` appears all over `CanonicalABI.md`, and
    /// a message that only says "wrong kind of handle" sends the reader to
    /// read the table by hand. Naming both sides is the difference between a
    /// one-line diagnosis and a debugging session.
    pub fn kind_name(&self) -> &'static str {
        match self {
            HandleEntry::ReadableStreamEnd(_) => "readable stream end",
            HandleEntry::WritableStreamEnd(_) => "writable stream end",
            HandleEntry::ReadableFutureEnd(_) => "readable future end",
            HandleEntry::WritableFutureEnd(_) => "writable future end",
            HandleEntry::Subtask { .. } => "subtask",
            HandleEntry::OwnedResource { .. } => "owned resource",
            HandleEntry::BorrowedResource { .. } => "borrowed resource",
            HandleEntry::ErrorContext { .. } => "error-context",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SubtaskState {
    Starting,
    Started,
    Returned,
    Cancelled,
}

/// Per-instance handle table.
#[derive(Debug, Clone)]
pub struct HandleTable {
    entries: HashMap<u32, HandleEntry>,
    ends_by_id: HashMap<(u64, bool), Vec<u32>>,
    next_id: u32,
    free: Vec<u32>,
}

impl Default for HandleTable {
    fn default() -> Self {
        Self::new()
    }
}

impl HandleTable {
    pub const MAX_INDEX: u32 = (1 << 28) - 1;

    pub fn new() -> Self {
        HandleTable {
            entries: HashMap::new(),
            ends_by_id: HashMap::new(),
            next_id: 1,
            free: Vec::new(),
        }
    }

    fn end_key(entry: &HandleEntry) -> Option<(u64, bool)> {
        match entry {
            HandleEntry::ReadableStreamEnd(end) | HandleEntry::ReadableFutureEnd(end) => {
                Some((end.id, true))
            }
            HandleEntry::WritableStreamEnd(end) | HandleEntry::WritableFutureEnd(end) => {
                Some((end.id, false))
            }
            _ => None,
        }
    }

    pub fn can_allocate(&self, count: usize) -> bool {
        let fresh = if self.next_id > Self::MAX_INDEX {
            0
        } else {
            (Self::MAX_INDEX - self.next_id + 1) as usize
        };
        self.free.len().saturating_add(fresh) >= count
    }

    /// Allocate a handle, reusing the most recently removed index first.
    pub fn insert(&mut self, entry: HandleEntry) -> Result<u32, &'static str> {
        let id = if let Some(id) = self.free.pop() {
            id
        } else {
            if self.next_id > Self::MAX_INDEX {
                return Err("trap: component handle table exhausted");
            }
            let id = self.next_id;
            self.next_id += 1;
            id
        };
        if let Some(key) = Self::end_key(&entry) {
            self.ends_by_id.entry(key).or_default().push(id);
        }
        self.entries.insert(id, entry);
        Ok(id)
    }

    pub fn get(&self, id: u32) -> Option<&HandleEntry> {
        self.entries.get(&id)
    }

    pub fn get_mut(&mut self, id: u32) -> Option<&mut HandleEntry> {
        self.entries.get_mut(&id)
    }

    /// Keep the guest-visible handle while relinking its underlying end.
    pub fn retarget_end(&mut self, handle: u32, new_id: u64) -> bool {
        let Some(entry) = self.entries.get_mut(&handle) else { return false };
        let Some(old_key) = Self::end_key(entry) else { return false };
        match entry {
            HandleEntry::ReadableStreamEnd(end)
            | HandleEntry::WritableStreamEnd(end)
            | HandleEntry::ReadableFutureEnd(end)
            | HandleEntry::WritableFutureEnd(end) => end.id = new_id,
            _ => unreachable!(),
        }
        let new_key = Self::end_key(entry).expect("end remains an end");
        if old_key != new_key {
            if let Some(handles) = self.ends_by_id.get_mut(&old_key) {
                handles.retain(|id| *id != handle);
                if handles.is_empty() { self.ends_by_id.remove(&old_key); }
            }
            self.ends_by_id.entry(new_key).or_default().push(handle);
        }
        true
    }

    pub fn writable_end_handles(&self, id: u64, stream: bool) -> Vec<u32> {
        self.ends_by_id.get(&(id, false)).into_iter().flatten()
            .copied()
            .filter(|handle| matches!(
                (stream, self.entries.get(handle)),
                (true, Some(HandleEntry::WritableStreamEnd(_)))
                    | (false, Some(HandleEntry::WritableFutureEnd(_)))
            ))
            .collect()
    }

    pub(crate) fn has_future_ends(&self, id: u64) -> bool {
        [true, false].into_iter().any(|readable| {
            self.ends_by_id.get(&(id, readable)).into_iter().flatten().any(|handle| {
                matches!(self.entries.get(handle),
                    Some(HandleEntry::ReadableFutureEnd(_) | HandleEntry::WritableFutureEnd(_)))
            })
        })
    }

    /// Remove and return an existing handle entry.
    pub fn remove(&mut self, id: u32) -> Option<HandleEntry> {
        let entry = self.entries.remove(&id)?;
        if let Some(key) = Self::end_key(&entry) {
            if let Some(ids) = self.ends_by_id.get_mut(&key) {
                ids.retain(|other| *other != id);
                if ids.is_empty() {
                    self.ends_by_id.remove(&key);
                }
            }
        }
        self.free.push(id);
        Some(entry)
    }

    pub fn remove_if(
        &mut self,
        id: u32,
        matches: impl FnOnce(&HandleEntry) -> bool,
    ) -> Option<HandleEntry> {
        if !matches(self.entries.get(&id)?) {
            return None;
        }
        self.remove(id)
    }

    pub fn lower_owned_resource(
        &mut self,
        type_id: u32,
        value: Value,
    ) -> Result<u32, &'static str> {
        self.insert(HandleEntry::OwnedResource {
            type_id,
            value,
            num_lends: 0,
        })
    }

    pub fn lift_owned_resource(&mut self, id: u32, type_id: u32) -> Result<Value, &'static str> {
        // Canonical lift removes first; even a mismatched handle is consumed on trap.
        match self.remove(id) {
            Some(HandleEntry::OwnedResource {
                type_id: actual,
                num_lends: 0,
                value,
            }) if actual == type_id => Ok(value),
            Some(HandleEntry::OwnedResource { num_lends, .. }) if num_lends != 0 => {
                Err("resource has outstanding lends")
            }
            Some(HandleEntry::OwnedResource { .. }) => Err("resource type mismatch"),
            _ => Err("handle is not an owned resource"),
        }
    }

    pub fn lift_borrowed_resource(&mut self, id: u32, type_id: u32) -> Result<Value, &'static str> {
        match self.get_mut(id) {
            Some(HandleEntry::OwnedResource {
                type_id: actual,
                value,
                num_lends,
            })
            | Some(HandleEntry::BorrowedResource {
                type_id: actual,
                value,
                num_lends,
                ..
            }) if *actual == type_id => {
                *num_lends = num_lends
                    .checked_add(1)
                    .ok_or("resource lend count overflow")?;
                Ok(value.clone())
            }
            Some(HandleEntry::OwnedResource { .. })
            | Some(HandleEntry::BorrowedResource { .. }) => Err("resource type mismatch"),
            _ => Err("handle is not a resource"),
        }
    }

    pub fn lower_borrowed_resource(
        &mut self,
        type_id: u32,
        value: Value,
        scope: &mut CMTask,
        implementing_instance: bool,
    ) -> Result<u32, &'static str> {
        let Value::I32(rep) = value else {
            return Err("resource representation is not i32");
        };
        if implementing_instance {
            return Ok(rep as u32);
        }
        let count = scope
            .num_borrows
            .checked_add(1)
            .ok_or("borrow count overflow")?;
        let id = self.insert(HandleEntry::BorrowedResource {
            type_id,
            value: Value::I32(rep),
            scope_task: scope.id,
            num_lends: 0,
        })?;
        scope.num_borrows = count;
        Ok(id)
    }

    pub fn release_resource_lend(&mut self, id: u32) -> Result<(), &'static str> {
        match self.get_mut(id) {
            Some(HandleEntry::OwnedResource { num_lends, .. })
            | Some(HandleEntry::BorrowedResource { num_lends, .. })
                if *num_lends > 0 =>
            {
                *num_lends -= 1;
                Ok(())
            }
            _ => Err("resource has no outstanding lend"),
        }
    }

    /// Release a `Copying` end that was waiting on the stream/future `id`,
    /// because the other side has just made progress.
    ///
    /// This is the observable half of `stream_event` / `set_pending_event`:
    /// the spec parks a copy in `COPYING` and resets it to `IDLE` when the
    /// event is DELIVERED. Without that reset a `BLOCKED` copy is a dead end —
    /// the end can never be read again, because every subsequent copy traps on
    /// "not IDLE", and the guest has no way back short of dropping the handle.
    ///
    /// `readable` selects which side to release: a writer making data available
    /// frees a parked READER, and vice versa.
    pub fn release_copying(&mut self, id: u64, readable: bool) {
        if let Some(handles) = self.ends_by_id.get(&(id, readable)) {
            for handle in handles {
                let end = match (self.entries.get_mut(handle), readable) {
                    (Some(HandleEntry::ReadableStreamEnd(e)), true)
                    | (Some(HandleEntry::ReadableFutureEnd(e)), true)
                    | (Some(HandleEntry::WritableStreamEnd(e)), false)
                    | (Some(HandleEntry::WritableFutureEnd(e)), false) => e,
                    _ => continue,
                };
                if end.id == id && end.state == CopyState::Copying {
                    end.state = CopyState::Idle;
                }
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{CopyState, HandleEntry, HandleTable, StreamEnd};
    use crate::cm_task::CMTask;
    use crate::value::Value;

    fn entry() -> HandleEntry {
        HandleEntry::ErrorContext {
            debug_message: String::new(),
        }
    }

    #[test]
    fn removed_indices_are_reused_lifo_without_reusing_zero() {
        let mut table = HandleTable::default();
        let a = table.insert(entry()).unwrap();
        let b = table.insert(entry()).unwrap();
        let c = table.insert(entry()).unwrap();
        assert_eq!([a, b, c], [1, 2, 3]);
        assert!(table.get(0).is_none());
        assert!(table.remove(b).is_some());
        assert!(table.remove(a).is_some());
        assert!(table.remove(a).is_none());
        assert_eq!(table.insert(entry()).unwrap(), a);
        assert_eq!(table.insert(entry()).unwrap(), b);
        assert_eq!(table.insert(entry()).unwrap(), 4);
    }

    #[test]
    fn index_limit_traps_but_a_freed_slot_remains_reusable() {
        let mut table = HandleTable::new();
        table.next_id = HandleTable::MAX_INDEX;
        let last = table.insert(entry()).unwrap();
        assert_eq!(last, HandleTable::MAX_INDEX);
        assert_eq!(
            table.insert(entry()).unwrap_err(),
            "trap: component handle table exhausted"
        );
        assert!(!table.can_allocate(1));
        assert!(table.remove(last).is_some());
        assert!(table.can_allocate(1));
        assert_eq!(table.insert(entry()).unwrap(), last);
    }

    #[test]
    fn wrong_kind_removal_preserves_entry_and_reuse_order() {
        let mut table = HandleTable::new();
        let id = table.insert(entry()).unwrap();
        assert!(table.remove_if(id, |_| false).is_none());
        assert!(table.get(id).is_some());
        assert_eq!(table.insert(entry()).unwrap(), id + 1);
        assert!(
            table
                .remove_if(id, |entry| matches!(
                    entry,
                    HandleEntry::ErrorContext { .. }
                ))
                .is_some()
        );
        assert_eq!(table.insert(entry()).unwrap(), id);
    }

    #[test]
    fn owned_resource_transfer_consumes_handle_without_destroying_rep() {
        let mut table = HandleTable::new();
        let id = table.lower_owned_resource(7, Value::I32(42)).unwrap();
        assert!(matches!(
            table.lift_owned_resource(id, 7),
            Ok(Value::I32(42))
        ));
        assert!(table.get(id).is_none());
        assert_eq!(table.lower_owned_resource(7, Value::I32(43)).unwrap(), id);
    }

    #[test]
    fn borrowed_resource_blocks_transfer_until_lend_is_released() {
        let mut table = HandleTable::new();
        let id = table.lower_owned_resource(7, Value::I32(42)).unwrap();
        assert_eq!(
            table.lift_borrowed_resource(id, 8).unwrap_err(),
            "resource type mismatch"
        );
        assert!(matches!(
            table.lift_borrowed_resource(id, 7),
            Ok(Value::I32(42))
        ));
        assert!(matches!(
            table.get(id),
            Some(HandleEntry::OwnedResource { num_lends: 1, .. })
        ));
        assert_eq!(table.release_resource_lend(id), Ok(()));
        assert!(matches!(
            table.lift_owned_resource(id, 7),
            Ok(Value::I32(42))
        ));
    }

    #[test]
    fn failed_owned_lift_consumes_the_handle_as_canonical_abi_specifies() {
        let mut table = HandleTable::new();
        let id = table.lower_owned_resource(7, Value::I32(42)).unwrap();
        assert_eq!(
            table.lift_owned_resource(id, 8).unwrap_err(),
            "resource type mismatch"
        );
        assert!(table.get(id).is_none());
    }

    #[test]
    fn borrowed_resource_is_scoped_only_outside_implementing_instance() {
        let mut table = HandleTable::new();
        let mut task = CMTask::new(9);
        let direct = table
            .lower_borrowed_resource(7, Value::I32(42), &mut task, true)
            .unwrap();
        assert_eq!(direct, 42);
        assert_eq!(task.num_borrows, 0);
        assert!(table.get(direct).is_none());

        let handle = table
            .lower_borrowed_resource(7, Value::I32(42), &mut task, false)
            .unwrap();
        assert_eq!(task.num_borrows, 1);
        assert!(matches!(
            table.get(handle),
            Some(HandleEntry::BorrowedResource {
                type_id: 7,
                value: Value::I32(42),
                scope_task: 9,
                num_lends: 0,
            })
        ));
    }

    #[test]
    fn borrow_count_overflow_does_not_allocate_a_handle() {
        let mut table = HandleTable::new();
        let mut task = CMTask::new(9);
        task.num_borrows = u32::MAX;
        assert_eq!(
            table
                .lower_borrowed_resource(7, Value::I32(42), &mut task, false)
                .unwrap_err(),
            "borrow count overflow"
        );
        assert_eq!(task.num_borrows, u32::MAX);
        assert!(table.get(1).is_none());
        assert_eq!(table.lower_owned_resource(7, Value::I32(42)).unwrap(), 1);
    }

    #[test]
    fn release_copying_respects_id_direction_kind_and_reused_slot() {
        let mut table = HandleTable::new();
        let copying = |id| StreamEnd {
            state: CopyState::Copying,
            ..StreamEnd::new(id)
        };
        let old = table
            .insert(HandleEntry::ReadableStreamEnd(copying(10)))
            .unwrap();
        let reader = table
            .insert(HandleEntry::ReadableFutureEnd(copying(10)))
            .unwrap();
        let writer = table
            .insert(HandleEntry::WritableStreamEnd(copying(10)))
            .unwrap();
        let other = table
            .insert(HandleEntry::ReadableStreamEnd(copying(11)))
            .unwrap();
        assert!(table.remove(old).is_some());
        let reused = table
            .insert(HandleEntry::ReadableStreamEnd(copying(12)))
            .unwrap();
        assert_eq!(reused, old);

        table.release_copying(10, true);
        assert!(
            matches!(table.get(reader), Some(HandleEntry::ReadableFutureEnd(e)) if e.state == CopyState::Idle)
        );
        assert!(
            matches!(table.get(writer), Some(HandleEntry::WritableStreamEnd(e)) if e.state == CopyState::Copying)
        );
        assert!(
            matches!(table.get(other), Some(HandleEntry::ReadableStreamEnd(e)) if e.state == CopyState::Copying)
        );
        assert!(
            matches!(table.get(reused), Some(HandleEntry::ReadableStreamEnd(e)) if e.state == CopyState::Copying)
        );
        table.release_copying(10, false);
        assert!(
            matches!(table.get(writer), Some(HandleEntry::WritableStreamEnd(e)) if e.state == CopyState::Idle)
        );
    }

    #[test]
    fn cloned_table_keeps_endpoint_lookup_independent() {
        let mut original = HandleTable::new();
        let handle = original
            .insert(HandleEntry::ReadableStreamEnd(StreamEnd {
                state: CopyState::Copying,
                ..StreamEnd::new(8)
            }))
            .unwrap();
        let mut snapshot = original.clone();
        snapshot.release_copying(8, true);
        assert!(
            matches!(snapshot.get(handle), Some(HandleEntry::ReadableStreamEnd(e)) if e.state == CopyState::Idle)
        );
        assert!(
            matches!(original.get(handle), Some(HandleEntry::ReadableStreamEnd(e)) if e.state == CopyState::Copying)
        );
    }

    #[test]
    #[ignore = "microbenchmark; run explicitly on a quiet machine"]
    fn bench_release_copying_among_many_ends() {
        let mut table = HandleTable::new();
        let mut target = 0;
        for id in 1..=10_000 {
            let handle = table
                .insert(HandleEntry::ReadableStreamEnd(StreamEnd {
                    state: CopyState::Copying,
                    ..StreamEnd::new(id)
                }))
                .unwrap();
            if id == 5_000 {
                target = handle;
            }
        }
        let started = std::time::Instant::now();
        for _ in 0..10_000 {
            table.release_copying(5_000, true);
            if let Some(HandleEntry::ReadableStreamEnd(end)) = table.get_mut(target) {
                end.state = CopyState::Copying;
            }
        }
        eprintln!("10,000 wakes among 10,000 ends: {:?}", started.elapsed());
    }
}
