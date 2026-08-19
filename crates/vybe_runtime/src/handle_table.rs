//! CM3 runtime handle table per `CanonicalABI.md §HandleTable`.
//! Maps i32 indices to typed component resources for a single component instance.
//! Distinct from `TypeRegistry` (compile-time type metadata).

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
    pub state: CopyState,
    pub in_waitable_set: bool,
}

impl StreamEnd {
    pub fn new(id: u64) -> Self {
        StreamEnd {
            id,
            state: CopyState::Idle,
            in_waitable_set: false,
        }
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
    OwnedResource { type_id: u32, value: Value },
    /// A borrowed resource lent from another task (scope_task is the lending CMTask id).
    BorrowedResource {
        type_id: u32,
        value: Value,
        scope_task: u32,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SubtaskState {
    Starting,
    Started,
    Returned,
    Cancelled,
}

/// Per-instance handle table.
#[derive(Debug, Clone, Default)]
pub struct HandleTable {
    entries: HashMap<u32, HandleEntry>,
    next_id: u32,
}

impl HandleTable {
    pub fn new() -> Self {
        HandleTable {
            entries: HashMap::new(),
            next_id: 1,
        }
    }

    /// Allocate a new handle and return its i32 index.
    pub fn insert(&mut self, entry: HandleEntry) -> u32 {
        let id = self.next_id;
        self.next_id += 1;
        self.entries.insert(id, entry);
        id
    }

    pub fn get(&self, id: u32) -> Option<&HandleEntry> {
        self.entries.get(&id)
    }

    pub fn get_mut(&mut self, id: u32) -> Option<&mut HandleEntry> {
        self.entries.get_mut(&id)
    }

    /// Remove and return a handle entry; trap if not found.
    pub fn remove(&mut self, id: u32) -> Option<HandleEntry> {
        self.entries.remove(&id)
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
        for entry in self.entries.values_mut() {
            let end = match (entry, readable) {
                (HandleEntry::ReadableStreamEnd(e), true)
                | (HandleEntry::ReadableFutureEnd(e), true)
                | (HandleEntry::WritableStreamEnd(e), false)
                | (HandleEntry::WritableFutureEnd(e), false) => e,
                _ => continue,
            };
            if end.id == id && end.state == CopyState::Copying {
                end.state = CopyState::Idle;
            }
        }
    }
}
