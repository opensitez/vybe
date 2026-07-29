//! CM3 runtime handle table per `CanonicalABI.md §HandleTable`.
//! Maps i32 indices to typed component resources for a single component instance.
//! Distinct from `TypeRegistry` (compile-time type metadata).

use crate::value::Value;
use std::collections::HashMap;

/// Typed resource entry stored in the handle table.
#[derive(Debug, Clone)]
pub enum HandleEntry {
    /// Readable end of a stream<T>: the stream_id in the EventLoop registry.
    ReadableStreamEnd(u64),
    /// Writable end of a stream<T>.
    WritableStreamEnd(u64),
    /// Readable end of a future<T>: the future_id in the EventLoop registry.
    ReadableFutureEnd(u64),
    /// Writable end of a future<T>.
    WritableFutureEnd(u64),
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
}
