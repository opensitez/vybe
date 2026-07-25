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
    /// Readable end of a future<T> (future_id from EventLoop).
    Future(u64),
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
}

impl WaitableSet {
    pub fn new(id: u32) -> Self {
        WaitableSet {
            id,
            members: Vec::new(),
        }
    }

    /// Join a waitable into this set.
    pub fn join(&mut self, w: Waitable) {
        self.members.push(w);
    }

    /// Poll all members: returns the first ready event (code, handle_u32).
    /// Returns None if no member is ready.
    pub fn poll_ready(&self, el: &EventLoop) -> Option<(EventCode, u64)> {
        for w in &self.members {
            match w {
                Waitable::Stream(id) => {
                    if el.stream_has_item(*id) || el.stream_is_eof(*id) {
                        return Some((EventCode::StreamRead, *id));
                    }
                }
                Waitable::Future(id) => {
                    if let Some(rec) = el.future_states.get(id) {
                        if rec.phase != FuturePhase::Pending {
                            return Some((EventCode::FutureRead, *id));
                        }
                    }
                }
                Waitable::Subtask(id) => {
                    if let Some(rec) = el.future_states.get(id) {
                        if rec.phase != FuturePhase::Pending {
                            return Some((EventCode::Subtask, *id));
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
}

impl WaitableRegistry {
    pub fn new() -> Self {
        WaitableRegistry {
            sets: HashMap::new(),
            next_id: 1,
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
        self.sets.get_mut(&id)
    }

    pub fn remove(&mut self, id: u32) -> Option<WaitableSet> {
        self.sets.remove(&id)
    }
}
