//! CM3 async task lifecycle per `CanonicalABI.md §Task`.
//! Named `CMTask` to avoid collision with `event_loop::Task`.

/// Lifecycle phase of a CM3 async task.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TaskPhase {
    /// Created but not yet entered the function body.
    Initial,
    /// The async function body is executing.
    Started,
    /// `task.return` has been called; results delivered.
    /// The function body may still run to drain streams.
    Returned,
}

/// A single CM3 async task.
#[derive(Debug, Clone)]
pub struct CMTask {
    /// Unique task ID (i32 in the handle table of the calling component).
    pub id: u32,
    pub phase: TaskPhase,
    /// Outstanding borrowed handles lent to this task; must reach 0 before task.return.
    pub num_borrows: u32,
    /// Outstanding borrows this task has lent to subtasks.
    pub num_lends: u32,
    /// Backpressure COUNTER (CM3 `backpressure.inc`/`dec`): the instance
    /// resists new incoming calls while > 0. The old boolean
    /// `backpressure.set` (canon 0x08) was dropped from the spec.
    pub backpressure: u32,
}

impl CMTask {
    pub fn new(id: u32) -> Self {
        CMTask {
            id,
            phase: TaskPhase::Initial,
            num_borrows: 0,
            num_lends: 0,
            backpressure: 0,
        }
    }

    /// Transition Initial → Started; returns false if already past Initial.
    pub fn start(&mut self) -> bool {
        if self.phase == TaskPhase::Initial {
            self.phase = TaskPhase::Started;
            true
        } else {
            false
        }
    }

    /// Transition Started → Returned; returns false if already Returned (trap).
    pub fn mark_returned(&mut self) -> bool {
        if self.phase == TaskPhase::Returned {
            false
        } else {
            self.phase = TaskPhase::Returned;
            true
        }
    }
}
