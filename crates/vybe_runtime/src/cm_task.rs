//! CM3 async task lifecycle per `CanonicalABI.md §Task`.
//! Named `CMTask` to avoid collision with `event_loop::Task`.

use crate::value::Value;

/// Lifecycle state of a CM3 async task — `CanonicalABI.md` `class Task`:
///
/// ```python
/// class State(Enum):
///   INITIAL = 1
///   STARTED = 2
///   PENDING_CANCEL = 3
///   CANCEL_DELIVERED = 4
///   RESOLVED = 5
/// ```
///
/// All five, with the spec's names. The two cancellation states are not
/// decoration: `Task.cancel` is `trap_if(self.state != CANCEL_DELIVERED)`, so
/// without them `task.cancel` cannot enforce its precondition and a cancel that
/// was never delivered to core wasm looks exactly like one that was.
///
/// `RESOLVED` is the single terminal state for BOTH paths — `return_` resolves
/// with a value, `cancel` resolves with none — which is why the discriminator
/// is [`CMTask::result`] being `Some`/`None`, not a separate phase.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TaskPhase {
    /// Created but not yet entered the function body.
    Initial,
    /// The async function body is executing.
    Started,
    /// Cancellation was requested but has not yet been delivered to core wasm.
    PendingCancel,
    /// Cancellation has been delivered; `task.cancel` is now permitted.
    CancelDelivered,
    /// `task.return` or `task.cancel` has resolved the task. The function body
    /// may still run to drain streams.
    Resolved,
}

/// A single CM3 async task.
#[derive(Debug, Clone)]
pub struct CMTask {
    /// Unique task ID (i32 in the handle table of the calling component).
    pub id: u32,
    pub phase: TaskPhase,
    /// The resolved value. `Some` after `task.return`, `None` after
    /// `task.cancel` — the spec's `on_resolve(result)` vs `on_resolve(None)`.
    ///
    /// It exists because `canon_task_return` ends `task.return_(result)` and
    /// `return []`: the result belongs to the TASK, and returning it to the
    /// operand stack instead both loses it and leaves a value the spec says is
    /// not there.
    pub result: Option<Value>,
    /// Outstanding borrowed handles lent to this task; must reach 0 before task.return.
    pub num_borrows: u32,
    /// Outstanding borrows this task has lent to subtasks.
    pub num_lends: u32,
    /// Backpressure COUNTER (CM3 `backpressure.inc`/`dec`): the instance
    /// resists new incoming calls while > 0. The old boolean
    /// `backpressure.set` (canon 0x08) was dropped from the spec.
    pub backpressure: u32,
}

/// Why a resolve attempt was refused. Each maps to one spec `trap_if`, so the
/// caller can name the violated precondition instead of reporting "failed".
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ResolveError {
    /// `trap_if(self.state == Task.State.RESOLVED)` — already resolved.
    AlreadyResolved,
    /// `trap_if(self.num_borrows > 0)` — lent handles not yet returned.
    OutstandingBorrows(u32),
    /// `trap_if(self.state != Task.State.CANCEL_DELIVERED)` — cancellation was
    /// never delivered to core wasm, so the callee cannot know to cancel.
    CancelNotDelivered(TaskPhase),
}

impl std::fmt::Display for ResolveError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ResolveError::AlreadyResolved => write!(f, "task already resolved"),
            ResolveError::OutstandingBorrows(n) => {
                write!(f, "{n} borrowed handle(s) still lent to this task")
            }
            ResolveError::CancelNotDelivered(phase) => write!(
                f,
                "cancellation not delivered to the task (phase {phase:?}, requires CancelDelivered)"
            ),
        }
    }
}

impl CMTask {
    pub fn new(id: u32) -> Self {
        CMTask {
            id,
            phase: TaskPhase::Initial,
            result: None,
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

    /// `Task.return_` — resolve WITH a value.
    ///
    /// ```python
    /// def return_(self, result):
    ///   trap_if(self.state == Task.State.RESOLVED)
    ///   trap_if(self.num_borrows > 0)
    ///   self.on_resolve(result)
    ///   self.state = Task.State.RESOLVED
    /// ```
    ///
    /// The `num_borrows` guard is what lets the CALLER know every borrow it
    /// lent has come back by the time its `OnResolve` runs — dropping it would
    /// let a caller reclaim a handle the callee still holds.
    pub fn return_(&mut self, result: Value) -> Result<(), ResolveError> {
        if self.phase == TaskPhase::Resolved {
            return Err(ResolveError::AlreadyResolved);
        }
        if self.num_borrows > 0 {
            return Err(ResolveError::OutstandingBorrows(self.num_borrows));
        }
        self.result = Some(result);
        self.phase = TaskPhase::Resolved;
        Ok(())
    }

    /// `Task.cancel` — resolve with NO value.
    ///
    /// ```python
    /// def cancel(self):
    ///   trap_if(self.state != Task.State.CANCEL_DELIVERED)
    ///   trap_if(self.num_borrows > 0)
    ///   self.on_resolve(None)
    ///   self.state = Task.State.RESOLVED
    /// ```
    ///
    /// The first guard is the one that needs [`TaskPhase::CancelDelivered`] to
    /// exist: cancelling a task that was never told to cancel is a trap, not a
    /// no-op, because the callee would not have unwound.
    pub fn cancel(&mut self) -> Result<(), ResolveError> {
        if self.phase != TaskPhase::CancelDelivered {
            return Err(ResolveError::CancelNotDelivered(self.phase));
        }
        if self.num_borrows > 0 {
            return Err(ResolveError::OutstandingBorrows(self.num_borrows));
        }
        self.result = None;
        self.phase = TaskPhase::Resolved;
        Ok(())
    }

    /// Request cancellation: Started → PendingCancel. Returns false when the
    /// task is already resolved or already cancelling — neither is a fresh
    /// request.
    pub fn request_cancel(&mut self) -> bool {
        if matches!(self.phase, TaskPhase::Initial | TaskPhase::Started) {
            self.phase = TaskPhase::PendingCancel;
            true
        } else {
            false
        }
    }

    /// Deliver a pending cancellation to core wasm: PendingCancel →
    /// CancelDelivered. This is the transition `Task.cancel` requires.
    pub fn deliver_cancel(&mut self) -> bool {
        if self.phase == TaskPhase::PendingCancel {
            self.phase = TaskPhase::CancelDelivered;
            true
        } else {
            false
        }
    }

    /// `Task.deliver_pending_cancel` — `CanonicalABI.md:941`:
    ///
    /// ```python
    /// def deliver_pending_cancel(self, cancellable) -> bool:
    ///   if cancellable and self.state == Task.State.PENDING_CANCEL:
    ///     self.state = Task.State.CANCEL_DELIVERED
    ///     return True
    ///   return False
    /// ```
    ///
    /// Every blocking thread built-in calls this FIRST, and the `cancellable`
    /// gate is the whole point: a caller that did not opt in is not told, and
    /// the request stays `PendingCancel` so a later `cancellable` call gets it.
    /// Delivering regardless would consume the request against a caller that
    /// cannot act on it — the spec's "will only indicate cancellation once".
    pub fn deliver_pending_cancel(&mut self, cancellable: bool) -> bool {
        cancellable && self.deliver_cancel()
    }

    /// Whether the task has reached its terminal state.
    pub fn resolved(&self) -> bool {
        self.phase == TaskPhase::Resolved
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A non-cancellable caller must NOT consume the pending request. If it
    /// did, the request would be marked delivered against a caller that never
    /// learned of it, and the later `cancellable` call — the one that could
    /// actually propagate — would see nothing pending and run on.
    #[test]
    fn a_non_cancellable_caller_leaves_the_request_pending() {
        let mut t = CMTask::new(1);
        t.start();
        assert!(t.request_cancel());
        assert_eq!(t.phase, TaskPhase::PendingCancel);

        assert!(!t.deliver_pending_cancel(false), "not opted in: not told");
        assert_eq!(
            t.phase,
            TaskPhase::PendingCancel,
            "the request must SURVIVE a non-cancellable caller"
        );

        assert!(t.deliver_pending_cancel(true), "opted in: told");
        assert_eq!(t.phase, TaskPhase::CancelDelivered);

        assert!(
            !t.deliver_pending_cancel(true),
            "indicated ONCE — a second cancellable call sees nothing"
        );
    }

    #[test]
    fn return_resolves_once_and_keeps_the_value() {
        let mut t = CMTask::new(1);
        t.start();
        assert_eq!(t.return_(Value::I32(7)), Ok(()));
        assert_eq!(t.result, Some(Value::I32(7)), "the result belongs to the TASK");
        assert!(t.resolved());
        // `trap_if(self.state == RESOLVED)` — a second return is a trap.
        assert_eq!(t.return_(Value::I32(8)), Err(ResolveError::AlreadyResolved));
    }

    #[test]
    fn return_refuses_while_borrows_are_outstanding() {
        let mut t = CMTask::new(1);
        t.start();
        t.num_borrows = 2;
        assert_eq!(t.return_(Value::I32(1)), Err(ResolveError::OutstandingBorrows(2)));
        assert!(!t.resolved(), "a refused return must not resolve the task");
    }

    #[test]
    fn cancel_requires_delivery_first() {
        // The whole reason CancelDelivered exists: cancelling a task that was
        // never told to cancel is a trap, because the callee never unwound.
        let mut t = CMTask::new(1);
        t.start();
        assert_eq!(
            t.cancel(),
            Err(ResolveError::CancelNotDelivered(TaskPhase::Started))
        );

        assert!(t.request_cancel());
        assert_eq!(
            t.cancel(),
            Err(ResolveError::CancelNotDelivered(TaskPhase::PendingCancel)),
            "requested is not delivered"
        );

        assert!(t.deliver_cancel());
        assert_eq!(t.cancel(), Ok(()));
        assert!(t.resolved());
        assert_eq!(t.result, None, "cancel resolves with NO value");
    }

    #[test]
    fn cancel_and_return_share_one_terminal_state() {
        // Both paths end at RESOLVED; the discriminator is `result`, not a
        // separate phase — so a return after a cancel still traps.
        let mut t = CMTask::new(1);
        t.start();
        t.request_cancel();
        t.deliver_cancel();
        t.cancel().unwrap();
        assert_eq!(t.return_(Value::I32(1)), Err(ResolveError::AlreadyResolved));
    }

    #[test]
    fn a_resolved_task_cannot_be_cancelled_again() {
        let mut t = CMTask::new(1);
        t.start();
        t.return_(Value::I32(1)).unwrap();
        assert!(!t.request_cancel(), "already resolved");
        assert_eq!(
            t.cancel(),
            Err(ResolveError::CancelNotDelivered(TaskPhase::Resolved))
        );
    }
}
