//! 🧵 Component Model threads — `CanonicalABI.md` §`Thread`.
//!
//! These are NOT `wasi:threads`. That proposal (`primitives/threading.rs`,
//! `add_import("wasi:threads", "thread-spawn")`) spawns real OS threads over
//! shared memory and is Phase 1 / unversioned. A CM thread is **cooperative,
//! instance-local, and continuation-based**: it never runs concurrently with
//! its siblings, and switching between two of them is a stack switch, not a
//! scheduler preemption.
//!
//! The substrate already exists — `cont.new` / `resume` / `switch` (spec bytes
//! 0xE0..=0xE6) are all dispatched, and a continuation is a
//! `Value::Object(ObjectKind::Continuation)`. What was missing is the model on
//! top: the thread state machine, the instance-local table, and the built-ins
//! that drive them.
//!
//! ```python
//! class Thread:
//!   cont: Optional[Continuation]
//!   ready_func: Optional[Callable[[], bool]]
//!   task: Task
//!   cancellable: bool
//!   index: Optional[int]
//!   storage: tuple[int,int]
//! ```

use crate::value::Value;

/// Length of a thread's `context.get`/`context.set` array — `storage = [0,0]`.
/// Mirrors [`crate::vm::CONTEXT_STORAGE_SLOTS`]; the spec fact is the same one.
pub const THREAD_STORAGE_SLOTS: usize = 2;

/// The readiness predicate of a `waiting` thread.
///
/// ```python
/// def waiting(self): return not self.running() and self.ready_func is not None
/// def ready(self):   return self.waiting() and self.ready_func()
/// ```
///
/// Every 🧵 built-in that starts a thread waiting uses `lambda: True` —
/// `resume_later`, `yield_`, `yield_then_resume`, `yield_then_promote`. The
/// CONDITION-carrying form (`Thread.wait_until(ready_func)`) belongs to the
/// synchronous blocking built-ins (stream/future reads), which park a `Fiber`
/// instead and are a separate mechanism today.
///
/// Modelled as an enum rather than a boxed closure so the gap is legible: a
/// `Condition` variant is the extension point, and its absence is a statement
/// about which built-ins exist, not a limitation hidden behind a `Box<dyn Fn>`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReadyWhen {
    /// `lambda: True` — ready the moment it starts waiting.
    Always,
}

impl ReadyWhen {
    /// Evaluate the predicate — `self.ready_func()`.
    pub fn is_ready(self) -> bool {
        match self {
            ReadyWhen::Always => true,
        }
    }
}

/// The three states a `Thread` can be in. Derived, never stored — the spec
/// defines them as predicates over `cont` and `ready_func`, and storing a
/// fourth field would let the two disagree.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ThreadState {
    /// `cont is None` — actively executing on the stack.
    Running,
    /// Has a continuation and no readiness predicate: waits to be `resume`d by
    /// another thread in the same component instance.
    Suspended,
    /// Has a continuation and a readiness predicate: waits to be resumed
    /// nondeterministically by the host once the predicate holds.
    Waiting,
}

/// A single Component Model thread.
#[derive(Debug)]
pub struct Thread {
    /// The suspended continuation, or `None` while this thread is running.
    /// A `Value::Object(ObjectKind::Continuation)` produced by `cont.new`.
    pub cont: Option<Value>,
    /// `None` ⇒ suspended; `Some` ⇒ waiting, and readiness is its predicate.
    pub ready_func: Option<ReadyWhen>,
    /// The task this thread executes on behalf of (`CMTask::id`).
    pub task_id: u32,
    /// Whether the CURRENT blocking operation opted in to cancellation.
    /// Set by each `block`/`switch_to` and read when a cancel is delivered.
    pub cancellable: bool,
    /// Index in the instance's `threads` table, once registered.
    pub index: Option<u32>,
    /// `storage: tuple[int,int]` — thread-local storage for
    /// `canon context.get` / `context.set`, zero-initialized.
    pub storage: [Value; THREAD_STORAGE_SLOTS],
}

impl Thread {
    /// `Thread.__init__` — a new thread starts **suspended**, holding a
    /// continuation it has not yet entered.
    pub fn new(task_id: u32, cont: Value) -> Self {
        let t = Thread {
            cont: Some(cont),
            ready_func: None,
            task_id,
            cancellable: false,
            index: None,
            storage: [Value::I32(0), Value::I32(0)],
        };
        debug_assert_eq!(t.state(), ThreadState::Suspended);
        t
    }

    /// `running()` / `suspended()` / `waiting()`, as one derived answer.
    pub fn state(&self) -> ThreadState {
        match (&self.cont, &self.ready_func) {
            (None, _) => ThreadState::Running,
            (Some(_), None) => ThreadState::Suspended,
            (Some(_), Some(_)) => ThreadState::Waiting,
        }
    }

    pub fn running(&self) -> bool {
        self.state() == ThreadState::Running
    }

    pub fn suspended(&self) -> bool {
        self.state() == ThreadState::Suspended
    }

    pub fn waiting(&self) -> bool {
        self.state() == ThreadState::Waiting
    }

    /// `ready()` — waiting AND its predicate holds. Only a `ready` thread may
    /// be promoted to by `{suspend,yield}_then_promote`.
    pub fn ready(&self) -> bool {
        self.waiting() && self.ready_func.is_some_and(ReadyWhen::is_ready)
    }

    /// `start_waiting_internal` — `assert(not self.waiting() and not
    /// self.ready_func)`. The caller adds it to the store's waiting list.
    pub fn start_waiting_internal(&mut self, ready_func: ReadyWhen) -> Result<(), ThreadError> {
        if self.waiting() || self.ready_func.is_some() {
            return Err(ThreadError::AlreadyWaiting);
        }
        self.ready_func = Some(ready_func);
        Ok(())
    }

    /// `stop_waiting_internal` — `assert(self.waiting() and self.ready_func)`
    /// and `assert(cancelled or self.ready())`. Clearing the predicate on a
    /// thread that is neither cancelled nor ready would resume it before its
    /// condition held.
    pub fn stop_waiting_internal(&mut self, cancelled: bool) -> Result<(), ThreadError> {
        if !self.waiting() {
            return Err(ThreadError::NotWaiting(self.state()));
        }
        if !cancelled && !self.ready() {
            return Err(ThreadError::NotReady);
        }
        self.ready_func = None;
        Ok(())
    }

    /// `resume_later` — `assert(self.suspended())`, then start waiting with a
    /// predicate that is already true, so the thread is immediately `ready`.
    ///
    /// This does NOT switch to the thread; it makes it eligible to be resumed
    /// at some nondeterministic later point chosen by the embedder.
    pub fn resume_later(&mut self) -> Result<(), ThreadError> {
        if !self.suspended() {
            return Err(ThreadError::NotSuspended(self.state()));
        }
        self.start_waiting_internal(ReadyWhen::Always)?;
        debug_assert!(self.ready());
        Ok(())
    }

    /// Take the continuation for a resume, leaving the thread `running`.
    /// `Thread.resume` does `cont = thread.cont; thread.cont = None`.
    pub fn take_cont(&mut self) -> Option<Value> {
        self.cont.take()
    }

    /// Store the continuation a resume handed back, leaving the thread
    /// suspended again.
    pub fn put_cont(&mut self, cont: Value) {
        self.cont = Some(cont);
    }
}

/// Why a thread transition was refused. One variant per spec assertion, so a
/// failure names the rule it broke rather than reporting "invalid state".
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ThreadError {
    /// `assert(self.suspended())` — `resume_later`, `{suspend,yield}_then_resume`.
    NotSuspended(ThreadState),
    /// `assert(self.waiting() and self.ready_func)`.
    NotWaiting(ThreadState),
    /// `assert(not self.waiting() and not self.ready_func)`.
    AlreadyWaiting,
    /// `assert(cancelled or self.ready())`.
    NotReady,
    /// `inst.threads.get(i)` found nothing.
    NoSuchThread(u32),
}

impl std::fmt::Display for ThreadError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ThreadError::NotSuspended(s) => write!(f, "thread is {s:?}, must be Suspended"),
            ThreadError::NotWaiting(s) => write!(f, "thread is {s:?}, must be Waiting"),
            ThreadError::AlreadyWaiting => write!(f, "thread is already waiting"),
            ThreadError::NotReady => write!(f, "thread is waiting but not ready"),
            ThreadError::NoSuchThread(i) => write!(f, "no thread at index {i}"),
        }
    }
}

/// The component instance's `threads` table, addressed by index.
///
/// ⛔ Lives on the VM rather than a `ComponentInstance`, because we have no
/// such structure — the same deviation `backpressure` and `context_slots`
/// carry. Recorded in `cmplan.md`; it means one table is shared where the spec
/// gives each instance its own.
///
/// Indices are never reused while a thread is live. `unregister` leaves a hole
/// rather than compacting, because an index is a value core wasm holds: moving
/// a live thread to a different index would silently redirect a
/// `thread.resume-later` that is already in flight.
#[derive(Debug, Default)]
pub struct ThreadTable {
    slots: Vec<Option<Thread>>,
}

impl ThreadTable {
    pub fn new() -> Self {
        ThreadTable { slots: Vec::new() }
    }

    /// `task.register_thread(new_thread)` — assign the next index and store.
    pub fn register(&mut self, mut thread: Thread) -> u32 {
        let index = match self.slots.iter().position(Option::is_none) {
            Some(i) => i,
            None => {
                self.slots.push(None);
                self.slots.len() - 1
            }
        } as u32;
        thread.index = Some(index);
        self.slots[index as usize] = Some(thread);
        index
    }

    /// `task.unregister_thread` — called when a thread's function returns.
    pub fn unregister(&mut self, index: u32) -> Option<Thread> {
        self.slots.get_mut(index as usize)?.take()
    }

    pub fn get(&self, index: u32) -> Option<&Thread> {
        self.slots.get(index as usize)?.as_ref()
    }

    pub fn get_mut(&mut self, index: u32) -> Option<&mut Thread> {
        self.slots.get_mut(index as usize)?.as_mut()
    }

    /// Number of LIVE threads, not slots — holes left by `unregister` do not
    /// count.
    pub fn len(&self) -> usize {
        self.slots.iter().filter(|s| s.is_some()).count()
    }

    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }

    /// Indices of every thread currently `ready` — the candidates
    /// `Store.tick` may nondeterministically resume.
    pub fn ready_indices(&self) -> Vec<u32> {
        self.slots
            .iter()
            .enumerate()
            .filter_map(|(i, s)| s.as_ref()?.ready().then_some(i as u32))
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn cont() -> Value {
        // The state machine never dereferences the continuation; it only cares
        // whether one is present. A placeholder keeps these tests independent
        // of `cont.new`'s object shape.
        Value::I32(0)
    }

    fn suspended_thread() -> Thread {
        Thread::new(1, cont())
    }

    #[test]
    fn a_new_thread_is_suspended_not_running() {
        // `Thread.__init__` ends `assert(self.suspended())` — a fresh thread
        // holds a continuation it has not entered.
        let t = suspended_thread();
        assert_eq!(t.state(), ThreadState::Suspended);
        assert!(!t.running() && !t.waiting() && !t.ready());
    }

    #[test]
    fn state_is_derived_from_cont_and_ready_func() {
        let mut t = suspended_thread();
        assert_eq!(t.state(), ThreadState::Suspended);

        t.start_waiting_internal(ReadyWhen::Always).unwrap();
        assert_eq!(t.state(), ThreadState::Waiting);

        // `running()` is `cont is None` — taking the continuation IS the
        // transition, which is why the state is derived and not stored.
        t.take_cont();
        assert_eq!(t.state(), ThreadState::Running);
    }

    #[test]
    fn resume_later_makes_a_suspended_thread_ready() {
        let mut t = suspended_thread();
        t.resume_later().unwrap();
        assert!(t.waiting(), "resume_later moves suspended → waiting");
        assert!(t.ready(), "with `lambda: True`, waiting is immediately ready");
    }

    #[test]
    fn resume_later_refuses_a_thread_that_is_not_suspended() {
        // `assert(self.suspended())`. Calling it twice would double-add the
        // thread to the store's waiting list.
        let mut t = suspended_thread();
        t.resume_later().unwrap();
        assert_eq!(
            t.resume_later(),
            Err(ThreadError::NotSuspended(ThreadState::Waiting))
        );
    }

    #[test]
    fn stop_waiting_refuses_a_thread_that_is_neither_ready_nor_cancelled() {
        // `assert(cancelled or self.ready())` — clearing the predicate early
        // would resume a thread before its condition held.
        let mut t = suspended_thread();
        assert_eq!(
            t.stop_waiting_internal(false),
            Err(ThreadError::NotWaiting(ThreadState::Suspended))
        );

        t.resume_later().unwrap();
        // Ready, so stopping is allowed and returns it to suspended.
        t.stop_waiting_internal(false).unwrap();
        assert!(t.suspended());
    }

    #[test]
    fn a_cancelled_thread_may_stop_waiting_even_when_not_ready() {
        let mut t = suspended_thread();
        t.start_waiting_internal(ReadyWhen::Always).unwrap();
        assert_eq!(t.stop_waiting_internal(true), Ok(()));
        assert!(t.suspended());
    }

    #[test]
    fn the_table_assigns_indices_and_stamps_them_on_the_thread() {
        let mut table = ThreadTable::new();
        let a = table.register(suspended_thread());
        let b = table.register(suspended_thread());
        assert_ne!(a, b);
        assert_eq!(table.get(a).unwrap().index, Some(a));
        assert_eq!(table.get(b).unwrap().index, Some(b));
        assert_eq!(table.len(), 2);
    }

    #[test]
    fn unregister_leaves_a_hole_and_does_not_move_live_threads() {
        // An index is a value core wasm holds. Compacting would silently
        // redirect a `thread.resume-later` already in flight.
        let mut table = ThreadTable::new();
        let a = table.register(suspended_thread());
        let b = table.register(suspended_thread());
        table.unregister(a).unwrap();
        assert!(table.get(a).is_none());
        assert_eq!(
            table.get(b).unwrap().index,
            Some(b),
            "b must not shift into a's slot"
        );
        assert_eq!(table.len(), 1, "len counts live threads, not slots");
    }

    #[test]
    fn ready_indices_lists_only_ready_threads() {
        let mut table = ThreadTable::new();
        let parked = table.register(suspended_thread());
        let woken = table.register(suspended_thread());
        assert!(table.ready_indices().is_empty(), "suspended is not ready");

        table.get_mut(woken).unwrap().resume_later().unwrap();
        assert_eq!(table.ready_indices(), vec![woken]);
        assert!(table.get(parked).unwrap().suspended());
    }
}

#[cfg(test)]
mod lifecycle_tests {
    use super::*;

    /// The implicit thread `canon_lift` spawns is RUNNING, not suspended —
    /// it is the thread executing the lifted call. `Thread::new` starts
    /// suspended (correct for `thread.new-indirect`), so `canon lift` takes
    /// the continuation immediately.
    #[test]
    fn the_implicit_thread_is_running_not_suspended() {
        let mut t = Thread::new(1, Value::Undefined);
        assert!(t.suspended(), "Thread::new starts suspended");
        t.take_cont();
        assert!(t.running(), "the implicit thread is the one executing");
        assert!(!t.suspended() && !t.waiting());
    }

    /// `thread.resume-later` traps unless the target is suspended. The
    /// implicit thread is running, so it is not a legal target — which is the
    /// case that would otherwise resume the thread that is already executing.
    #[test]
    fn resume_later_refuses_the_running_implicit_thread() {
        let mut t = Thread::new(1, Value::Undefined);
        t.take_cont();
        assert_eq!(
            t.resume_later(),
            Err(ThreadError::NotSuspended(ThreadState::Running))
        );
    }

    /// Nested lifted calls: a `realloc` is itself a `canon_lift`, so a second
    /// implicit thread is registered while the first is live. Teardown must
    /// restore the outer index, not clear it.
    #[test]
    fn nested_lifted_calls_restore_the_outer_thread_index() {
        let mut table = ThreadTable::new();
        let mut outer = Thread::new(1, Value::Undefined);
        outer.take_cont();
        let outer_idx = table.register(outer);

        let mut inner = Thread::new(2, Value::Undefined);
        inner.take_cont();
        let inner_idx = table.register(inner);
        assert_ne!(outer_idx, inner_idx, "the inner call gets its own slot");

        // Inner call returns: unregister inner, restore outer.
        table.unregister(inner_idx);
        assert!(table.get(inner_idx).is_none());
        assert_eq!(
            table.get(outer_idx).unwrap().index,
            Some(outer_idx),
            "clearing instead of restoring would strand the outer thread"
        );
        assert_eq!(table.len(), 1);
    }

    /// A thread created by `thread.new-indirect` IS a legal `resume-later`
    /// target, and becomes ready — the contrast that shows the guard is about
    /// state, not about which built-in created the thread.
    #[test]
    fn an_explicit_thread_is_a_legal_resume_later_target() {
        let mut table = ThreadTable::new();
        let idx = table.register(Thread::new(1, Value::I32(0)));
        let t = table.get_mut(idx).unwrap();
        assert!(t.suspended());
        t.resume_later().unwrap();
        assert!(t.ready());
        assert_eq!(table.ready_indices(), vec![idx]);
    }
}

#[cfg(test)]
mod handoff_tests {
    use super::*;

    /// The 2x2 the four compound handoffs implement, asserted on the MODEL so
    /// the decision table is pinned independently of the dispatch plumbing:
    ///
    ///   suspend = park me      yield   = leave me runnable
    ///   resume  = switch always (target must be Suspended)
    ///   promote = switch only if target is READY, else fall back
    #[test]
    fn resume_requires_suspended_but_promote_requires_ready() {
        // A freshly created thread is Suspended but NOT ready: a legal
        // `*_then_resume` target, and NOT a legal `*_then_promote` switch.
        let t = Thread::new(1, Value::I32(0));
        assert!(t.suspended(), "resume's precondition holds");
        assert!(!t.ready(), "promote's precondition does NOT");
    }

    #[test]
    fn resume_later_makes_a_thread_a_promote_target_but_not_a_resume_target() {
        // The exact inverse, which is what makes the two built-ins distinct
        // rather than one being a special case of the other.
        let mut t = Thread::new(1, Value::I32(0));
        t.resume_later().unwrap();
        assert!(t.ready(), "promote will switch to it");
        assert!(
            !t.suspended(),
            "and `*_then_resume` must REFUSE it — it is Waiting, not Suspended"
        );
    }

    #[test]
    fn promote_clears_the_predicate_before_switching() {
        // `other.stop_waiting_internal(cancelled = False)` — a promoted thread
        // stops waiting and becomes Suspended again, so it is not left on the
        // ready list while it is the one running.
        let mut t = Thread::new(1, Value::I32(0));
        t.resume_later().unwrap();
        t.stop_waiting_internal(false).unwrap();
        assert!(t.suspended());
        assert!(!t.ready(), "must not still look ready once entered");
    }

    #[test]
    fn yield_leaves_the_yielding_thread_ready_not_parked() {
        // `yield_then_*` does `start_waiting_internal(lambda: True)` on SELF
        // before switching, so the yielding thread can be picked up again.
        // `suspend_then_*` does not, leaving it merely Suspended.
        let mut yielding = Thread::new(1, Value::I32(0));
        yielding
            .start_waiting_internal(ReadyWhen::Always)
            .unwrap();
        assert!(yielding.ready(), "a yielder stays runnable");

        let parked = Thread::new(1, Value::I32(0));
        assert!(parked.suspended() && !parked.ready(), "a suspender does not");
    }

    #[test]
    fn a_thread_cannot_be_left_both_waiting_and_re_yielded() {
        // `start_waiting_internal` asserts `not self.waiting()`. Two yields
        // without an intervening resume would double-add it to the store's
        // waiting list.
        let mut t = Thread::new(1, Value::I32(0));
        t.start_waiting_internal(ReadyWhen::Always).unwrap();
        assert_eq!(
            t.start_waiting_internal(ReadyWhen::Always),
            Err(ThreadError::AlreadyWaiting)
        );
    }
}
