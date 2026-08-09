//! The scheduling SLOT — mechanism in the VM, policy installed by a host.
//!
//! There is no WASM event loop: core wasm is synchronous, and what WASM
//! defines is the suspension MECHANISM (JSPI, stack-switching) plus WASI's
//! time and readiness. Which callback runs next — jobs before tasks, drain to
//! empty or one per turn — is host-spec territory (ECMA-262 §9.5,
//! HTML's processing model), so the loop that decides it does not belong in
//! this crate. The host implements this trait and installs it at plugin
//! registration, exactly like host functions; the VM keeps only a fallback
//! that preserves bare-VM behaviour for its own tests.
//!
//! `turn(&mut VM)` is the inversion: the HOST drives, the VM provides the
//! mechanism (`run_scheduled_callback`, `resume_scheduled_fiber`, the work
//! queues). State lives on the VM, so implementations are stateless structs.

use crate::error::VMError;
use crate::vm::VM;

pub trait Scheduler: Send + Sync {
    /// One turn of host-scheduled work — the job queue under the module's
    /// job queue drained to empty, then at most one deferred task.
    /// Returns `true` if any work ran.
    fn turn(&self, vm: &mut VM) -> Result<bool, VMError>;
    /// Anything pending (jobs, timers, suspended fibers awaiting settlement)?
    fn has_pending(&self, vm: &VM) -> bool;
    /// Block until the nearest deadline/readiness (`wasi:io/poll` shape).
    fn wait(&self, vm: &VM);
}

/// A host-owned source of time-deferred work, `wasi:io/poll` shaped: the VM
/// (or the installed scheduler) can ask whether anything is registered, when
/// to wake, and pop one DUE callback per turn. The STORAGE lives with the
/// host that owns the concept — HTML's timer wheel is `platforms/web`'s, not
/// this crate's — and is registered here at plugin init, exactly like host
/// functions. The VM never sees a fire time being set or a timer id being
/// cancelled; it only polls readiness.
pub trait DeferredSource: Send + Sync {
    /// Any entries registered, due or not?
    fn has_pending(&self) -> bool;
    /// Earliest wake deadline (monotonic ms, `event_loop::monotonic_now_ms`
    /// clock), if any entry is pending.
    fn earliest_deadline_ms(&self) -> Option<f64>;
    /// Pop ONE due entry's callback — first-registered-due-first, matching
    /// the drain's one-task-per-turn contract. `None` if nothing is due yet.
    fn pop_due(&self) -> Option<crate::value::Value>;
    /// Visit every queued callback. Fiber capture needs this: a callback's
    /// open upvalues must be closed before the stack it indexes is saved.
    fn for_each_callback(&self, f: &mut dyn FnMut(&crate::value::Value));

    /// Drop every queued entry — called by [`crate::VM::reset_to`].
    ///
    /// A queued callback is a `Value` belonging to the program that queued it.
    /// A source that keeps one across a reset hands the NEXT program a closure
    /// over code that no longer exists: `reset_to` truncates the chunks a
    /// callback indexes, and those indices are then reused by whatever runs
    /// next. Draining it would run the new tenant's bytes through the old
    /// tenant's closure.
    ///
    /// Default no-op only for sources that hold no per-program state; any
    /// source that queues `Value`s MUST implement it.
    fn clear_pending(&self) {}
}
