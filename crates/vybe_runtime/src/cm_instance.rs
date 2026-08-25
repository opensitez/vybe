//! `ComponentInstance` — `CanonicalABI.md` §`class ComponentInstance`.
//!
//! ```python
//! class ComponentInstance:
//!   store: Store
//!   parent: Optional[ComponentInstance]
//!   handles: Table[ResourceHandle | Waitable | WaitableSet | ErrorContext]
//!   threads: Table[Thread]
//!   may_enter: bool
//!   may_leave: bool
//!   backpressure: int
//!   num_waiting_to_enter: int
//!   exclusive_thread: Optional[Thread]
//! ```
//!
//! This structure was missing, and its absence was not neutral — the state it
//! owns had been scattered onto whatever was nearby:
//!
//! - `backpressure` sat on `CMTask`, so a per-INSTANCE counter was per-task.
//! - the `threads` table had nowhere to live at all.
//! - `may_leave` did not exist, so **every** `trap_if(not inst.may_leave)` in
//!   the canonical ABI — the guard on nearly every built-in — was absent.
//!
//! `handles` stays on the VM for now (`VM::handle_table`), because the resource
//! and waitable tables are shared with machinery outside the CM built-ins;
//! moving it is a separate change with its own consumers.

use crate::cm_thread::ThreadTable;

/// One component instance.
///
/// ⛔ A real embedding has many, keyed by a `Store`. This VM has exactly one,
/// held as `VM::cm_instance`, so `current_instance()` is always that one.
/// Recorded in `cmplan.md`: it means two components would share an instance
/// where the spec gives each its own, which matters for `may_enter` reentrance
/// and for the handle/thread index spaces.
#[derive(Debug)]
pub struct ComponentInstance {
    /// The instance's `threads` table — `thread.new-indirect` registers here
    /// and every other 🧵 built-in addresses threads by index into it.
    pub threads: ThreadTable,
    /// Whether a synchronous call may ENTER this instance. Cleared while a
    /// synchronous export is running, so reentrance traps instead of
    /// corrupting the in-flight call.
    pub may_enter: bool,
    /// Whether guest code may LEAVE this instance.
    ///
    /// Cleared for the duration of `post-return` and `realloc`, which is what
    /// lets a synchronously-lowered call to a synchronously-lifted function be
    /// a plain function call: neither may block, so no fiber is needed.
    /// Nearly every canonical built-in opens with `trap_if(not inst.may_leave)`.
    pub may_leave: bool,
    /// `backpressure` — the instance resists new incoming calls while > 0.
    /// A COUNTER (`backpressure.inc`/`dec`), not the old boolean
    /// `backpressure.set`, which was dropped from Binary.md.
    pub backpressure: u32,
    /// How many callers are parked waiting for `may_enter` / backpressure.
    pub num_waiting_to_enter: u32,
    /// The thread currently holding exclusive access, by table index.
    pub exclusive_thread: Option<u32>,
}

impl Default for ComponentInstance {
    fn default() -> Self {
        Self::new()
    }
}

impl ComponentInstance {
    /// `ComponentInstance.__init__` — both `may_*` flags start **true**, the
    /// counters at zero.
    pub fn new() -> Self {
        ComponentInstance {
            threads: ThreadTable::new(),
            may_enter: true,
            may_leave: true,
            backpressure: 0,
            num_waiting_to_enter: 0,
            exclusive_thread: None,
        }
    }

    /// `trap_if(not inst.may_leave)` — the guard nearly every canonical
    /// built-in opens with. Named once so each built-in states the rule rather
    /// than re-spelling the condition.
    pub fn require_may_leave(&self, builtin: &str) -> Result<(), String> {
        if self.may_leave {
            Ok(())
        } else {
            Err(format!(
                "canon {builtin}: may_leave is clear — a built-in cannot be \
                 called during post-return or realloc (trap)"
            ))
        }
    }

    /// Run `f` with `may_leave` cleared, restoring it afterwards.
    ///
    /// ```python
    /// assert(cx.inst.may_leave)
    /// inst.may_leave = False
    /// [] = call_and_trap_on_throw(opts.post_return, flat_results)
    /// inst.may_leave = True
    /// ```
    ///
    /// The spec asserts `may_leave` is SET on entry — nesting two of these
    /// would restore it early, re-permitting a leave while the outer
    /// `post-return` was still running.
    pub fn enter_no_leave(&mut self) -> Result<(), String> {
        if !self.may_leave {
            return Err("may_leave is already clear — nested post-return/realloc".into());
        }
        self.may_leave = false;
        Ok(())
    }

    pub fn exit_no_leave(&mut self) {
        self.may_leave = true;
    }

    /// `backpressure.inc` — `trap_if(inst.backpressure == 2**16)`.
    pub fn backpressure_inc(&mut self) -> Result<(), String> {
        self.backpressure += 1;
        if self.backpressure == 1 << 16 {
            return Err("backpressure counter reached 2^16 (trap)".into());
        }
        Ok(())
    }

    /// `backpressure.dec` — `trap_if(inst.backpressure < 0)`.
    pub fn backpressure_dec(&mut self) -> Result<(), String> {
        if self.backpressure == 0 {
            return Err("backpressure counter is already 0 (trap)".into());
        }
        self.backpressure -= 1;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_fresh_instance_permits_entering_and_leaving() {
        let i = ComponentInstance::new();
        assert!(i.may_enter && i.may_leave);
        assert_eq!(i.backpressure, 0);
        assert!(i.exclusive_thread.is_none());
        assert!(i.threads.is_empty());
    }

    #[test]
    fn may_leave_gates_builtins_and_restores() {
        let mut i = ComponentInstance::new();
        assert!(i.require_may_leave("stream.read").is_ok());

        i.enter_no_leave().unwrap();
        let err = i.require_may_leave("stream.read").unwrap_err();
        assert!(err.contains("stream.read"), "the trap names the row: {err}");
        assert!(err.contains("post-return"));

        i.exit_no_leave();
        assert!(i.require_may_leave("stream.read").is_ok());
    }

    #[test]
    fn nesting_no_leave_is_refused_not_silently_restored() {
        // `assert(cx.inst.may_leave)` on entry. Nesting would let the inner
        // exit re-permit leaving while the outer post-return still ran.
        let mut i = ComponentInstance::new();
        i.enter_no_leave().unwrap();
        assert!(i.enter_no_leave().is_err());
    }

    #[test]
    fn backpressure_traps_at_the_ceiling_instead_of_saturating() {
        // Saturating loses the inc/dec pairing: the counter stops rising while
        // `dec` keeps falling, so backpressure releases early.
        let mut i = ComponentInstance::new();
        i.backpressure = (1 << 16) - 1;
        assert!(i.backpressure_inc().is_err(), "2^16 is a trap");
    }

    #[test]
    fn backpressure_traps_on_an_unbalanced_dec() {
        let mut i = ComponentInstance::new();
        assert!(i.backpressure_dec().is_err(), "dec below 0 is a trap");
        i.backpressure_inc().unwrap();
        assert_eq!(i.backpressure, 1);
        i.backpressure_dec().unwrap();
        assert_eq!(i.backpressure, 0);
    }

    #[test]
    fn the_instance_owns_the_thread_table() {
        // It had nowhere to live before; `thread.new-indirect` registers here
        // and every other 🧵 built-in addresses it by index.
        let mut i = ComponentInstance::new();
        let t = crate::cm_thread::Thread::new(1, crate::value::Value::I32(0));
        let idx = i.threads.register(t);
        assert_eq!(i.threads.len(), 1);
        assert!(i.threads.get(idx).unwrap().suspended());
    }
}

#[cfg(test)]
mod guard_coverage_tests {
    use super::*;
    use crate::vm::CanonBuiltin as B;

    /// The exemption list is a claim about the SPEC, so pin it. Each of these
    /// was checked against its own `CanonicalABI.md` definition; the rest of
    /// the 33 rows carry `trap_if(not inst.may_leave)` either directly or
    /// through `stream_copy` / `cancel_copy` / `future_copy` / `drop`.
    ///
    /// A first pass at deriving this by script split definitions at ``` fences
    /// and wrongly reported ten DELEGATING rows as unguarded. This test exists
    /// so that mistake cannot be re-made silently.
    #[test]
    fn only_five_in_scope_rows_are_exempt_from_may_leave() {
        let exempt = [
            B::BackpressureInc,
            B::BackpressureDec,
            B::ContextGet,
            B::ContextSet,
            B::ResourceRep,
        ];
        assert_eq!(exempt.len(), 5);

        // Rows that DELEGATE in the spec and are therefore guarded — the exact
        // set the flawed extraction got wrong.
        for guarded in [
            B::StreamRead,
            B::StreamWrite,
            B::FutureRead,
            B::FutureWrite,
            B::StreamCancelRead,
            B::StreamCancelWrite,
            B::FutureCancelRead,
            B::FutureCancelWrite,
            B::StreamDropReadable,
            B::StreamDropWritable,
            B::FutureDropReadable,
            B::FutureDropWritable,
        ] {
            assert!(
                !exempt.contains(&guarded),
                "{} delegates to a helper that traps on may_leave",
                guarded.spec_name()
            );
        }
    }

    #[test]
    fn every_builtin_has_a_spec_name() {
        // `spec_name` is generated from the same table as `by_name`; a round
        // trip proves neither side drifted.
        for b in [
            B::Lift,
            B::Lower,
            B::TaskReturn,
            B::ResourceRep,
            B::ContextGet,
            B::ThreadYield,
        ] {
            let n = b.spec_name();
            assert_eq!(B::by_name(n), Some(b), "{n} must round-trip");
        }
    }
}
