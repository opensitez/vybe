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
    /// declared discipline (`vm.scheduling`), then at most one deferred task.
    /// Returns `true` if any work ran.
    fn turn(&self, vm: &mut VM) -> Result<bool, VMError>;
    /// Anything pending (jobs, timers, suspended fibers awaiting settlement)?
    fn has_pending(&self, vm: &VM) -> bool;
    /// Block until the nearest deadline/readiness (`wasi:io/poll` shape).
    fn wait(&self, vm: &VM);
}
