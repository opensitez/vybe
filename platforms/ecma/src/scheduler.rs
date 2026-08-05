//! The ECMA-262 §9.5 job scheduler — the POLICY half of async, installed into
//! the VM's scheduler slot at plugin registration.
//!
//! This is the loop that used to live inside the WASM runtime (`jspi.rs`),
//! moved to the layer whose spec defines it: jobs and their drain-to-empty
//! discipline are ECMA-262 §8.6/§9.5 (`HostEnqueuePromiseJob`), the
//! one-task-per-turn structure is HTML's processing model. Time-deferred work
//! (timers) is not this crate's either — HTML owns it, `platforms/web`'s
//! wheel stores it, and the drain reaches it only through the VM's
//! `DeferredSource` registrations (`wasi:io/poll` shape).
//!
//! Stateless by design: all state (the ready queue, fibers, the host wheels)
//! lives on the VM or with the host that owns it; the scheduler reads it
//! through the mechanism surface (`run_scheduled_callback`,
//! `resume_scheduled_fiber`, `event_loop`, `next_due_deferred`).

use vybe_runtime::VM;
use vybe_runtime::error::VMError;
use vybe_runtime::event_loop::Task;
use vybe_runtime::scheduler::Scheduler;

pub struct EcmaScheduler;

impl Scheduler for EcmaScheduler {
    fn turn(&self, vm: &mut VM) -> Result<bool, VMError> {
        let mut ran = false;

        // 1. The job queue: drained to EMPTY before the next task — the §9.5
        // job checkpoint. This is mechanics, not a per-language property: a
        // language whose contract differs says so in its normalized AST ops
        // (which tier its lowerings enqueue on), never via a runtime flag.
        loop {
            let task = vm.event_loop.borrow_mut().next_immediate();
            let Some(task) = task else { break };
            ran = true;
            match task {
                Task::Callback { callback, value } => {
                    vm.run_scheduled_callback(&callback, &[value])?;
                }
                Task::ResumeFiber(fiber) => {
                    vm.resume_scheduled_fiber(fiber)?;
                }
            }
        }

        // 2. Wait for, then run, at most ONE deferred task (HTML: one task
        // per turn; the next job checkpoint follows on the next turn).
        if vm.deferred_pending() {
            self.wait(vm);
            if let Some(callback) = vm.next_due_deferred() {
                vm.run_scheduled_callback(&callback, &[])?;
                ran = true;
            }
        }
        Ok(ran)
    }

    fn has_pending(&self, vm: &VM) -> bool {
        vm.event_loop.borrow().has_pending() || vm.deferred_pending()
    }

    fn wait(&self, vm: &VM) {
        vm.wait_for_deferred();
    }
}
