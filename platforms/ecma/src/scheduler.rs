//! The ECMA-262 §9.5 job scheduler — the POLICY half of async, installed into
//! the VM's scheduler slot at plugin registration.
//!
//! This is the loop that used to live inside the WASM runtime (`jspi.rs`),
//! moved to the layer whose spec defines it: jobs and their drain-to-empty
//! discipline are ECMA-262 §8.6/§9.5 (`HostEnqueuePromiseJob`), the
//! one-task-per-turn structure is HTML's processing model. The module's
//! DECLARED contract (`vm.scheduling`, from the profile's `[async]` section)
//! chooses between drain-to-empty (`TieredJobs`) and asyncio's
//! ready-at-turn-start FIFO (`SingleReadyQueue`).
//!
//! Stateless by design: all state (queues, fibers, the declared policy) lives
//! on the VM; the scheduler reads it through the mechanism surface
//! (`run_scheduled_callback`, `resume_scheduled_fiber`, `event_loop`).

use vybe_runtime::VM;
use vybe_runtime::error::VMError;
use vybe_runtime::event_loop::Task;
use vybe_runtime::scheduler::Scheduler;

pub struct EcmaScheduler;

impl Scheduler for EcmaScheduler {
    fn turn(&self, vm: &mut VM) -> Result<bool, VMError> {
        let mut ran = false;

        // 1. The job queue, under the module's declared discipline.
        let drain_to_empty = matches!(
            vm.scheduling.queues,
            vybe_ast::QueueDiscipline::TieredJobs
        );
        let mut ready_at_turn_start = if drain_to_empty {
            usize::MAX
        } else {
            vm.event_loop.borrow().immediate.len()
        };
        loop {
            if ready_at_turn_start == 0 {
                break;
            }
            ready_at_turn_start = ready_at_turn_start.saturating_sub(1);
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
                _ => {}
            }
        }

        // 2. Wait for, then run, at most ONE deferred task (HTML: one task
        // per turn; the next job checkpoint follows on the next turn).
        if vm.event_loop.borrow().has_pending() {
            self.wait(vm);
            let timer = vm.event_loop.borrow_mut().next_ready_timer();
            if let Some(Task::Timer { callback, .. }) = timer {
                vm.run_scheduled_callback(&callback, &[])?;
                ran = true;
            }
        }
        Ok(ran)
    }

    fn has_pending(&self, vm: &VM) -> bool {
        vm.event_loop.borrow().has_pending()
    }

    fn wait(&self, vm: &VM) {
        vm.event_loop.borrow().wait_for_next();
    }
}
