//! JS Promise Integration (JSPI) and fiber save/resume.
//!
//! The WASM stack-switching proposal's machinery — when user code does
//! `await` on a pending Promise, the current fiber is saved (stack +
//! frame stack snapshot) and the VM yields back to the event loop.
//! When the Promise resolves, `resume_fiber` restores the saved state
//! and execution continues from the suspended opcode.
//!
//! `run_event_loop` drains microtasks (Promise callbacks) and
//! macrotasks (timers) until no pending work remains.

use std::collections::HashMap;
use std::sync::{Arc, Mutex};
use crate::chunk::Chunk;
use crate::error::VMError;
use crate::event_loop::{EventLoop, Task};
use crate::fiber::{Fiber, SavedFrame};
use crate::opcode::Op;
use crate::shared_memory::SharedMemory;
use crate::value::{Function, Object, ObjectKind, Upvalue, UpvalueLocation, Value};
use crate::vm::{
    VM, CallFrame, ExceptionHandler, FinalizerEntry, LabelEntry,
    ExecResult, HostContext, HostFn, ImportTarget,
    MAX_FRAMES, MAX_STACK,
};

impl VM {
    pub(crate) fn run_event_loop(&mut self) -> Result<(), VMError> {
        loop {
            let has_pending = self.event_loop.borrow().has_pending();
            if !has_pending { break; }

            // 1. Drain all microtasks
            let microtasks: Vec<Task> = {
                let mut el = self.event_loop.borrow_mut();
                let mut tasks = Vec::new();
                while let Some(task) = el.next_microtask() {
                    tasks.push(task);
                }
                tasks
            };

            for task in microtasks {
                match task {
                    Task::Microtask { callback, value } => {
                        self.invoke(&callback, &[value])?;
                    }
                    Task::ResumeFiber(fiber) => {
                        self.resume_fiber(fiber)?;
                    }
                    _ => {}
                }
            }

            // 2. Wait for and process one macrotask (timer)
            {
                let el = self.event_loop.borrow();
                el.wait_for_next();
            }
            let timer = self.event_loop.borrow_mut().next_ready_timer();
            if let Some(Task::Timer { callback, .. }) = timer {
                self.invoke(&callback, &[])?;
            }
        }
        Ok(())
    }

    /// Restore a saved fiber as the current VM state without kicking
    /// the execution loop. Used by the stack-switching dispatch so
    /// SUSPEND can return to the caller's exact pre-RESUME snapshot
    /// (stack + frames + upvalues) while the outer dispatch loop
    /// continues executing from the restored `ip`. If `push_value` is
    /// `Some`, it's pushed onto the restored stack — that's the
    /// yielded value the caller of RESUME sees.
    pub(crate) fn resume_fiber_with(&mut self, fiber: Fiber, push_value: Option<Value>)
        -> Result<(), VMError>
    {
        self.stack = fiber.stack;
        self.frames = fiber.frames.into_iter().map(|f| CallFrame {
            chunk_index: f.chunk_index,
            ip: f.ip,
            base: f.base,
            upvalues: f.upvalues,
        }).collect();
        self.open_upvalues = fiber.open_upvalues;
        if let Some(val) = push_value { self.push(val)?; }
        Ok(())
    }

    /// Resume a suspended fiber — restore its state and continue execution.
    pub(crate) fn resume_fiber(&mut self, fiber: Fiber) -> Result<Value, VMError> {
        // Restore state from fiber
        self.stack = fiber.stack;
        self.frames = fiber.frames.into_iter().map(|f| CallFrame {
            chunk_index: f.chunk_index,
            ip: f.ip,
            base: f.base,
            upvalues: f.upvalues,
        }).collect();
        self.open_upvalues = fiber.open_upvalues;

        // Push the resolved value onto the stack (this is what `await` returns)
        if let Some(val) = fiber.resume_value {
            self.push(val)?;
        }

        // Continue execution
        match self.execute_with_async()? {
            ExecResult::Done(val) => Ok(val),
            ExecResult::Suspended(_) => Ok(Value::Null), // re-suspended, event loop will handle
        }
    }

    /// JSPI: Resolve a suspended promise and resume execution.
    /// Called by the runtime/event loop when an async operation completes.
    /// `promise_id` identifies which suspension to resume.
    /// `value` is the resolved value that becomes the return of the host call.
    pub fn jspi_resolve(&mut self, promise_id: u64, value: Value) -> Result<Value, VMError> {
        let fiber = self.event_loop.borrow_mut().resolve_promise(promise_id, value);
        if let Some(fiber) = fiber {
            self.resume_fiber(fiber)
        } else {
            Ok(Value::Null)
        }
    }

    /// Check if there are any JSPI-suspended fibers waiting for resolution.
    pub fn has_pending_jspi(&self) -> bool {
        self.event_loop.borrow().has_pending()
    }

    /// Save the current execution state to a Fiber.
    pub(crate) fn save_fiber(&mut self) -> Fiber {
        let frames = self.frames.drain(..).map(|f| SavedFrame {
            chunk_index: f.chunk_index,
            ip: f.ip,
            base: f.base,
            upvalues: f.upvalues,
        }).collect();
        let stack = self.stack.drain(..).collect();
        let upvalues = self.open_upvalues.drain(..).collect();
        Fiber::new(stack, frames, upvalues)
    }
}
