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

use crate::error::VMError;
use crate::event_loop::Task;
use crate::fiber::{Fiber, SavedFrame};
use crate::value::Value;
use crate::vm::{CallFrame, ExecResult, VM};

impl VM {
    pub(crate) fn run_event_loop(&mut self) -> Result<(), VMError> {
        loop {
            let has_pending = self.event_loop.borrow().has_pending();
            if !has_pending {
                break;
            }

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
    pub fn resume_fiber_with(
        &mut self,
        fiber: Fiber,
        push_value: Option<Value>,
    ) -> Result<(), VMError> {
        self.stack = fiber.stack;
        self.frames = fiber
            .frames
            .into_iter()
            .map(|f| CallFrame {
                chunk_index: f.chunk_index,
                ip: f.ip,
                base: f.base,
                label_base: f.label_base,
                upvalues: f.upvalues,
            })
            .collect();
        self.open_upvalues = fiber.open_upvalues;
        self.label_stack = fiber.label_stack;
        self.exception_handlers = fiber.exception_handlers;
        self.active_continuations = fiber.active_continuations;
        if let Some(val) = push_value {
            self.push(val)?;
        }
        Ok(())
    }

    /// Resume a suspended fiber — restore its state and continue execution.
    pub(crate) fn resume_fiber(&mut self, fiber: Fiber) -> Result<Value, VMError> {
        // Restore state from fiber
        self.stack = fiber.stack;
        self.frames = fiber
            .frames
            .into_iter()
            .map(|f| CallFrame {
                chunk_index: f.chunk_index,
                ip: f.ip,
                base: f.base,
                label_base: f.label_base,
                upvalues: f.upvalues,
            })
            .collect();
        self.open_upvalues = fiber.open_upvalues;
        self.label_stack = fiber.label_stack;
        self.exception_handlers = fiber.exception_handlers;
        self.active_continuations = fiber.active_continuations;

        // Rejected promise: throw the reason into the resuming fiber so that
        // enclosing try/catch blocks fire correctly. This is the JSPI-compliant
        // behavior — rejected promise resumption is equivalent to a THROW at
        // the suspension point.
        if let Some(exc) = fiber.resume_exception {
            self.raise_exception_value(exc)?;
            // raise_exception_value either jumps to a handler or returns Err.
            // If it jumped (returned Ok implicitly via continue), fall through
            // to execute_with_async so the handler body runs.
        } else if let Some(val) = fiber.resume_value {
            // Push the fulfilled value (this is what `await` returns)
            self.push(val)?;
        }

        // Continue execution
        match self.execute_with_async()? {
            ExecResult::Done(val) => Ok(val),
            ExecResult::Suspended { .. } => Ok(Value::Null), // re-suspended, event loop will handle
        }
    }

    /// JSPI: Resolve a suspended promise and resume execution.
    /// Called by the runtime/event loop when an async operation completes.
    /// `promise_id` identifies which suspension to resume.
    /// `value` is the resolved value that becomes the return of the host call.
    pub fn jspi_resolve(&mut self, promise_id: u64, value: Value) -> Result<Value, VMError> {
        let fiber = self
            .event_loop
            .borrow_mut()
            .resolve_promise(promise_id, value);
        if let Some(fiber) = fiber {
            self.resume_fiber(fiber)
        } else {
            Ok(Value::Null)
        }
    }

    /// Check if there are any JSPI-suspended fibers waiting for resolution.
    pub fn has_pending_jspi(&self) -> bool {
        !self.event_loop.borrow().waiting_fibers.is_empty()
    }

    /// Save the current execution state to a Fiber.
    pub fn save_fiber(&mut self) -> Fiber {
        // Close open upvalues for all lambdas stored in the macrotask queue.
        // These callbacks escape the current stack frame — they will run in a
        // fresh execution context after this fiber suspends. Any Open(slot)
        // upvalue would then index an invalid stack. We resolve them now using
        // the current (about to be saved) stack.
        {
            let stack = &self.stack;
            use crate::value::{ObjectKind, UpvalueLocation};
            let el_ref = self.event_loop.borrow();
            for task in el_ref.macrotasks.iter() {
                let callback = match task {
                    crate::event_loop::Task::Timer { callback, .. } => callback,
                    _ => continue,
                };
                if let Value::Object(obj) = callback {
                    let o = obj.lock().unwrap();
                    if let ObjectKind::Function(func) = &o.kind {
                        for uv in &func.upvalues {
                            let mut u = uv.lock().unwrap();
                            if let UpvalueLocation::Open(slot) = u.location {
                                let val = stack.get(slot).cloned().unwrap_or(Value::Null);
                                u.location = UpvalueLocation::Closed(val);
                            }
                        }
                    }
                }
            }
        }

        let frames = self
            .frames
            .drain(..)
            .map(|f| SavedFrame {
                chunk_index: f.chunk_index,
                ip: f.ip,
                base: f.base,
                label_base: f.label_base,
                upvalues: f.upvalues,
            })
            .collect();
        let stack = self.stack.drain(..).collect();
        let upvalues = self.open_upvalues.drain(..).collect();
        let labels = self.label_stack.drain(..).collect();
        let handlers = self.exception_handlers.drain(..).collect();
        let conts = self.active_continuations.drain(..).collect();
        Fiber::new(stack, frames, upvalues)
            .with_labels(labels)
            .with_exception_handlers(handlers)
            .with_continuations(conts)
    }
}
