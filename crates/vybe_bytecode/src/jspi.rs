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
use std::sync::Arc;

impl VM {
    pub(crate) fn run_event_loop(&mut self) -> Result<(), VMError> {
        loop {
            let has_pending = self.event_loop.borrow().has_pending();
            if !has_pending {
                break;
            }

            // 1. Drain all microtasks — loop until none remain.
            // Microtask callbacks can schedule new microtasks (e.g.
            // .then() → .finally()), so drain iteratively per the
            // HTML spec microtask checkpoint algorithm.
            loop {
                let task = self.event_loop.borrow_mut().next_microtask();
                let Some(task) = task else { break };
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
        self.cur_fiber_id = fiber.fiber_id;
        self.cur_fiber_result_promise = fiber.result_promise;
        self.async_floors = fiber.async_floors;
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
        self.cur_fiber_id = fiber.fiber_id;
        self.cur_fiber_result_promise = fiber.result_promise;
        self.async_floors = fiber.async_floors;

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
        match self.execute_with_async() {
            Ok(ExecResult::Done(val)) => {
                // JSPI promising boundary: an async body that suspended earlier
                // has now run to completion — settle the pending result Promise
                // its caller holds (fulfil/reject with the body's outcome).
                if let Some(result_promise) = self.cur_fiber_result_promise.take() {
                    self.settle_async_result_promise(&result_promise, &val);
                }
                Ok(val)
            }
            Ok(ExecResult::Suspended { .. }) => Ok(Value::Null), // re-suspended, event loop will handle
            Err(e) => {
                // Uncaught JS throw out of a resumed async body → reject the
                // pending result Promise (§27.7: async throws become
                // rejections). Only a genuine thrown JS value qualifies —
                // an internal VM fault (no last_exception) must PROPAGATE,
                // not be laundered into a rejection.
                if let Some(exc) = self.last_exception.take() {
                    if let Some(result_promise) = self.cur_fiber_result_promise.take() {
                        self.settle_promise_via_host(&result_promise, "rejected", exc);
                        return Ok(Value::Null);
                    }
                    self.last_exception = Some(exc);
                }
                Err(e)
            }
        }
    }

    /// Settle an async call's result Promise with the completed body's return
    /// value. JS async bodies already return a settled Promise (the compiler's
    /// async_try wrap), so adopt its state; a raw value fulfils directly.
    /// Routed through the host's `ecma:promise.__settle_*` so pending `.then`
    /// reactions are drained and dependent awaiting fibers are woken.
    pub(crate) fn settle_async_result_promise(&mut self, result_promise: &Value, val: &Value) {
        let (state, inner) = {
            if let Value::Object(obj) = val {
                let o = obj.lock().unwrap();
                let is_promise = o
                    .properties
                    .get("__type")
                    .map(|t| format!("{}", t) == "Promise")
                    .unwrap_or(false);
                if is_promise {
                    let s = o
                        .properties
                        .get("__state")
                        .map(|v| format!("{}", v))
                        .unwrap_or_default();
                    let v = o
                        .properties
                        .get("__value")
                        .cloned()
                        .unwrap_or(Value::Undefined);
                    (s, v)
                } else {
                    ("fulfilled".to_string(), val.clone())
                }
            } else {
                ("fulfilled".to_string(), val.clone())
            }
        };
        match state.as_str() {
            "rejected" => self.settle_promise_via_host(result_promise, "rejected", inner),
            "pending" => {
                // Body returned a still-pending promise: adopt its eventual
                // state by registering host reactions that settle ours.
                self.settle_promise_via_host(result_promise, "adopt", val.clone());
            }
            _ => self.settle_promise_via_host(result_promise, "fulfilled", inner),
        }
    }

    /// JSPI `WebAssembly.promising` boundary — invoking an async function.
    ///
    /// Runs the async body inline until it either returns (result pushed,
    /// exactly like a plain call — JS async bodies return a settled Promise
    /// via the compiler's async_try wrap) or hits `await` on a pending
    /// promise. On suspension ONLY the async call's own frames / stack region /
    /// labels / handlers are captured as a fresh fiber registered against the
    /// awaited promise; the caller receives a PENDING result Promise and keeps
    /// executing — the JSPI-mandated delimited suspension (suspend to the
    /// promising export's boundary, resume off the event queue), never a
    /// whole-program block. On the fiber's final completion the result Promise
    /// is settled with the body's outcome (see `resume_fiber`).
    pub(crate) fn call_async(
        &mut self,
        func: &crate::value::Function,
        argc: usize,
    ) -> Result<(), VMError> {
        let floor = self.frames.len();
        self.async_floors.push(floor);
        if let Err(e) = self.call_function_direct(func, argc) {
            self.async_floors.pop();
            return Err(e);
        }
        if self.frames.len() <= floor {
            // Completed inline without a frame (host-style) — result on stack.
            self.async_floors.pop();
            return Ok(());
        }
        let call_base = self.frames[floor].base;
        let label_floor = self.frames[floor].label_base;
        match self.execute_until(floor + 1) {
            Ok(val) => {
                self.async_floors.pop();
                // The boundary return in execute_until skips the callee's
                // stack truncation — reclaim the callee's stack region.
                self.stack.truncate(call_base);
                self.push(val)?;
                Ok(())
            }
            Err(e) if e.message.starts_with("__jspi__:") => {
                self.async_floors.pop();
                let promise_id: u64 = e.message["__jspi__:".len()..].parse().unwrap_or(0);
                let result_promise =
                    self.suspend_async_call(floor, call_base, label_floor, promise_id);
                // Await on an ALREADY-SETTLED promise: JSPI still resumes via
                // the event queue — wake the just-registered fiber with an
                // immediate microtask carrying the settled value/rejection.
                if let Some((id, value, is_exception)) = self.pending_settled_await.take() {
                    if id == promise_id {
                        let mut el = self.event_loop.borrow_mut();
                        let woken = if is_exception {
                            el.reject_promise(id, value)
                        } else {
                            el.resolve_promise(id, value)
                        };
                        if let Some(fiber) = woken {
                            el.microtasks.push_back(Task::ResumeFiber(fiber));
                        }
                    }
                }
                self.push(result_promise)?;
                Ok(())
            }
            Err(e) => {
                self.async_floors.pop();
                Err(e)
            }
        }
    }

    /// Bounded suspension of an async call: capture the frames above `floor`
    /// and the stack above `call_base` as a fresh fiber (rebased to zero so a
    /// standard event-loop `resume_fiber` restores it), register it against
    /// the awaited promise, and hand back a new pending result Promise.
    fn suspend_async_call(
        &mut self,
        floor: usize,
        call_base: usize,
        label_floor: usize,
        promise_id: u64,
    ) -> Value {
        // Snapshot upvalues that point into the departing stack region.
        self.close_upvalues(call_base);

        let frames: Vec<SavedFrame> = self
            .frames
            .split_off(floor)
            .into_iter()
            .map(|f| SavedFrame {
                chunk_index: f.chunk_index,
                ip: f.ip,
                base: f.base - call_base,
                label_base: f.label_base - label_floor,
                upvalues: f.upvalues,
            })
            .collect();
        let stack: Vec<Value> = self.stack.split_off(call_base);
        let mut labels = self.label_stack.split_off(label_floor);
        for l in &mut labels {
            l.stack_height = l.stack_height.saturating_sub(call_base);
        }
        // Handlers installed by the body travel with it (rebased); the
        // caller's handlers stay.
        let mut kept = Vec::new();
        let mut captured = Vec::new();
        for mut h in std::mem::take(&mut self.exception_handlers) {
            if h.frame_depth > floor {
                h.frame_depth -= floor;
                h.stack_depth = h.stack_depth.saturating_sub(call_base);
                h.label_depth = h.label_depth.saturating_sub(label_floor);
                captured.push(h);
            } else {
                kept.push(h);
            }
        }
        self.exception_handlers = kept;

        let mut fiber = Fiber::new(stack, frames, Vec::new())
            .with_labels(labels)
            .with_exception_handlers(captured);
        fiber.fiber_id = self.next_fiber_id;
        self.next_fiber_id += 1;

        // Pending result Promise handed to the caller; settled at completion.
        let result_id = self.event_loop.borrow_mut().next_promise_id();
        let result_promise = Self::make_pending_promise(result_id);
        fiber.result_promise = Some(result_promise.clone());

        self.event_loop.borrow_mut().suspend_fiber(promise_id, fiber);
        result_promise
    }

    /// Build a bare pending Promise object (`__type`/`__state`/`__id`) — the
    /// same shape the ECMA host engine uses, so host `.then`/settle/adopt and
    /// `do_await` all interoperate with it.
    fn make_pending_promise(id: u64) -> Value {
        let mut obj = crate::value::Object::new();
        obj.properties
            .insert("__type".into(), Value::String(Arc::from("Promise")));
        obj.properties
            .insert("__state".into(), Value::String(Arc::from("pending")));
        obj.properties.insert("__id".into(), Value::F64(id as f64));
        Value::Object(Arc::new(std::sync::Mutex::new(obj)))
    }

    /// Invoke `ecma:promise.__settle_fulfilled` / `__settle_rejected` (or
    /// `__adopt` for pending adoption) on the host so reaction draining and
    /// fiber wake-ups happen through the one spec engine.
    pub(crate) fn settle_promise_via_host(&mut self, promise: &Value, state: &str, value: Value) {
        let name = match state {
            "rejected" => "__settle_rejected",
            "adopt" => "__adopt",
            _ => "__settle_fulfilled",
        };
        let idx = self
            .host_registry
            .get(&("ecma:promise".to_string(), name.to_string()))
            .copied();
        if let Some(idx) = idx {
            let host_fn = self.host_fns[idx].clone();
            let args = [promise.clone(), value];
            let mut ctx = self.make_host_context();
            let _ = host_fn(&mut ctx, &args);
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

        // Close ALL open upvalues before the stack is drained. An Open(slot)
        // upvalue indexes the live stack; after the switch a closure created on
        // this stack (e.g. a generator entry capturing enclosing locals) would
        // read out-of-bounds slots on the new stack. Closing snapshots the
        // value into the shared cell — same semantics as close-on-frame-return.
        self.close_upvalues(0);

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
        let mut fiber = Fiber::new(stack, frames, upvalues)
            .with_labels(labels)
            .with_exception_handlers(handlers)
            .with_continuations(conts)
            .with_fiber_id(self.cur_fiber_id);
        // A suspending async body carries its pending result Promise with it,
        // so completion after any number of re-suspensions can settle it.
        fiber.result_promise = self.cur_fiber_result_promise.take();
        // Async-call floors index THIS fiber's frame stack — swap them out with
        // the rest of the state so they never index another fiber's frames.
        fiber.async_floors = std::mem::take(&mut self.async_floors);
        fiber
    }
}
