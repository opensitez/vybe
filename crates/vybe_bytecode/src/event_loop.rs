//! Event loop — processes async tasks, timers, and promise callbacks.
//! Follows the browser/Node.js model:
//!   1. Run synchronous code
//!   2. Drain microtask queue (Promise.then callbacks)
//!   3. Process one macrotask (setTimeout callback)
//!   4. Repeat until all queues empty
//!
//! Timer scheduling uses a monotonic clock (wasi:clocks/monotonic-clock
//! semantics) so fire times are immune to wall-clock jumps.

use crate::fiber::Fiber;
use crate::value::Value;
use std::collections::VecDeque;

/// Close all open upvalues captured in a lambda Value.
/// Timer callbacks escape their creating stack frame and run in a fresh
/// execution context, so any Open(slot) upvalue would index an invalid stack.
/// This converts them to Closed(value) using the current stack snapshot.
fn close_upvalues_in_value(val: &Value, stack: &[Value]) {
    use crate::value::{ObjectKind, UpvalueLocation};
    if let Value::Object(obj) = val {
        let o = obj.lock().unwrap();
        if let ObjectKind::Function(ref func) = o.kind {
            for uv in &func.upvalues {
                let mut u = uv.lock().unwrap();
                if let UpvalueLocation::Open(slot) = u.location {
                    let captured = stack.get(slot).cloned().unwrap_or(Value::Null);
                    u.location = UpvalueLocation::Closed(captured);
                }
            }
        }
    }
}

/// A task in the event loop.
#[derive(Debug)]
pub enum Task {
    /// A suspended fiber waiting to resume with a value.
    ResumeFiber(Fiber),
    /// A timer callback — function value + scheduled fire time (ms, monotonic)
    /// and a unique cancellable ID.
    Timer {
        callback: Value,
        fire_at_ms: f64,
        id: u64,
    },
    /// A microtask — Promise.then/catch callback with a value.
    Microtask { callback: Value, value: Value },
}

/// The event loop — manages pending async work.
#[derive(Debug)]
pub struct EventLoop {
    /// Microtask queue (Promise callbacks) — higher priority.
    pub microtasks: VecDeque<Task>,
    /// Macrotask queue (setTimeout callbacks).
    pub macrotasks: VecDeque<Task>,
    /// Suspended fibers waiting for Promise resolution.
    pub waiting_fibers: Vec<(u64, Fiber)>, // (promise_id, fiber)
    /// Next promise ID.
    next_promise_id: u64,
    /// Next timer ID (separate counter from promise IDs).
    next_timer_id: u64,
}

impl EventLoop {
    pub fn new() -> Self {
        EventLoop {
            microtasks: VecDeque::new(),
            macrotasks: VecDeque::new(),
            waiting_fibers: Vec::new(),
            next_promise_id: 1,
            next_timer_id: 1,
        }
    }

    /// Generate a unique promise ID.
    pub fn next_promise_id(&mut self) -> u64 {
        let id = self.next_promise_id;
        self.next_promise_id += 1;
        id
    }

    /// Schedule a microtask (Promise.then callback).
    pub fn queue_microtask(&mut self, callback: Value, value: Value) {
        self.microtasks
            .push_back(Task::Microtask { callback, value });
    }

    /// Schedule a macrotask (setTimeout callback). Does not return an ID.
    pub fn queue_timer(&mut self, callback: Value, delay_ms: f64) {
        let id = self.next_timer_id;
        self.next_timer_id += 1;
        let now = current_time_ms();
        self.macrotasks.push_back(Task::Timer {
            callback,
            fire_at_ms: now + delay_ms,
            id,
        });
    }

    /// Schedule a macrotask and return its cancellable ID.
    /// Use this from setTimeout host functions.
    pub fn queue_timer_id(&mut self, callback: Value, delay_ms: f64) -> u64 {
        let id = self.next_timer_id;
        self.next_timer_id += 1;
        let now = current_time_ms();
        self.macrotasks.push_back(Task::Timer {
            callback,
            fire_at_ms: now + delay_ms,
            id,
        });
        id
    }

    /// Cancel a timer by ID. Returns true if the timer was found and removed.
    pub fn cancel_timer(&mut self, id: u64) -> bool {
        if let Some(pos) = self
            .macrotasks
            .iter()
            .position(|t| matches!(t, Task::Timer { id: tid, .. } if *tid == id))
        {
            self.macrotasks.remove(pos);
            true
        } else {
            false
        }
    }

    /// Suspend a fiber — it will resume when the promise with the given ID resolves.
    pub fn suspend_fiber(&mut self, promise_id: u64, fiber: Fiber) {
        self.waiting_fibers.push((promise_id, fiber));
    }

    /// Resolve a promise — wake the fiber waiting for it (fulfillment).
    pub fn resolve_promise(&mut self, promise_id: u64, value: Value) -> Option<Fiber> {
        if let Some(pos) = self
            .waiting_fibers
            .iter()
            .position(|(id, _)| *id == promise_id)
        {
            let (_, mut fiber) = self.waiting_fibers.remove(pos);
            fiber.resume_value = Some(value);
            Some(fiber)
        } else {
            None
        }
    }

    /// Reject a promise — wake the fiber waiting for it (rejection).
    /// The fiber will throw the value instead of returning it.
    pub fn reject_promise(&mut self, promise_id: u64, reason: Value) -> Option<Fiber> {
        if let Some(pos) = self
            .waiting_fibers
            .iter()
            .position(|(id, _)| *id == promise_id)
        {
            let (_, mut fiber) = self.waiting_fibers.remove(pos);
            fiber.resume_exception = Some(reason);
            Some(fiber)
        } else {
            None
        }
    }

    /// Get the next ready microtask.
    pub fn next_microtask(&mut self) -> Option<Task> {
        self.microtasks.pop_front()
    }

    /// Get the next ready macrotask (timer whose fire time has passed).
    pub fn next_ready_timer(&mut self) -> Option<Task> {
        let now = current_time_ms();
        if let Some(pos) = self
            .macrotasks
            .iter()
            .position(|t| matches!(t, Task::Timer { fire_at_ms, .. } if *fire_at_ms <= now))
        {
            Some(self.macrotasks.remove(pos).unwrap())
        } else {
            None
        }
    }

    /// Check if there's any pending work.
    pub fn has_pending(&self) -> bool {
        !self.microtasks.is_empty()
            || !self.macrotasks.is_empty()
            || !self.waiting_fibers.is_empty()
    }

    /// Sleep until the next timer fires (or return immediately if microtasks pending).
    /// Uses the monotonic clock for accurate scheduling.
    pub fn wait_for_next(&self) {
        if !self.microtasks.is_empty() || !self.waiting_fibers.is_empty() {
            return; // microtasks are processed immediately
        }
        if let Some(earliest) = self
            .macrotasks
            .iter()
            .filter_map(|t| {
                if let Task::Timer { fire_at_ms, .. } = t {
                    Some(*fire_at_ms)
                } else {
                    None
                }
            })
            .reduce(f64::min)
        {
            let now = current_time_ms();
            if earliest > now {
                let sleep_ms = (earliest - now) as u64;
                // Native sleep — equivalent to wasi:clocks/monotonic-clock subscribe-duration.
                std::thread::sleep(std::time::Duration::from_millis(sleep_ms));
            }
        }
    }
}

/// Monotonic milliseconds since first call.
/// Aligned with wasi:clocks/monotonic-clock semantics: values are only
/// meaningful relative to each other, not as wall-clock timestamps.
fn current_time_ms() -> f64 {
    static START: std::sync::OnceLock<std::time::Instant> = std::sync::OnceLock::new();
    let start = START.get_or_init(std::time::Instant::now);
    start.elapsed().as_secs_f64() * 1000.0
}
