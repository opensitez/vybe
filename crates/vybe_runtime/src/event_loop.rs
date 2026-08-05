//! Work queues — the MECHANISM a host schedules on.
//!
//! Deliberately free of host vocabulary. "Microtask", "macrotask", "job" and
//! "task" are ECMA-262 §9.5 and HTML concepts; naming VM state after them is
//! what made one language's contract everyone's. This layer owns two things
//! WASM can justify: ordered queues of pending work, and monotonic fire times.
//! WHICH tier a callback belongs in, and how far each is drained per turn, is
//! the host's policy — applied by the drain
//! loop.
//!
//! The storage stays here because it holds `Fiber`s, which are VM state
//! (JSPI / stack-switching). Only the naming moved.
//!
//! Timer scheduling uses a monotonic clock (wasi:clocks/monotonic-clock
//! semantics) so fire times are immune to wall-clock jumps.

use crate::fiber::Fiber;
use crate::value::Value;
use std::collections::{HashMap, VecDeque};

/// Close all open upvalues captured in a lambda Value.
/// Timer callbacks escape their creating stack frame and run in a fresh
/// execution context, so any Open(slot) upvalue would index an invalid stack.
/// This converts them to Closed(value) using the current stack snapshot.
#[allow(dead_code)]
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

/// Phase of a CM3 future in the EventLoop registry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FuturePhase {
    Pending,
    Resolved,
    Rejected }

/// Registered state for a single future<T>.
#[derive(Debug)]
pub struct FutureRecord {
    pub phase: FuturePhase,
    pub value: Option<Value> }

/// Registered state for a single stream<T>.
#[derive(Debug)]
pub struct StreamRecord {
    pub buffer: VecDeque<Value>,
    pub closed: bool }

/// A task in the event loop.
#[derive(Debug)]
pub enum Task {
    /// A suspended fiber waiting to resume with a value.
    ResumeFiber(Fiber),
    /// A callback with its argument. The host decides what an entry MEANS
    /// (an ECMA job, a settled reaction); the VM only orders them.
    Callback { callback: Value, value: Value } }

/// The event loop — manages pending async work.
#[derive(Debug)]
pub struct EventLoop {
    /// The ONE ready queue — work that became runnable, in arrival order.
    /// ECMA-262 calls what goes here a *job*; the VM does not. Time-deferred
    /// work (HTML's timer wheel) is host storage, registered as a
    /// `scheduler::DeferredSource` — it never lives in this struct.
    pub immediate: VecDeque<Task>,
    /// Suspended fibers waiting for Promise resolution.
    pub waiting_fibers: Vec<(u64, Fiber)>, // (promise_id, fiber)
    /// Next promise ID.
    next_promise_id: u64,
    /// CM3 future registry: future_id → FutureRecord
    pub future_states: HashMap<u64, FutureRecord>,
    /// Fibers suspended waiting for a specific future to resolve.
    pub future_waiting_fibers: Vec<(u64, Fiber)>, // (future_id, fiber)
    next_future_id: u64,
    /// CM3 stream registry: stream_id → StreamRecord
    pub stream_buffers: HashMap<u64, StreamRecord>,
    /// Fibers suspended waiting for data in a specific stream.
    pub stream_waiting_fibers: Vec<(u64, Fiber)>, // (stream_id, fiber)
    next_stream_id: u64 }

impl EventLoop {
    pub fn new() -> Self {
        EventLoop {
            immediate: VecDeque::new(),
            waiting_fibers: Vec::new(),
            next_promise_id: 1,
            future_states: HashMap::new(),
            future_waiting_fibers: Vec::new(),
            next_future_id: 1,
            stream_buffers: HashMap::new(),
            stream_waiting_fibers: Vec::new(),
            next_stream_id: 1 }
    }

    /// Generate a unique promise ID.
    pub fn next_promise_id(&mut self) -> u64 {
        let id = self.next_promise_id;
        self.next_promise_id += 1;
        id
    }

    /// Enqueue on tier 0. Named for the tier, not for what the host puts there.
    pub fn queue_immediate(&mut self, callback: Value, value: Value) {
        self.immediate
            .push_back(Task::Callback { callback, value });
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

    /// Pop the next ready entry.
    pub fn next_immediate(&mut self) -> Option<Task> {
        self.immediate.pop_front()
    }

    // ── CM3 futures ─────────────────────────────────────────────────────────

    /// Allocate a new future, returning its ID.
    pub fn create_future(&mut self) -> u64 {
        let id = self.next_future_id;
        self.next_future_id += 1;
        self.future_states.insert(
            id,
            FutureRecord {
                phase: FuturePhase::Pending,
                value: None },
        );
        id
    }

    /// Suspend a fiber waiting for the given future to resolve.
    pub fn suspend_future(&mut self, future_id: u64, fiber: Fiber) {
        self.future_waiting_fibers.push((future_id, fiber));
    }

    /// Resolve a future — wake the fiber waiting for it.
    pub fn resolve_future(&mut self, future_id: u64, value: Value) -> Option<Fiber> {
        if let Some(rec) = self.future_states.get_mut(&future_id) {
            rec.phase = FuturePhase::Resolved;
            rec.value = Some(value.clone());
        }
        if let Some(pos) = self
            .future_waiting_fibers
            .iter()
            .position(|(id, _)| *id == future_id)
        {
            let (_, mut fiber) = self.future_waiting_fibers.remove(pos);
            fiber.resume_value = Some(value);
            Some(fiber)
        } else {
            None
        }
    }

    /// Reject a future — wake the fiber waiting for it (will throw the reason).
    pub fn reject_future(&mut self, future_id: u64, reason: Value) -> Option<Fiber> {
        if let Some(rec) = self.future_states.get_mut(&future_id) {
            rec.phase = FuturePhase::Rejected;
            rec.value = Some(reason.clone());
        }
        if let Some(pos) = self
            .future_waiting_fibers
            .iter()
            .position(|(id, _)| *id == future_id)
        {
            let (_, mut fiber) = self.future_waiting_fibers.remove(pos);
            fiber.resume_exception = Some(reason);
            Some(fiber)
        } else {
            None
        }
    }

    // ── CM3 streams ─────────────────────────────────────────────────────────

    /// Allocate a new stream, returning its ID.
    pub fn create_stream(&mut self) -> u64 {
        let id = self.next_stream_id;
        self.next_stream_id += 1;
        self.stream_buffers.insert(
            id,
            StreamRecord {
                buffer: VecDeque::new(),
                closed: false },
        );
        id
    }

    /// Suspend a fiber waiting for the next item from a stream.
    pub fn suspend_stream_reader(&mut self, stream_id: u64, fiber: Fiber) {
        self.stream_waiting_fibers.push((stream_id, fiber));
    }

    /// Push one item to a stream. If a fiber is waiting, wake it directly with the item.
    /// Returns the fiber to queue as ResumeFiber, if any.
    pub fn stream_push(&mut self, stream_id: u64, item: Value) -> Option<Fiber> {
        if let Some(pos) = self
            .stream_waiting_fibers
            .iter()
            .position(|(id, _)| *id == stream_id)
        {
            // Direct wake — don't buffer, give item straight to the waiting fiber.
            let (_, mut fiber) = self.stream_waiting_fibers.remove(pos);
            fiber.resume_value = Some(item);
            Some(fiber)
        } else {
            if let Some(rec) = self.stream_buffers.get_mut(&stream_id) {
                rec.buffer.push_back(item);
            }
            None
        }
    }

    /// Pop one buffered item from a stream. Returns None if the buffer is empty.
    pub fn stream_pop(&mut self, stream_id: u64) -> Option<Value> {
        self.stream_buffers
            .get_mut(&stream_id)
            .and_then(|rec| rec.buffer.pop_front())
    }

    /// Check whether a stream's buffer is empty AND the stream is closed (EOF).
    pub fn stream_is_eof(&self, stream_id: u64) -> bool {
        self.stream_buffers
            .get(&stream_id)
            .map_or(true, |rec| rec.closed && rec.buffer.is_empty())
    }

    /// Check whether a stream has buffered items ready to read.
    pub fn stream_has_item(&self, stream_id: u64) -> bool {
        self.stream_buffers
            .get(&stream_id)
            .map_or(false, |rec| !rec.buffer.is_empty())
    }

    /// Close a stream. If a fiber is waiting, wake it with EOF (resume_value = None → Null).
    pub fn stream_close(&mut self, stream_id: u64) -> Option<Fiber> {
        if let Some(rec) = self.stream_buffers.get_mut(&stream_id) {
            rec.closed = true;
        }
        if let Some(pos) = self
            .stream_waiting_fibers
            .iter()
            .position(|(id, _)| *id == stream_id)
        {
            let (_, mut fiber) = self.stream_waiting_fibers.remove(pos);
            // EOF sentinel: resume_value stays None; dispatch will push Value::Null
            fiber.resume_value = Some(Value::Null);
            Some(fiber)
        } else {
            None
        }
    }

    /// Check if there's any pending work.
    pub fn has_pending(&self) -> bool {
        !self.immediate.is_empty()
    }
}

/// Monotonic milliseconds since first call.
/// Aligned with wasi:clocks/monotonic-clock semantics: values are only
/// meaningful relative to each other, not as wall-clock timestamps.
/// Public because host-owned deferred sources (the `platforms/web` timer
/// wheel) must stamp fire times on the SAME clock the drain compares against.
pub fn monotonic_now_ms() -> f64 {
    static START: std::sync::OnceLock<std::time::Instant> = std::sync::OnceLock::new();
    let start = START.get_or_init(std::time::Instant::now);
    start.elapsed().as_secs_f64() * 1000.0
}
