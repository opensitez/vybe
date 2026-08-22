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

/// Flatten one buffered stream item into `stream<u8>` bytes.
///
/// The same conversion [`VM::stream_drain`] applies, kept identical on purpose:
/// two byte-views of one stream that disagree would show up as corrupted output
/// far from here. `I32` is a single byte (a `u8` element), `String` is its
/// UTF-8 bytes, and an array is the concatenation of its elements.
fn flatten_into(item: &Value, out: &mut Vec<u8>) {
    use crate::value::ObjectKind;
    match item {
        Value::I32(b) => out.push(*b as u8),
        Value::I64(b) => out.push(*b as u8),
        Value::F64(b) => out.push(*b as u8),
        Value::String(s) => out.extend_from_slice(s.as_bytes()),
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                for e in elems.iter() {
                    flatten_into(e, out);
                }
            }
        }
        _ => {}
    }
}

/// Phase of a CM3 future in the EventLoop registry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FuturePhase {
    Pending,
    Resolved,
    Rejected,
}

/// Registered state for a single future<T>.
#[derive(Debug)]
pub struct FutureRecord {
    pub phase: FuturePhase,
    pub value: Option<Value>,
}

/// Registered state for a single stream<T>.
#[derive(Debug)]
pub struct StreamRecord {
    pub buffer: VecDeque<Value>,
    pub closed: bool,
    /// The `T` of `stream<T>`, when it is not `u8`.
    ///
    /// `None` means the raw byte stream every producer used before typed
    /// elements existed, and it keeps that path EXACTLY as it was: `u8` is not
    /// spelled as some 1-byte stand-in here, because `canon stream.read`'s
    /// byte path and its element path are different code, and conflating the
    /// two types is how a `stream<u8>` would silently acquire element strides.
    pub elem: Option<crate::component::ValType>,
    /// Bytes flattened out of `buffer` but not yet handed to a reader.
    /// `canon stream.read` copies a byte count, which need not land on an item
    /// boundary; the remainder waits here. See [`EventLoop::stream_read_bytes`].
    pub pending: Vec<u8>,
}

/// A task in the event loop.
#[derive(Debug)]
pub enum Task {
    /// A suspended fiber waiting to resume with a value.
    ResumeFiber(Fiber),
    /// A callback with its argument. The host decides what an entry MEANS
    /// (an ECMA job, a settled reaction); the VM only orders them.
    Callback { callback: Value, value: Value },
}

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
    next_stream_id: u64,
    /// Fibers parked inside a SYNCHRONOUS `canon stream.read`.
    ///
    /// Distinct from `stream_waiting_fibers`, and the difference is the whole
    /// point: a fiber there is HANDED the item and resumes with it as a value,
    /// which is the shape a `yield`-style read wants. A fiber here must find
    /// the data still IN the stream, because what resumes it is a copy into
    /// linear memory at a remembered `(ptr, n)` — see [`Fiber::pending_copy`].
    /// Waking one by consuming its item would lose that item outright.
    stream_sync_readers: Vec<(u64, Fiber)>, // (stream_id, fiber)
    /// The same, for a synchronous `canon future.read`.
    future_sync_readers: Vec<(u64, Fiber)>, // (future_id, fiber)
    /// Host functions that can produce more elements for a stream on demand,
    /// as `stream_id → (module, name)`.
    ///
    /// A `stream<T>` whose elements arrive over time — `wasi:sockets`'
    /// `listen()`, where each element is an inbound connection — cannot be
    /// filled by the call that created it: a host function returns once. The
    /// producer is what the reader consults instead, just before it would
    /// park: run the host's `accept` now, then look again.
    ///
    /// A (module, name) PAIR rather than a closure, because a closure would
    /// have to be `Send + Sync` and carry the VM state it needs to mint
    /// resources and push. This is the same indirection the VM already uses to
    /// reach `ecma:promise.resolve` internally, and it keeps every socket fact
    /// in `platforms/wasi` — the VM only knows that a stream may name one.
    stream_producers: HashMap<u64, (String, String)>,
}

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
            next_stream_id: 1,
            stream_sync_readers: Vec::new(),
            future_sync_readers: Vec::new(),
            stream_producers: HashMap::new(),
        }
    }

    /// Name the host function that can produce more elements for `stream_id`.
    pub fn set_stream_producer(&mut self, stream_id: u64, module: &str, name: &str) {
        self.stream_producers
            .insert(stream_id, (module.to_string(), name.to_string()));
    }

    /// The producer registered for `stream_id`, if any.
    pub fn stream_producer(&self, stream_id: u64) -> Option<(String, String)> {
        self.stream_producers.get(&stream_id).cloned()
    }

    /// Forget a stream's producer — called when the stream closes, so a
    /// listener that has gone away is not polled forever.
    pub fn clear_stream_producer(&mut self, stream_id: u64) {
        self.stream_producers.remove(&stream_id);
    }

    /// Park a fiber inside a synchronous `canon stream.read`.
    pub fn suspend_stream_sync_reader(&mut self, stream_id: u64, fiber: Fiber) {
        self.stream_sync_readers.push((stream_id, fiber));
    }

    /// Park a fiber inside a synchronous `canon future.read`.
    pub fn suspend_future_sync_reader(&mut self, future_id: u64, fiber: Fiber) {
        self.future_sync_readers.push((future_id, fiber));
    }

    /// Requeue every fiber parked on `stream_id`, leaving the data in place.
    ///
    /// Enqueued directly rather than returned: a push can release readers of
    /// several ends of the same stream, and the one-fiber return of
    /// [`stream_push`] has no way to say so. Silently returning only the first
    /// would park the rest forever.
    fn wake_stream_sync_readers(&mut self, stream_id: u64) {
        let mut i = 0;
        while i < self.stream_sync_readers.len() {
            if self.stream_sync_readers[i].0 == stream_id {
                let (_, fiber) = self.stream_sync_readers.remove(i);
                self.immediate.push_back(Task::ResumeFiber(fiber));
            } else {
                i += 1;
            }
        }
    }

    /// The same, for a future that has settled either way.
    fn wake_future_sync_readers(&mut self, future_id: u64) {
        let mut i = 0;
        while i < self.future_sync_readers.len() {
            if self.future_sync_readers[i].0 == future_id {
                let (_, fiber) = self.future_sync_readers.remove(i);
                self.immediate.push_back(Task::ResumeFiber(fiber));
            } else {
                i += 1;
            }
        }
    }

    /// Fibers parked in a synchronous copy with nothing left to wake them.
    ///
    /// The event loop ends when nothing is runnable. A fiber parked here at
    /// that moment is a DEADLOCK — its producer is gone — and reporting it is
    /// the difference between an error and a program that silently stops
    /// mid-read. Returned as a count so the caller can name it.
    pub fn parked_sync_copies(&self) -> usize {
        self.stream_sync_readers.len() + self.future_sync_readers.len()
    }

    /// Generate a unique promise ID.
    pub fn next_promise_id(&mut self) -> u64 {
        let id = self.next_promise_id;
        self.next_promise_id += 1;
        id
    }

    /// Enqueue on tier 0. Named for the tier, not for what the host puts there.
    pub fn queue_immediate(&mut self, callback: Value, value: Value) {
        self.immediate.push_back(Task::Callback { callback, value });
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
                value: None,
            },
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
        // Recorded first: a parked `future.read` resumes by reading
        // `future_states`, so it must find the settled phase already there.
        self.wake_future_sync_readers(future_id);
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
        // A rejected future is the writable end going away: the copy can never
        // happen, so the parked reader is released to answer `DROPPED` rather
        // than waiting on a value that will never arrive.
        self.wake_future_sync_readers(future_id);
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

    /// Allocate a new `stream<u8>`, returning its ID.
    pub fn create_stream(&mut self) -> u64 {
        self.create_stream_of(None)
    }

    /// Allocate a new `stream<T>`. `None` is `stream<u8>`.
    pub fn create_stream_of(&mut self, elem: Option<crate::component::ValType>) -> u64 {
        let id = self.next_stream_id;
        self.next_stream_id += 1;
        self.stream_buffers.insert(
            id,
            StreamRecord {
                buffer: VecDeque::new(),
                closed: false,
                pending: Vec::new(),
                elem,
            },
        );
        id
    }

    /// The `T` of `stream<T>`, or `None` for a byte stream.
    pub fn stream_elem(&self, stream_id: u64) -> Option<crate::component::ValType> {
        self.stream_buffers
            .get(&stream_id)
            .and_then(|rec| rec.elem.clone())
    }

    /// Pop up to `max` whole ITEMS off a typed stream.
    ///
    /// The counterpart of [`stream_read_bytes`] for `stream<T>` where `T` is
    /// not `u8`: `canon stream.read` counts ELEMENTS, and an element of a
    /// record type has no meaning as a byte run — flattening one the way the
    /// byte path does would hand the guest the record's field bytes with no
    /// layout at all.
    pub fn stream_read_items(&mut self, stream_id: u64, max: usize) -> Vec<Value> {
        let Some(rec) = self.stream_buffers.get_mut(&stream_id) else {
            return Vec::new();
        };
        let take = max.min(rec.buffer.len());
        rec.buffer.drain(..take).collect()
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
            // Buffered FIRST, then the parked synchronous readers are released
            // — they resume by copying out of that buffer, so waking them
            // before the item lands would have them find the stream still
            // empty and park again.
            self.wake_stream_sync_readers(stream_id);
            None
        }
    }

    /// Pop one buffered item from a stream. Returns None if the buffer is empty.
    pub fn stream_pop(&mut self, stream_id: u64) -> Option<Value> {
        self.stream_buffers
            .get_mut(&stream_id)
            .and_then(|rec| rec.buffer.pop_front())
    }

    /// Read up to `max` BYTES out of a `stream<u8>` — the unit `canon
    /// stream.read` copies in, as against [`stream_pop`]'s whole items.
    ///
    /// A buffered item is not necessarily one byte: a host that wrote a string
    /// pushed one `Value` holding many. So an item is flattened on demand and
    /// whatever the caller could not take stays in `pending`, to be handed out
    /// by the next read. Without that cursor a 100-byte item read with `n = 10`
    /// would either overrun the guest's buffer or silently drop 90 bytes —
    /// and the guest would have no way to detect either.
    ///
    /// Conversion matches [`VM::stream_drain`]: `I32` is one byte, `String` is
    /// its UTF-8 bytes, an array is the concatenation of its elements.
    pub fn stream_read_bytes(&mut self, stream_id: u64, max: usize) -> Vec<u8> {
        let Some(rec) = self.stream_buffers.get_mut(&stream_id) else {
            return Vec::new();
        };
        while rec.pending.len() < max {
            let Some(item) = rec.buffer.pop_front() else {
                break;
            };
            flatten_into(&item, &mut rec.pending);
        }
        let take = max.min(rec.pending.len());
        rec.pending.drain(..take).collect()
    }

    /// True when a `stream<u8>` has no bytes left AND no items that could
    /// produce any. Distinct from [`stream_has_item`], which cannot see bytes
    /// already flattened into the cursor.
    pub fn stream_has_bytes(&self, stream_id: u64) -> bool {
        self.stream_buffers
            .get(&stream_id)
            .is_some_and(|rec| !rec.pending.is_empty() || !rec.buffer.is_empty())
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
        // A closed stream takes no more elements, so its producer is dead
        // weight — and polling a listener whose stream is gone would keep
        // accepting connections nobody can read.
        self.stream_producers.remove(&stream_id);
        // EOF is a RESULT, not a reason to stay parked: a synchronous reader
        // waiting here now has an answer (`DROPPED`, or a final short copy of
        // whatever is still buffered), so it must be released too.
        self.wake_stream_sync_readers(stream_id);
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
