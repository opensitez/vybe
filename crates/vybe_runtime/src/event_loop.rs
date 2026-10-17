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
use std::collections::{HashMap, HashSet, VecDeque};
use std::sync::Arc;

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
    Rejected,
}

/// Registered state for a single future<T>.
#[derive(Debug)]
pub struct FutureRecord {
    pub phase: FuturePhase,
    pub value: Option<Value>,
}

type ByteConsumer = Box<dyn FnMut(Option<&[u8]>) -> Result<(), Value>>;

struct HostByteConsumer {
    future: u64,
    consume: ByteConsumer,
}

impl std::fmt::Debug for HostByteConsumer {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("HostByteConsumer").field("future", &self.future).finish_non_exhaustive()
    }
}

/// Registered state for a single stream<T>.
#[derive(Debug)]
pub struct StreamRecord {
    pub buffer: VecDeque<StreamItem>,
    pub closed: bool,
    // EOF alone does not release the reader's ownership of buffered data.
    dropped_ends: u8,
    /// The `T` of `stream<T>`, when it is not `u8`.
    ///
    /// `None` means the raw byte stream every producer used before typed
    /// elements existed, and it keeps that path EXACTLY as it was: `u8` is not
    /// spelled as some 1-byte stand-in here, because `canon stream.read`'s
    /// byte path and its element path are different code, and conflating the
    /// two types is how a `stream<u8>` would silently acquire element strides.
    pub elem: Option<Arc<crate::component::ValType>>,
    /// Unconsumed suffix of a byte run already taken from `buffer`.
    /// `canon stream.read` copies a byte count, which need not land on an item
    /// boundary; the remainder waits here. See [`EventLoop::stream_read_bytes`].
    pub pending: Vec<u8>,
    pending_start: usize,
    zero_read: bool,
    zero_write: bool,
    read_completion: Option<u32>,
    write_completion: Option<u32>,
    /// Guest write progress remains pending until delivered. Host-produced
    /// buffered chunks do not use this accounting.
    guest_write_total: usize,
    guest_write_remaining: usize,
}

#[derive(Debug)]
pub enum StreamItem {
    Value(Value),
    Bytes(Vec<u8>),
}

impl StreamItem {
    fn into_value(self) -> Value {
        match self {
            Self::Value(value) => value,
            Self::Bytes(bytes) => Value::Object(crate::heap::alloc(
                crate::value::Object::new_array(
                    bytes.into_iter().map(|byte| Value::I32(byte as i32)).collect(),
                ),
            )),
        }
    }
}

fn stream_payload_is_self_contained(ty: &crate::component::ValType) -> bool {
    use crate::component::ValType;
    match ty {
        ValType::I32 | ValType::U32 | ValType::I64 | ValType::U64
        | ValType::F32 | ValType::F64 | ValType::S8 | ValType::U8
        | ValType::S16 | ValType::U16 | ValType::Bool | ValType::Char
        | ValType::String | ValType::Flags(_) => true,
        ValType::List(element) | ValType::ListFixed(element, _)
        | ValType::Option(element) => stream_payload_is_self_contained(element),
        ValType::Record(fields) => fields.iter()
            .all(|(_, field)| stream_payload_is_self_contained(field)),
        ValType::Variant(cases) => cases.iter().all(|(_, payload)|
            payload.as_ref().is_none_or(stream_payload_is_self_contained)),
        ValType::Result(ok, error) => [ok, error].into_iter().all(|payload|
            payload.as_deref().is_none_or(stream_payload_is_self_contained)),
        ValType::Own(_) | ValType::Borrow(_) | ValType::Stream(_)
        | ValType::Future(_) | ValType::ErrorContext | ValType::Any => false,
    }
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
    abandoned_futures: HashSet<u64>,
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
    stream_sync_readers: HashMap<u64, Vec<Fiber>>,
    /// The same, for a synchronous `canon future.read`.
    future_sync_readers: HashMap<u64, Vec<Fiber>>,
    sync_reader_count: usize,
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
    stream_consumers: HashMap<u64, HostByteConsumer>,
    pub(crate) pending_guest_consumers: VecDeque<u64>,
}

impl EventLoop {
    pub fn new() -> Self {
        EventLoop {
            immediate: VecDeque::new(),
            waiting_fibers: Vec::new(),
            next_promise_id: 1,
            future_states: HashMap::new(),
            abandoned_futures: HashSet::new(),
            future_waiting_fibers: Vec::new(),
            next_future_id: 1,
            stream_buffers: HashMap::new(),
            stream_waiting_fibers: Vec::new(),
            next_stream_id: 1,
            stream_sync_readers: HashMap::new(),
            future_sync_readers: HashMap::new(),
            sync_reader_count: 0,
            stream_producers: HashMap::new(),
            stream_consumers: HashMap::new(),
            pending_guest_consumers: VecDeque::new(),
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
        self.stream_sync_readers
            .entry(stream_id)
            .or_default()
            .push(fiber);
        self.sync_reader_count += 1;
    }

    /// Park a fiber inside a synchronous `canon future.read`.
    pub fn suspend_future_sync_reader(&mut self, future_id: u64, fiber: Fiber) {
        self.future_sync_readers
            .entry(future_id)
            .or_default()
            .push(fiber);
        self.sync_reader_count += 1;
    }

    /// Requeue every fiber parked on `stream_id`, leaving the data in place.
    ///
    /// Enqueued directly rather than returned: a push can release readers of
    /// several ends of the same stream, and the one-fiber return of
    /// [`stream_push`] has no way to say so. Silently returning only the first
    /// would park the rest forever.
    fn wake_stream_sync_readers(&mut self, stream_id: u64) {
        if let Some(readers) = self.stream_sync_readers.remove(&stream_id) {
            self.sync_reader_count -= readers.len();
            for fiber in readers {
                self.immediate.push_back(Task::ResumeFiber(fiber));
            }
        }
    }

    /// The same, for a future that has settled either way.
    fn wake_future_sync_readers(&mut self, future_id: u64) {
        if let Some(readers) = self.future_sync_readers.remove(&future_id) {
            self.sync_reader_count -= readers.len();
            for fiber in readers {
                self.immediate.push_back(Task::ResumeFiber(fiber));
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
        self.sync_reader_count
    }

    pub(crate) fn has_sync_copy_waiter(&self, id: u64, stream: bool, handle: u32) -> bool {
        let readers = if stream { &self.stream_sync_readers } else { &self.future_sync_readers };
        let matches = |fiber: &Fiber| fiber.pending_copy.as_ref()
            .is_some_and(|copy| copy.handle == handle);
        readers.get(&id).is_some_and(|fibers| fibers.iter().any(matches))
            || self.immediate.iter().any(|task| match task {
                Task::ResumeFiber(fiber) => matches(fiber),
                Task::Callback { .. } => false,
            })
    }

    /// Resume only this suspended core call, leaving unrelated host jobs in order.
    pub(crate) fn take_ready_copy(&mut self, fiber_id: u64, thread: Option<u32>) -> Option<Fiber> {
        let index = self.immediate.iter().position(|task| {
            matches!(task, Task::ResumeFiber(fiber)
                if fiber.fiber_id == fiber_id && fiber.cm_thread == thread
                    && fiber.pending_copy.is_some())
        })?;
        match self.immediate.remove(index)? {
            Task::ResumeFiber(fiber) => Some(fiber),
            _ => unreachable!("matched a copy fiber"),
        }
    }

    pub(crate) fn mark_copy_thread_exit(&mut self, end_id: u64, fiber_id: u64, thread: u32) -> bool {
        for readers in [&mut self.stream_sync_readers, &mut self.future_sync_readers] {
            if let Some(fiber) = readers.get_mut(&end_id).and_then(|fibers| fibers.iter_mut()
                .find(|fiber| fiber.fiber_id == fiber_id && fiber.cm_thread == Some(thread)))
            {
                fiber.cm_exit_on_completion = true;
                return true;
            }
        }
        false
    }

    /// Remove a blocked copy when a cooperative peer traps. The caller restores
    /// this stack before propagating the trap across its canonical boundary.
    pub(crate) fn take_parked_copy(&mut self, end_id: u64, fiber_id: u64, thread: Option<u32>) -> Option<Fiber> {
        for readers in [&mut self.stream_sync_readers, &mut self.future_sync_readers] {
            if let std::collections::hash_map::Entry::Occupied(mut entry) = readers.entry(end_id) {
                if let Some(index) = entry.get().iter().position(|fiber|
                    fiber.fiber_id == fiber_id && fiber.cm_thread == thread)
                {
                    let fiber = entry.get_mut().remove(index);
                    if entry.get().is_empty() { entry.remove(); }
                    self.sync_reader_count -= 1;
                    return Some(fiber);
                }
            }
        }
        None
    }

    /// The streams a synchronous reader is currently parked on.
    ///
    /// What the loop asks before it concludes the program is stuck: a stream
    /// with a registered producer is not stuck, it just has nobody who has
    /// been asked to fill it yet. Deduplicated because several readers may
    /// park on one stream and the producer only needs driving once.
    pub fn parked_stream_ids(&self) -> Vec<u64> {
        let mut ids: Vec<u64> = self.stream_sync_readers.keys().copied().collect();
        ids.sort_unstable();
        ids
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

    /// The last canonical end disappeared. Retain pending host producers until
    /// settlement, but do not retain their unobservable completion payloads.
    pub(crate) fn release_future(&mut self, id: u64) {
        if self.future_sync_readers.contains_key(&id)
            || self.future_waiting_fibers.iter().any(|(future, _)| *future == id) {
            return;
        }
        match self.future_states.get(&id).map(|record| record.phase) {
            Some(FuturePhase::Pending) => { self.abandoned_futures.insert(id); }
            Some(_) => { self.future_states.remove(&id); }
            None => {}
        }
    }

    /// Relink the surviving writer of `source` to the surviving reader of
    /// `destination`. No future value is copied while both are pending.
    pub fn forward_future(&mut self, source: u64, destination: u64) -> bool {
        if source == destination {
            return self.future_states.remove(&source).is_some();
        }
        if !self.future_states.contains_key(&destination) { return false; }
        let Some(record) = self.future_states.remove(&source) else { return false };
        match record.phase {
            FuturePhase::Pending => {}
            FuturePhase::Resolved => {
                if let Some(value) = record.value {
                    if let Some(fiber) = self.resolve_future(destination, value) {
                        self.immediate.push_back(Task::ResumeFiber(fiber));
                    }
                }
            }
            FuturePhase::Rejected => {
                if let Some(reason) = record.value {
                    if let Some(fiber) = self.reject_future(destination, reason) {
                        self.immediate.push_back(Task::ResumeFiber(fiber));
                    }
                }
            }
        }
        true
    }

    /// Suspend a fiber waiting for the given future to resolve.
    pub fn suspend_future(&mut self, future_id: u64, fiber: Fiber) {
        self.future_waiting_fibers.push((future_id, fiber));
    }

    /// Resolve a future — wake the fiber waiting for it.
    pub fn resolve_future(&mut self, future_id: u64, value: Value) -> Option<Fiber> {
        if self.abandoned_futures.remove(&future_id) {
            self.future_states.remove(&future_id);
            return None;
        }
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
        if self.abandoned_futures.remove(&future_id) {
            self.future_states.remove(&future_id);
            return None;
        }
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
                dropped_ends: 0,
                pending: Vec::new(),
                pending_start: 0,
                elem: elem.map(Arc::new),
                zero_read: false,
                zero_write: false,
                read_completion: None,
                write_completion: None,
                guest_write_total: 0,
                guest_write_remaining: 0,
            },
        );
        id
    }

    /// Append the source's owned chunks after already-buffered destination
    /// chunks, then let the source writer feed the destination reader directly.
    pub fn forward_stream(&mut self, source: u64, destination: u64) -> bool {
        if source == destination {
            self.stream_producers.remove(&source);
            return self.stream_buffers.remove(&source).is_some();
        }
        if !self.stream_buffers.contains_key(&destination) { return false; }
        let Some(mut source_record) = self.stream_buffers.remove(&source) else { return false };
        if source_record.pending_start < source_record.pending.len() {
            if source_record.pending_start != 0 {
                source_record.pending.drain(..source_record.pending_start);
            }
            source_record.buffer.push_front(StreamItem::Bytes(source_record.pending));
        }
        let destination_record = self.stream_buffers.get_mut(&destination).expect("checked above");
        if destination_record.buffer.is_empty() {
            destination_record.buffer = std::mem::take(&mut source_record.buffer);
        } else {
            destination_record.buffer.append(&mut source_record.buffer);
        }
        destination_record.closed = source_record.closed;
        destination_record.dropped_ends = (destination_record.dropped_ends & 1)
            | (source_record.dropped_ends & 2);
        let ready = !destination_record.buffer.is_empty() || destination_record.closed;
        self.stream_producers.remove(&destination);
        if let Some(producer) = self.stream_producers.remove(&source) {
            self.stream_producers.insert(destination, producer);
        }
        if ready {
            self.wake_stream_sync_readers(destination);
        }
        if let Some(consumer) = self.stream_consumers.remove(&destination) {
            self.set_byte_consumer(destination, consumer.future, consumer.consume)
                .expect("forwarded native consumer must retain its byte stream");
        }
        true
    }

    /// The `T` of `stream<T>`, or `None` for a byte stream.
    pub fn stream_elem(&self, stream_id: u64) -> Option<Arc<crate::component::ValType>> {
        self.stream_buffers
            .get(&stream_id)
            .and_then(|rec| rec.elem.clone())
    }

    pub fn stream_zero_read(&mut self, stream_id: u64) -> Option<crate::canon_copy::CopyResult> {
        use crate::canon_copy::{pack, CopyResult};
        let rec = self.stream_buffers.get_mut(&stream_id)?;
        if rec.guest_write_remaining != 0 || rec.pending_start < rec.pending.len() || !rec.buffer.is_empty() {
            return Some(CopyResult::Completed);
        }
        if rec.closed { return Some(CopyResult::Dropped); }
        let wake_writer = rec.zero_write;
        if wake_writer {
            rec.zero_write = false;
            rec.write_completion = Some(pack(CopyResult::Completed, 0));
        }
        rec.zero_read = true;
        if wake_writer { self.wake_stream_sync_readers(stream_id); }
        None
    }

    pub fn stream_zero_write(&mut self, stream_id: u64) -> Option<crate::canon_copy::CopyResult> {
        use crate::canon_copy::CopyResult;
        if self.stream_consumers.contains_key(&stream_id) { return Some(CopyResult::Completed); }
        let rec = self.stream_buffers.get_mut(&stream_id)?;
        if rec.closed { return Some(CopyResult::Dropped); }
        if rec.zero_read { return Some(CopyResult::Completed); }
        rec.zero_write = true;
        None
    }

    pub fn stream_zero_completion(&self, stream_id: u64, readable: bool) -> Option<u32> {
        let rec = self.stream_buffers.get(&stream_id)?;
        if readable { rec.read_completion } else { rec.write_completion }
    }

    pub(crate) fn complete_guest_stream_read(&mut self, stream_id: u64, packed: u32) {
        if let Some(rec) = self.stream_buffers.get_mut(&stream_id) {
            debug_assert!(rec.read_completion.is_none());
            rec.read_completion = Some(packed);
        }
    }

    pub fn take_stream_zero_completion(&mut self, stream_id: u64, readable: bool) -> Option<u32> {
        let rec = self.stream_buffers.get_mut(&stream_id)?;
        let completion = if readable {
            rec.read_completion.take()
        } else {
            let completion = rec.write_completion.take();
            if completion.is_some() {
                rec.guest_write_total = 0;
                rec.guest_write_remaining = 0;
            }
            completion
        };
        self.reclaim_closed_stream(stream_id);
        completion
    }

    pub fn stream_zero_write_pending(&self, stream_id: u64) -> bool {
        self.stream_buffers.get(&stream_id).is_some_and(|rec| rec.zero_write)
    }

    pub fn stream_guest_write_pending(&self, stream_id: u64) -> bool {
        self.stream_buffers.get(&stream_id).is_some_and(|rec| rec.guest_write_remaining != 0)
    }

    pub fn begin_guest_stream_write(&mut self, stream_id: u64, count: usize) {
        if let Some(rec) = self.stream_buffers.get_mut(&stream_id) {
            rec.guest_write_total = count;
            rec.guest_write_remaining = count;
            if rec.zero_read {
                rec.zero_read = false;
                rec.read_completion = Some(crate::canon_copy::pack(crate::canon_copy::CopyResult::Completed, 0));
            }
        }
        self.wake_stream_sync_readers(stream_id);
    }

    pub(crate) fn has_stream_consumer(&self, stream_id: u64) -> bool {
        self.stream_consumers.contains_key(&stream_id)
    }

    pub(crate) fn pending_stream_read_capacity(&self, stream_id: u64) -> Option<usize> {
        self.stream_sync_readers.get(&stream_id)?.iter().find_map(|fiber| {
            let copy = fiber.pending_copy.as_ref()?;
            matches!(&copy.kind, crate::fiber::PendingCopyKind::StreamBytes |
                crate::fiber::PendingCopyKind::StreamTyped(_)).then_some(copy.n)
        })
    }

    fn finish_guest_stream_read(rec: &mut StreamRecord, count: usize) -> bool {
        if rec.guest_write_remaining == 0 { return false; }
        rec.guest_write_remaining = rec.guest_write_remaining.saturating_sub(count);
        if count != 0 {
            rec.write_completion = Some(crate::canon_copy::pack(
                crate::canon_copy::CopyResult::Completed,
                (rec.guest_write_total - rec.guest_write_remaining) as u32));
            return true;
        }
        false
    }

    pub fn stream_zero_read_pending(&self, stream_id: u64) -> bool {
        self.stream_buffers.get(&stream_id).is_some_and(|rec| rec.zero_read)
    }

    pub fn cancel_stream_zero(&mut self, stream_id: u64, readable: bool) -> u32 {
        if let Some(rec) = self.stream_buffers.get_mut(&stream_id) {
            if readable {
                rec.zero_read = false;
                return rec.read_completion.take().map_or(0, |packed| packed >> 4);
            } else {
                rec.zero_write = false;
                rec.write_completion = None;
                let copied = rec.guest_write_total - rec.guest_write_remaining;
                if rec.guest_write_remaining != 0 {
                    rec.buffer.clear();
                    rec.pending.clear();
                    rec.pending_start = 0;
                }
                rec.guest_write_total = 0;
                rec.guest_write_remaining = 0;
                return copied as u32;
            }
        }
        0
    }

    pub fn stream_note_read_attempt(&mut self, stream_id: u64) {
        let mut completed = false;
        if let Some(rec) = self.stream_buffers.get_mut(&stream_id) {
            if rec.zero_write {
                rec.zero_write = false;
                rec.write_completion = Some(crate::canon_copy::pack(
                    crate::canon_copy::CopyResult::Completed, 0));
                completed = true;
            }
        }
        if completed { self.wake_stream_sync_readers(stream_id); }
    }

    /// Pop up to `max` whole ITEMS off a typed stream.
    ///
    /// The counterpart of [`stream_read_bytes`] for `stream<T>` where `T` is
    /// not `u8`: `canon stream.read` counts ELEMENTS, and an element of a
    /// record type has no meaning as a byte run — flattening one the way the
    /// byte path does would hand the guest the record's field bytes with no
    /// layout at all.
    pub fn stream_read_items(&mut self, stream_id: u64, max: usize) -> Vec<Value> {
        self.stream_note_read_attempt(stream_id);
        let Some(rec) = self.stream_buffers.get_mut(&stream_id) else {
            return Vec::new();
        };
        let take = max.min(rec.buffer.len());
        let items = rec.buffer.drain(..take).map(StreamItem::into_value).collect();
        if Self::finish_guest_stream_read(rec, take) {
            self.wake_stream_sync_readers(stream_id);
        }
        items
    }

    /// Suspend a fiber waiting for the next item from a stream.
    pub fn suspend_stream_reader(&mut self, stream_id: u64, fiber: Fiber) {
        self.stream_waiting_fibers.push((stream_id, fiber));
    }

    /// Push one item to a stream. If a fiber is waiting, wake it directly with the item.
    /// Returns the fiber to queue as ResumeFiber, if any.
    pub fn stream_push(&mut self, stream_id: u64, item: Value) -> Result<Option<Fiber>, &'static str> {
        let record = self.stream_buffers.get(&stream_id).ok_or("invalid typed stream")?;
        if record.elem.is_none() { return Err("byte streams require explicit byte buffers"); }
        if record.dropped_ends & 1 != 0
            && record.elem.as_deref().is_some_and(stream_payload_is_self_contained)
        { return Ok(None); }
        Ok(self.push_typed_item(stream_id, item))
    }

    fn push_typed_item(&mut self, stream_id: u64, item: Value) -> Option<Fiber> {
        let mut zero_read_notified = false;
        if let Some(rec) = self.stream_buffers.get_mut(&stream_id) {
            if rec.zero_read {
                rec.zero_read = false;
                rec.read_completion = Some(crate::canon_copy::pack(
                    crate::canon_copy::CopyResult::Completed, 0));
                zero_read_notified = true;
            }
        }
        if zero_read_notified { self.wake_stream_sync_readers(stream_id); }
        if let Some(pos) = self
            .stream_waiting_fibers
            .iter()
            .position(|(id, _)| *id == stream_id)
        {
            // Direct wake — don't buffer, give item straight to the waiting fiber.
            let (_, mut fiber) = self.stream_waiting_fibers.remove(pos);
            if let Some(rec) = self.stream_buffers.get_mut(&stream_id) {
                if Self::finish_guest_stream_read(rec, 1) {
                    self.wake_stream_sync_readers(stream_id);
                }
            }
            fiber.resume_value = Some(item);
            Some(fiber)
        } else {
            if let Some(rec) = self.stream_buffers.get_mut(&stream_id) {
                rec.buffer.push_back(StreamItem::Value(item));
            }
            // Buffered FIRST, then the parked synchronous readers are released
            // — they resume by copying out of that buffer, so waking them
            // before the item lands would have them find the stream still
            // empty and park again.
            self.wake_stream_sync_readers(stream_id);
            None
        }
    }

    /// Buffer a byte run as one internal item. This changes no stream handle or
    /// canonical ABI shape; `canon stream.read` still copies individual bytes.
    pub fn stream_push_bytes(&mut self, stream_id: u64, bytes: Vec<u8>) {
        if bytes.is_empty() {
            return;
        }
        if self.stream_consumers.contains_key(&stream_id) {
            let count = bytes.len();
            self.consume_host_bytes(stream_id, Some(&bytes));
            let completed = self.stream_buffers.get_mut(&stream_id)
                .is_some_and(|record| !record.closed && Self::finish_guest_stream_read(record, count));
            if completed { self.wake_stream_sync_readers(stream_id); }
            return;
        }
        if let Some(rec) = self.stream_buffers.get_mut(&stream_id) {
            assert!(rec.elem.is_none(), "byte run pushed to typed stream");
            if rec.dropped_ends & 1 != 0 { return; }
            if rec.zero_read {
                rec.zero_read = false;
                rec.read_completion = Some(crate::canon_copy::pack(
                    crate::canon_copy::CopyResult::Completed, 0));
            }
            if let Some(pos) = self
                .stream_waiting_fibers
                .iter()
                .position(|(id, _)| *id == stream_id)
            {
                let (_, mut fiber) = self.stream_waiting_fibers.remove(pos);
                if Self::finish_guest_stream_read(rec, bytes.len()) {
                    self.wake_stream_sync_readers(stream_id);
                }
                fiber.resume_value = Some(StreamItem::Bytes(bytes).into_value());
                self.immediate.push_back(Task::ResumeFiber(fiber));
                return;
            }
            rec.buffer.push_back(StreamItem::Bytes(bytes));
            self.wake_stream_sync_readers(stream_id);
        }
    }

    /// Pop one buffered item from a stream. Returns None if the buffer is empty.
    pub fn stream_pop(&mut self, stream_id: u64) -> Option<Value> {
        self.stream_buffers
            .get_mut(&stream_id)
            .and_then(|rec| rec.buffer.pop_front())
            .map(StreamItem::into_value)
    }

    /// Read up to `max` BYTES out of a `stream<u8>` — the unit `canon
    /// stream.read` copies in, as against [`stream_pop`]'s whole items.
    ///
    /// A buffered run can contain many bytes. Whatever the caller could not
    /// take stays in `pending`, to be handed out by the next read. Without
    /// that cursor a 100-byte run read with `n = 10`
    /// would either overrun the guest's buffer or silently drop 90 bytes —
    /// and the guest would have no way to detect either.
    ///
    /// Byte runs are copied directly; typed items are not byte-coerced.
    pub fn stream_read_bytes(&mut self, stream_id: u64, max: usize) -> Vec<u8> {
        self.stream_note_read_attempt(stream_id);
        let Some(rec) = self.stream_buffers.get_mut(&stream_id) else {
            return Vec::new();
        };
        assert!(rec.elem.is_none(), "typed streams require typed reads");
        if rec.pending_start == rec.pending.len()
            && matches!(rec.buffer.front(), Some(StreamItem::Bytes(bytes))
                if bytes.len() <= max && (rec.buffer.len() == 1 || bytes.len() == max))
        {
            if let Some(StreamItem::Bytes(bytes)) = rec.buffer.pop_front() {
                if Self::finish_guest_stream_read(rec, bytes.len()) {
                    self.wake_stream_sync_readers(stream_id);
                }
                return bytes;
            }
        }
        while rec.pending.len() - rec.pending_start < max {
            let Some(item) = rec.buffer.pop_front() else {
                break;
            };
            if rec.pending_start != 0 {
                rec.pending.drain(..rec.pending_start);
                rec.pending_start = 0;
            }
            match item {
                StreamItem::Value(_) => unreachable!("typed item in a byte stream"),
                StreamItem::Bytes(bytes) if rec.pending.is_empty() => rec.pending = bytes,
                StreamItem::Bytes(bytes) => rec.pending.extend_from_slice(&bytes),
            }
        }
        let take = max.min(rec.pending.len() - rec.pending_start);
        if max == usize::MAX && rec.pending_start == 0 {
            let bytes = std::mem::take(&mut rec.pending);
            if Self::finish_guest_stream_read(rec, bytes.len()) {
                self.wake_stream_sync_readers(stream_id);
            }
            return bytes;
        }
        let bytes = rec.pending[rec.pending_start..rec.pending_start + take].to_vec();
        rec.pending_start += take;
        if rec.pending_start == rec.pending.len() {
            rec.pending.clear();
            rec.pending_start = 0;
        }
        if Self::finish_guest_stream_read(rec, bytes.len()) {
            self.wake_stream_sync_readers(stream_id);
        }
        bytes
    }

    pub fn stream_close_readable(&mut self, stream_id: u64) -> Option<Fiber> {
        if let Some(rec) = self.stream_buffers.get_mut(&stream_id) {
            rec.dropped_ends |= 1;
            // Only ownership-bearing payloads need canonical disposal rather
            // than ordinary value release (including resources nested in records).
            if rec.elem.as_deref().is_none_or(stream_payload_is_self_contained) {
                rec.buffer = VecDeque::new();
                rec.pending = Vec::new();
                rec.pending_start = 0;
            }
            if rec.guest_write_total != 0 {
                let copied = rec.guest_write_total - rec.guest_write_remaining;
                rec.write_completion = Some(crate::canon_copy::pack(
                    crate::canon_copy::CopyResult::Dropped, copied as u32));
                rec.guest_write_remaining = 0;
            }
        }
        self.close_stream_state(stream_id)
    }

    /// True when a `stream<u8>` has no bytes left AND no items that could
    /// produce any. Distinct from [`stream_has_item`], which cannot see bytes
    /// already moved into the cursor.
    pub fn stream_has_bytes(&self, stream_id: u64) -> bool {
        self.stream_buffers
            .get(&stream_id)
            .is_some_and(|rec| rec.pending_start < rec.pending.len() || !rec.buffer.is_empty())
    }

    /// Check whether a stream's buffer is empty AND the stream is closed (EOF).
    pub fn stream_is_eof(&self, stream_id: u64) -> bool {
        self.stream_buffers
            .get(&stream_id)
            .map_or(true, |rec| {
                rec.closed && rec.buffer.is_empty() && rec.pending_start == rec.pending.len()
            })
    }

    /// Check whether a stream has buffered items ready to read.
    pub fn stream_has_item(&self, stream_id: u64) -> bool {
        self.stream_buffers
            .get(&stream_id)
            .map_or(false, |rec| rec.pending_start < rec.pending.len() || !rec.buffer.is_empty())
    }

    /// Close a stream. If a fiber is waiting, wake it with EOF (resume_value = None → Null).
    pub fn stream_close(&mut self, stream_id: u64) -> Option<Fiber> {
        if let Some(rec) = self.stream_buffers.get_mut(&stream_id) {
            rec.dropped_ends |= 2;
        }
        self.close_stream_state(stream_id)
    }

    fn reclaim_closed_stream(&mut self, stream_id: u64) {
        let reclaim = self.stream_buffers.get(&stream_id).is_some_and(|rec| {
            rec.dropped_ends == 3 && rec.closed
                && rec.buffer.is_empty() && rec.pending.is_empty()
                && !rec.zero_read && !rec.zero_write
                && rec.read_completion.is_none() && rec.write_completion.is_none()
                && rec.guest_write_remaining == 0
        });
        if reclaim && !self.stream_consumers.contains_key(&stream_id)
            && !self.stream_producers.contains_key(&stream_id)
            && !self.stream_sync_readers.contains_key(&stream_id)
            && !self.stream_waiting_fibers.iter().any(|(id, _)| *id == stream_id)
        {
            self.stream_buffers.remove(&stream_id);
        }
    }

    fn close_stream_state(&mut self, stream_id: u64) -> Option<Fiber> {
        self.consume_host_bytes(stream_id, None);
        if let Some(rec) = self.stream_buffers.get_mut(&stream_id) {
            rec.closed = true;
            if rec.zero_read {
                rec.zero_read = false;
                rec.read_completion = Some(crate::canon_copy::pack(
                    crate::canon_copy::CopyResult::Dropped, 0));
            }
            if rec.zero_write {
                rec.zero_write = false;
                rec.write_completion = Some(crate::canon_copy::pack(
                    crate::canon_copy::CopyResult::Dropped, 0));
            }
        }
        // A closed stream takes no more elements, so its producer is dead
        // weight — and polling a listener whose stream is gone would keep
        // accepting connections nobody can read.
        self.stream_producers.remove(&stream_id);
        // EOF is a RESULT, not a reason to stay parked: a synchronous reader
        // waiting here now has an answer (`DROPPED`, or a final short copy of
        // whatever is still buffered), so it must be released too.
        self.wake_stream_sync_readers(stream_id);
        let fiber = if let Some(pos) = self
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
        };
        self.reclaim_closed_stream(stream_id);
        fiber
    }

    /// Install a host-owned byte buffer. `None` finalizes it at writer EOF.
    /// Existing buffered producers are drained once; open streams remain pending.
    pub fn set_byte_consumer(
        &mut self, stream_id: u64, future: u64, mut consume: ByteConsumer,
    ) -> Result<(), &'static str> {
        let Some(record) = self.stream_buffers.get(&stream_id) else {
            return Err("invalid stream");
        };
        if record.elem.is_some() || self.stream_consumers.contains_key(&stream_id) {
            return Err("stream is not an available byte stream");
        }
        self.stream_note_read_attempt(stream_id);
        let record = self.stream_buffers.get_mut(&stream_id).unwrap();
        let closed = record.closed;
        let pending = std::mem::take(&mut record.pending);
        let pending_start = std::mem::replace(&mut record.pending_start, 0);
        let buffered = std::mem::take(&mut record.buffer);
        let mut count = pending.len() - pending_start;
        for item in &buffered {
            match item {
                StreamItem::Bytes(bytes) => count += bytes.len(),
                StreamItem::Value(_) => unreachable!("typed item in a byte stream"),
            }
        }
        // Preserve the drain's copy-completion transition before callbacks,
        // without concatenating independently owned byte runs.
        if Self::finish_guest_stream_read(record, count) {
            self.wake_stream_sync_readers(stream_id);
        }
        let mut outcome = Ok(());
        if pending_start < pending.len() {
            outcome = consume(Some(&pending[pending_start..]));
        }
        for item in buffered {
            if outcome.is_err() { break; }
            match item {
                StreamItem::Bytes(bytes) if !bytes.is_empty() => outcome = consume(Some(&bytes)),
                StreamItem::Bytes(_) => {},
                StreamItem::Value(_) => unreachable!("typed item in a byte stream"),
            }
        }
        if closed && outcome.is_ok() { outcome = consume(None); }
        if closed || outcome.is_err() {
            self.finish_host_byte_consumer(stream_id, future, outcome);
        } else {
            self.stream_consumers.insert(stream_id, HostByteConsumer { future, consume });
            if self.stream_guest_write_pending(stream_id) {
                self.pending_guest_consumers.push_back(stream_id);
            }
        }
        Ok(())
    }

    fn consume_host_bytes(&mut self, stream_id: u64, bytes: Option<&[u8]>) {
        let Some(consumer) = self.stream_consumers.get_mut(&stream_id) else { return };
        let outcome = (consumer.consume)(bytes);
        if bytes.is_some() && outcome.is_ok() {
            return;
        }
        let future = consumer.future;
        self.stream_consumers.remove(&stream_id);
        self.finish_host_byte_consumer(stream_id, future, outcome);
    }

    fn finish_host_byte_consumer(&mut self, stream_id: u64, future: u64, outcome: Result<(), Value>) {
        if let Some(record) = self.stream_buffers.get_mut(&stream_id) {
            record.dropped_ends |= 1;
        }
        if outcome.is_err() {
            if let Some(record) = self.stream_buffers.get_mut(&stream_id) { record.closed = true; }
        }
        if let Some(fiber) = self.resolve_future(future, outcome.err().unwrap_or(Value::Null)) {
            self.immediate.push_back(Task::ResumeFiber(fiber));
        }
        self.reclaim_closed_stream(stream_id);
    }

    /// Check if there's any pending work.
    pub fn has_pending(&self) -> bool {
        !self.immediate.is_empty()
    }
}

#[cfg(test)]
mod host_byte_consumer_tests {
    use super::*;
    use std::cell::{Cell, RefCell};
    use std::rc::Rc;

    #[test]
    fn buffered_sink_lends_runs_without_joining_or_replaying_consumed_prefix() {
        let mut events = EventLoop::new();
        let stream = events.create_stream();
        let future = events.create_future();
        events.stream_push_bytes(stream, vec![0, 1, 2]);
        assert_eq!(events.stream_read_bytes(stream, 1), [0]);
        let run = vec![3, 4, 5];
        let address = run.as_ptr() as usize;
        events.stream_push_bytes(stream, run);
        let seen = Rc::new(RefCell::new(Vec::new()));
        let captured = seen.clone();
        let lent_run = Rc::new(Cell::new(false));
        let captured_run = lent_run.clone();
        events.set_byte_consumer(stream, future, Box::new(move |bytes| {
            if let Some(bytes) = bytes {
                if bytes == [3, 4, 5] {
                    assert_eq!(bytes.as_ptr() as usize, address);
                    captured_run.set(true);
                }
                captured.borrow_mut().extend_from_slice(bytes);
            }
            Ok(())
        })).unwrap();
        assert_eq!(&*seen.borrow(), &[1, 2, 3, 4, 5]);
        assert!(lent_run.get());
        assert!(!events.stream_has_bytes(stream));
        assert_eq!(events.future_states[&future].phase, FuturePhase::Pending);
        events.stream_close(stream);
        assert_eq!(events.future_states[&future].phase, FuturePhase::Resolved);
    }

    #[test]
    fn buffered_sink_stops_on_error_without_eof_or_later_callbacks() {
        let mut events = EventLoop::new();
        let stream = events.create_stream();
        let future = events.create_future();
        events.stream_push_bytes(stream, vec![1]);
        events.stream_push_bytes(stream, vec![2]);
        events.stream_close(stream);
        let calls = Rc::new(Cell::new(0));
        let captured = calls.clone();
        events.set_byte_consumer(stream, future, Box::new(move |_| {
            captured.set(captured.get() + 1);
            Err(Value::I32(42))
        })).unwrap();
        assert_eq!(calls.get(), 1);
        assert!(matches!(events.future_states[&future].value, Some(Value::I32(42))));
        assert!(!events.stream_has_bytes(stream));
    }

    #[test]
    fn byte_stream_rejects_value_producers_without_coercion() {
        let mut events = EventLoop::new();
        let stream = events.create_stream();
        let future = events.create_future();
        assert!(events.stream_push(stream, Value::I32(255)).is_err());
        assert!(events.stream_push(stream, Value::String(std::sync::Arc::from("ab"))).is_err());
        events.stream_push_bytes(stream, vec![0, 128, 255]);
        events.stream_close(stream);
        let output = Rc::new(RefCell::new(Vec::new()));
        let captured = output.clone();
        events.set_byte_consumer(stream, future, Box::new(move |bytes| {
            if let Some(bytes) = bytes { captured.borrow_mut().extend_from_slice(bytes); }
            Ok(())
        })).unwrap();
        assert_eq!(&*output.borrow(), &[0, 128, 255]);
        assert_eq!(events.future_states[&future].phase, FuturePhase::Resolved);
    }

    #[test]
    fn open_sink_accepts_bulk_chunks_and_resolves_only_at_eof() {
        let mut events = EventLoop::new();
        let stream = events.create_stream();
        let future = events.create_future();
        let output = Rc::new(RefCell::new(Vec::new()));
        let captured = output.clone();
        events.set_byte_consumer(stream, future, Box::new(move |bytes| {
            if let Some(bytes) = bytes { captured.borrow_mut().extend_from_slice(bytes); }
            Ok(())
        })).unwrap();
        events.stream_push_bytes(stream, vec![0, 255, 1]);
        events.stream_push_bytes(stream, vec![2; 65536]);
        events.stream_push_bytes(stream, vec![3]);
        assert_eq!(events.future_states[&future].phase, FuturePhase::Pending);
        assert!(!events.stream_has_bytes(stream));
        assert_eq!(output.borrow().len(), 65540);
        assert_eq!(&output.borrow()[..3], &[0, 255, 1]);
        events.stream_close(stream);
        assert_eq!(events.future_states[&future].phase, FuturePhase::Resolved);
        assert!(matches!(events.future_states[&future].value, Some(Value::Null)));
        assert!(!events.stream_consumers.contains_key(&stream));
    }

    #[test]
    fn sink_error_closes_reader_and_is_not_overwritten_by_eof() {
        let mut events = EventLoop::new();
        let stream = events.create_stream();
        let future = events.create_future();
        events.set_byte_consumer(stream, future, Box::new(|_| Err(Value::I32(42)))).unwrap();
        events.stream_push_bytes(stream, vec![1]);
        assert!(events.stream_is_eof(stream));
        assert!(matches!(events.future_states[&future].value, Some(Value::I32(42))));
        events.stream_close(stream);
        assert!(matches!(events.future_states[&future].value, Some(Value::I32(42))));
    }

    #[test]
    fn closed_producer_and_empty_stream_finalize_once() {
        for bytes in [vec![], vec![0, 128, 255]] {
            let mut events = EventLoop::new();
            let stream = events.create_stream();
            let future = events.create_future();
            events.stream_push_bytes(stream, bytes.clone());
            events.stream_close(stream);
            let seen = Rc::new(RefCell::new((Vec::new(), 0)));
            let captured = seen.clone();
            events.set_byte_consumer(stream, future, Box::new(move |bytes| {
                let mut seen = captured.borrow_mut();
                if let Some(bytes) = bytes { seen.0.extend_from_slice(bytes); } else { seen.1 += 1; }
                Ok(())
            })).unwrap();
            events.stream_close(stream);
            assert_eq!(seen.borrow().0, bytes);
            assert_eq!(seen.borrow().1, 1);
            assert_eq!(events.future_states[&future].phase, FuturePhase::Resolved);
        }
    }

    #[test]
    fn forwarded_chunks_reach_the_existing_host_reader_in_order() {
        let mut events = EventLoop::new();
        let source = events.create_stream();
        let destination = events.create_stream();
        let future = events.create_future();
        let output = Rc::new(RefCell::new(Vec::new()));
        let captured = output.clone();
        events.set_byte_consumer(destination, future, Box::new(move |bytes| {
            if let Some(bytes) = bytes { captured.borrow_mut().extend_from_slice(bytes); }
            Ok(())
        })).unwrap();
        events.stream_push_bytes(destination, vec![1]);
        events.stream_push_bytes(source, vec![2, 3]);
        assert!(events.forward_stream(source, destination));
        events.stream_push_bytes(destination, vec![4]);
        assert_eq!(&*output.borrow(), &[1, 2, 3, 4]);
        assert_eq!(events.future_states[&future].phase, FuturePhase::Pending);
        events.stream_close(destination);
        assert_eq!(events.future_states[&future].phase, FuturePhase::Resolved);
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

#[cfg(test)]
mod stream_tests {
    use super::{EventLoop, FuturePhase, Task, stream_payload_is_self_contained};
    use crate::canon_copy::{pack, CopyResult};
    use crate::fiber::Fiber;
    use crate::value::Value;

    fn reader(marker: i32) -> Fiber {
        Fiber::new(vec![Value::I32(marker)], Vec::new(), Vec::new())
    }

    fn next_reader_marker(events: &mut EventLoop) -> i32 {
        let Some(Task::ResumeFiber(fiber)) = events.immediate.pop_front() else {
            panic!("expected a resumed reader");
        };
        let [Value::I32(marker)] = fiber.stack.as_slice() else {
            panic!("expected reader marker");
        };
        *marker
    }

    #[test]
    fn zero_length_write_then_read_notifies_writer_but_leaves_reader_pending() {
        let mut events = EventLoop::new();
        let id = events.create_stream();
        assert_eq!(events.stream_zero_write(id), None);
        assert!(events.stream_zero_write_pending(id));
        assert_eq!(events.stream_zero_read(id), None);
        assert_eq!(events.take_stream_zero_completion(id, false), Some(pack(CopyResult::Completed, 0)));
        assert_eq!(events.stream_zero_completion(id, true), None);
        assert!(events.stream_zero_read_pending(id));
        events.stream_push_bytes(id, vec![1, 2]);
        assert_eq!(events.take_stream_zero_completion(id, true), Some(pack(CopyResult::Completed, 0)));
        assert_eq!(events.stream_read_bytes(id, 2), [1, 2]);
    }

    #[test]
    fn zero_length_read_then_write_and_cancel() {
        let mut events = EventLoop::new();
        let id = events.create_stream();
        assert_eq!(events.stream_zero_read(id), None);
        assert_eq!(events.stream_zero_write(id), Some(CopyResult::Completed));
        assert!(events.stream_zero_read_pending(id));
        assert_eq!(events.cancel_stream_zero(id, true), 0);
        events.stream_push_bytes(id, vec![7]);
        assert_eq!(events.stream_zero_completion(id, true), None);

        assert_eq!(events.stream_zero_write(id), None);
        assert!(events.stream_zero_write_pending(id));
        assert!(events.stream_read_bytes(id, 1).len() == 1);
        assert_eq!(events.take_stream_zero_completion(id, false), Some(pack(CopyResult::Completed, 0)));
    }

    #[test]
    fn typed_stream_zero_read_is_notified_by_whole_element() {
        let mut events = EventLoop::new();
        let id = events.create_stream_of(Some(crate::component::ValType::U32));
        assert_eq!(events.stream_zero_read(id), None);
        events.stream_push(id, Value::I32(42)).unwrap();
        assert_eq!(events.take_stream_zero_completion(id, true), Some(pack(CopyResult::Completed, 0)));
        assert_eq!(events.stream_read_items(id, 1), [Value::I32(42)]);
    }

    #[test]
    fn guest_write_after_zero_read_completes_when_consumed_or_dropped() {
        let mut events = EventLoop::new();
        let id = events.create_stream();
        assert_eq!(events.stream_zero_read(id), None);
        events.stream_push_bytes(id, vec![1, 2, 3, 4]);
        events.begin_guest_stream_write(id, 4);
        assert!(events.stream_guest_write_pending(id));
        assert_eq!(events.stream_read_bytes(id, 2), [1, 2]);
        assert_eq!(events.stream_zero_completion(id, false), Some(pack(CopyResult::Completed, 2)));
        events.stream_close_readable(id);
        assert_eq!(events.take_stream_zero_completion(id, false), Some(pack(CopyResult::Dropped, 2)));
    }

    #[test]
    fn cancelling_partially_read_guest_write_discards_unread_bytes() {
        let mut events = EventLoop::new();
        let id = events.create_stream();
        assert_eq!(events.stream_zero_read(id), None);
        events.stream_push_bytes(id, vec![1, 2, 3, 4]);
        events.begin_guest_stream_write(id, 4);
        assert_eq!(events.stream_read_bytes(id, 2), [1, 2]);
        assert_eq!(events.cancel_stream_zero(id, false), 2);
        assert!(!events.stream_has_bytes(id));
        assert_eq!(events.stream_zero_completion(id, false), None);
    }

    #[test]
    fn consumed_guest_write_event_does_not_reappear_on_later_close() {
        let mut events = EventLoop::new();
        let id = events.create_stream();
        assert_eq!(events.stream_zero_read(id), None);
        events.stream_push_bytes(id, vec![9]);
        events.begin_guest_stream_write(id, 1);
        assert_eq!(events.stream_read_bytes(id, 1), [9]);
        assert_eq!(events.take_stream_zero_completion(id, false), Some(pack(CopyResult::Completed, 1)));
        events.stream_close_readable(id);
        assert_eq!(events.stream_zero_completion(id, false), None);
    }

    #[test]
    fn readable_drop_overrides_unobserved_write_completion() {
        let mut events = EventLoop::new();
        let id = events.create_stream();
        assert_eq!(events.stream_zero_read(id), None);
        events.begin_guest_stream_write(id, 4);
        events.stream_push_bytes(id, vec![1, 2, 3, 4]);
        assert_eq!(events.stream_read_bytes(id, 4), [1, 2, 3, 4]);
        assert_eq!(events.stream_zero_completion(id, false), Some(pack(CopyResult::Completed, 4)));
        events.stream_close_readable(id);
        assert_eq!(events.take_stream_zero_completion(id, false), Some(pack(CopyResult::Dropped, 4)));
    }

    #[test]
    fn dropped_byte_streams_release_payload_and_records_in_either_order() {
        for reader_first in [true, false] {
            let mut events = EventLoop::new();
            for _ in 0..1000 {
                let id = events.create_stream();
                events.stream_push_bytes(id, vec![41; 4096]);
                if reader_first {
                    events.stream_close_readable(id);
                    let record = &events.stream_buffers[&id];
                    assert!(record.buffer.is_empty());
                    assert_eq!(record.pending.capacity(), 0);
                    events.stream_close(id);
                } else {
                    events.stream_close(id);
                    assert!(events.stream_has_bytes(id));
                    events.stream_close_readable(id);
                }
                assert!(!events.stream_buffers.contains_key(&id));
            }
        }
    }

    #[test]
    fn typed_payload_cleanup_distinguishes_nested_component_ownership() {
        use crate::component::ValType;
        for owned in [ValType::Own("file".into()), ValType::Borrow("file".into()),
            ValType::Future(Box::new(ValType::U8)), ValType::Stream(Box::new(ValType::U8)),
            ValType::ErrorContext, ValType::Any] {
            let nested = ValType::Record(vec![("items".into(), ValType::List(Box::new(
                ValType::Option(Box::new(owned)))))]);
            assert!(!stream_payload_is_self_contained(&nested));
        }
        let plain = ValType::Record(vec![("name".into(), ValType::String),
            ("kind".into(), ValType::Variant(vec![("file".into(), None),
                ("other".into(), Some(ValType::Option(Box::new(ValType::String))))]))]);
        assert!(stream_payload_is_self_contained(&plain));
        let mut events = EventLoop::new();
        let id = events.create_stream_of(Some(plain));
        let mut value = crate::value::Object::new();
        value.properties.insert("name".into(), Value::String("payload".into()));
        value.properties.insert("kind".into(), Value::String("file".into()));
        let value = Value::Object(crate::heap::alloc(value));
        events.stream_push(id, value.clone()).unwrap();
        events.stream_close_readable(id);
        assert!(!events.stream_has_item(id));
        events.stream_push(id, value).unwrap();
        assert!(!events.stream_has_item(id));
        events.stream_close(id);
        assert!(!events.stream_buffers.contains_key(&id));
    }

    #[test]
    fn resource_payload_is_not_silently_discarded_by_plain_cleanup() {
        let mut events = EventLoop::new();
        let id = events.create_stream_of(Some(crate::component::ValType::Own("file".into())));
        events.stream_push(id, Value::I32(42)).unwrap();
        events.stream_close_readable(id);
        events.stream_close(id);
        assert!(events.stream_has_item(id));
    }

    #[test]
    fn stream_cleanup_preserves_pending_completion_until_consumed() {
        let mut events = EventLoop::new();
        let id = events.create_stream();
        assert_eq!(events.stream_zero_write(id), None);
        events.stream_close_readable(id);
        events.stream_close(id);
        assert!(events.stream_buffers.contains_key(&id));
        assert_eq!(events.take_stream_zero_completion(id, false),
            Some(pack(CopyResult::Dropped, 0)));
        assert!(!events.stream_buffers.contains_key(&id));
    }

    #[test]
    fn late_byte_producer_cannot_refill_a_dropped_reader() {
        let mut events = EventLoop::new();
        let id = events.create_stream();
        events.stream_close_readable(id);
        events.stream_push_bytes(id, vec![41; 4096]);
        assert!(!events.stream_has_bytes(id));
        events.stream_close(id);
        assert!(!events.stream_buffers.contains_key(&id));
    }

    #[test]
    fn native_stream_consumer_releases_reader_only_after_finalization() {
        let mut events = EventLoop::new();
        let id = events.create_stream();
        let future = events.create_future();
        events.set_byte_consumer(id, future, Box::new(|_| Ok(()))).unwrap();
        events.stream_push_bytes(id, vec![41; 4096]);
        assert!(events.stream_buffers.contains_key(&id));
        assert_eq!(events.future_states[&future].phase, FuturePhase::Pending);
        events.stream_close(id);
        assert_eq!(events.future_states[&future].phase, FuturePhase::Resolved);
        assert!(!events.stream_buffers.contains_key(&id));
    }

    #[test]
    fn forwarded_closed_writer_preserves_reader_payload_and_cleanup() {
        let mut events = EventLoop::new();
        let source = events.create_stream();
        let destination = events.create_stream();
        events.stream_push_bytes(source, vec![1, 2, 3]);
        events.stream_close(source);
        assert!(events.forward_stream(source, destination));
        assert!(!events.stream_buffers.contains_key(&source));
        assert_eq!(events.stream_read_bytes(destination, 3), [1, 2, 3]);
        events.stream_close_readable(destination);
        assert!(!events.stream_buffers.contains_key(&destination));
    }

    #[test]
    fn zero_read_sees_buffered_data_before_eof() {
        let mut events = EventLoop::new();
        let id = events.create_stream();
        events.stream_push_bytes(id, vec![5]);
        events.stream_close(id);
        assert_eq!(events.stream_zero_read(id), Some(CopyResult::Completed));
        assert_eq!(events.stream_read_bytes(id, 1), [5]);
        assert_eq!(events.stream_zero_read(id), Some(CopyResult::Dropped));
    }

    #[test]
    fn synchronous_stream_readers_wake_in_order_without_waking_other_streams() {
        let mut events = EventLoop::new();
        let first = events.create_stream();
        let second = events.create_stream();
        events.suspend_stream_sync_reader(first, reader(1));
        events.suspend_stream_sync_reader(second, reader(2));
        events.suspend_stream_sync_reader(first, reader(3));
        assert_eq!(events.parked_sync_copies(), 3);
        assert_eq!(events.parked_stream_ids(), [first, second]);

        events.stream_push_bytes(first, vec![42]);
        assert_eq!(events.parked_sync_copies(), 1);
        assert_eq!(events.parked_stream_ids(), [second]);
        assert_eq!(next_reader_marker(&mut events), 1);
        assert_eq!(next_reader_marker(&mut events), 3);
        assert!(events.immediate.is_empty());
        assert_eq!(events.stream_read_bytes(first, 1), [42]);

        // A reader can park again when an earlier reader consumed the data.
        events.suspend_stream_sync_reader(first, reader(4));
        assert_eq!(events.parked_sync_copies(), 2);
        events.stream_close(second);
        assert_eq!(next_reader_marker(&mut events), 2);
        events.stream_close(first);
        assert_eq!(next_reader_marker(&mut events), 4);
        assert_eq!(events.parked_sync_copies(), 0);
        assert!(events.parked_stream_ids().is_empty());
    }

    #[test]
    fn synchronous_future_readers_wake_on_resolution_or_rejection() {
        let mut events = EventLoop::new();
        let first = events.create_future();
        let second = events.create_future();
        events.suspend_future_sync_reader(first, reader(1));
        events.suspend_future_sync_reader(second, reader(2));
        events.suspend_future_sync_reader(first, reader(3));
        assert_eq!(events.parked_sync_copies(), 3);

        events.resolve_future(first, Value::I32(42));
        assert_eq!(next_reader_marker(&mut events), 1);
        assert_eq!(next_reader_marker(&mut events), 3);
        assert_eq!(events.parked_sync_copies(), 1);
        events.reject_future(second, Value::I32(7));
        assert_eq!(next_reader_marker(&mut events), 2);
        assert_eq!(events.parked_sync_copies(), 0);
        assert!(events.immediate.is_empty());

        assert_eq!(EventLoop::new().parked_sync_copies(), 0);
    }

    #[test]
    fn byte_runs_preserve_order_and_partial_reads() {
        let mut events = EventLoop::new();
        let stream = events.create_stream();
        events.stream_push_bytes(stream, vec![0, 255, 1, 2, 3]);
        events.stream_push_bytes(stream, vec![4]);
        events.stream_push_bytes(stream, vec![5, 6]);

        assert_eq!(events.stream_read_bytes(stream, 2), [0, 255]);
        assert!(events.stream_has_bytes(stream));
        assert_eq!(events.stream_read_bytes(stream, 4), [1, 2, 3, 4]);
        assert_eq!(events.stream_read_bytes(stream, 4), [5, 6]);
        assert!(!events.stream_has_bytes(stream));
    }

    #[test]
    fn closing_a_partially_read_byte_stream_keeps_remaining_bytes_ready() {
        let mut events = EventLoop::new();
        let stream = events.create_stream();
        events.stream_push_bytes(stream, vec![1, 2, 3]);
        assert_eq!(events.stream_read_bytes(stream, 1), [1]);
        events.stream_close(stream);
        assert!(events.stream_has_item(stream));
        assert!(!events.stream_is_eof(stream));
        assert_eq!(events.stream_read_bytes(stream, 4), [2, 3]);
        assert!(events.stream_is_eof(stream));
    }

    #[test]
    fn typed_stream_keeps_whole_items() {
        let mut events = EventLoop::new();
        let stream = events.create_stream_of(Some(crate::component::ValType::I32));
        events.stream_push(stream, Value::I32(257)).unwrap();
        assert!(matches!(events.stream_read_items(stream, 1).as_slice(), [Value::I32(257)]));
    }

    #[test]
    fn typed_stream_reuses_its_element_type_between_copies() {
        use crate::component::ValType;

        let nested = ValType::List(Box::new(ValType::Option(Box::new(ValType::I64))));
        let mut events = EventLoop::new();
        let stream = events.create_stream_of(Some(nested.clone()));
        let first = events.stream_elem(stream).unwrap();
        let second = events.stream_elem(stream).unwrap();
        assert!(std::sync::Arc::ptr_eq(&first, &second));
        assert_eq!(*first, nested);
        let other = events.create_stream_of(Some(nested));
        assert!(!std::sync::Arc::ptr_eq(
            &first,
            &events.stream_elem(other).unwrap()
        ));
    }

    #[test]
    fn draining_multiple_byte_runs_does_not_stop_at_the_first() {
        let mut events = EventLoop::new();
        let stream = events.create_stream();
        events.stream_push_bytes(stream, vec![1, 2]);
        events.stream_push_bytes(stream, vec![3, 4]);
        assert_eq!(events.stream_read_bytes(stream, usize::MAX), [1, 2, 3, 4]);
    }

    #[test]
    fn full_drain_after_partial_read_keeps_only_remaining_bytes() {
        let mut events = EventLoop::new();
        let stream = events.create_stream();
        events.stream_push_bytes(stream, vec![1, 2, 3]);
        events.stream_push_bytes(stream, vec![4, 5]);
        assert_eq!(events.stream_read_bytes(stream, 2), [1, 2]);
        assert_eq!(events.stream_read_bytes(stream, usize::MAX), [3, 4, 5]);
        assert!(!events.stream_has_bytes(stream));
    }

    #[test]
    fn small_reads_cross_large_byte_run_boundaries() {
        let mut events = EventLoop::new();
        let stream = events.create_stream();
        let first: Vec<u8> = (0..65536).map(|index| index as u8).collect();
        let second: Vec<u8> = (0..65536).map(|index| (index >> 3) as u8).collect();
        events.stream_push_bytes(stream, first.clone());
        events.stream_push_bytes(stream, second.clone());
        events.stream_close(stream);

        let mut actual = Vec::new();
        while !events.stream_is_eof(stream) {
            actual.extend(events.stream_read_bytes(stream, 64));
        }
        assert_eq!(actual.len(), first.len() + second.len());
        assert_eq!(&actual[..first.len()], first);
        assert_eq!(&actual[first.len()..], second);
    }

    #[test]
    #[ignore = "local byte-stream drain diagnostic"]
    fn bench_large_byte_run_small_reads() {
        use std::time::Instant;

        let chunks: Vec<Vec<u8>> = (0..4)
            .map(|_| (0..65536).map(|index| index as u8).collect())
            .collect();
        for _ in 0..3 {
            let mut old_pending = Vec::new();
            let old_start = Instant::now();
            for chunk in &chunks {
                old_pending.extend_from_slice(chunk);
                while !old_pending.is_empty() {
                    std::hint::black_box(old_pending.drain(..64).collect::<Vec<_>>());
                }
            }
            let old_elapsed = old_start.elapsed();

            let mut events = EventLoop::new();
            let stream = events.create_stream();
            for chunk in &chunks {
                events.stream_push_bytes(stream, chunk.clone());
            }
            let cursor_start = Instant::now();
            for _ in 0..4096 {
                std::hint::black_box(events.stream_read_bytes(stream, 64));
            }
            let cursor_elapsed = cursor_start.elapsed();
            assert!(!events.stream_has_bytes(stream));
            println!("front drain: {old_elapsed:?}, cursor: {cursor_elapsed:?}");
        }
    }
}
