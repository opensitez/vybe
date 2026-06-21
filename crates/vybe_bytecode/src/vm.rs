use std::cell::RefCell;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use std::sync::{Arc, Mutex, Weak as ArcWeak};

use crate::chunk::Chunk;
use crate::error::VMError;
use crate::event_loop::EventLoop;
use crate::module_record::{ExportEntry, ModuleRecord};
use crate::opcode::Op;
use crate::shared_memory::SharedMemory;
use crate::value::{Object, ObjectKind, Upvalue, Value};

pub(crate) const MAX_FRAMES: usize = 256;
pub(crate) const MAX_STACK: usize = 65536;

/// Result of VM execution — may complete or suspend for async.
pub enum ExecResult {
    /// Execution completed with a value.
    Done(Value),
    /// Execution suspended — waiting for host/runtime resolution.
    Suspended { kind: SuspensionKind, id: u64 },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SuspensionKind {
    Await,
    Jspi,
    Future,
    StreamRead,
}

/// Restricted context passed to host functions.
/// Provides only the capabilities a host function needs:
/// - Invoke VM callbacks (for LINQ, event handlers, etc.)
/// - Access linear memory (for WASI filesystem, network, etc.)
/// - Queue microtasks/timers through the event loop
///
/// Does NOT expose: globals, stack, frames, bytecode, type registry.
/// This matches the WASM security model (Wasmtime Caller<State>).
pub struct HostContext<'a> {
    /// Invoke a VM function reference with arguments.
    /// This is the ONLY way host functions can call back into the VM.
    invoker: Option<&'a mut dyn FnMut(&Value, &[Value]) -> Value>,
    /// Linear memory access (WASM MVP memory[0]).
    pub memory: Option<&'a mut [u8]>,
    /// Event loop reference for queuing microtasks and timers.
    /// Cloned from VM.event_loop — valid for the lifetime of the host call.
    event_loop: Option<Rc<RefCell<EventLoop>>>,
    /// Raw pointer to VM.last_exception — set by THROW when no handler matches.
    /// Null when no VM is attached (HostContext::empty()).
    last_exception_slot: *mut Option<Value>,
    /// Raw pointer to VM globals for host-managed JS receiver binding.
    /// Null when no VM is attached (HostContext::empty()).
    globals_slot: *mut HashMap<String, Value>,
    /// Raw pointer to VM.stack for closing escaped upvalues in timer callbacks.
    /// Null when no VM is attached (HostContext::empty()).
    #[allow(dead_code)]
    stack_slot: *const Vec<Value>,
    /// Raw pointer to the CM3 handle table, so host functions receiving a
    /// canon `stream<u8>` / `future<T>` i32 handle (CanonicalABI lowering)
    /// can resolve it to the EventLoop stream/future id.
    /// Null when no VM is attached (HostContext::empty()).
    handle_table_slot: *const crate::handle_table::HandleTable,
}

// SAFETY: HostContext is always created and used on the VM's owning thread.
// The raw pointer to last_exception_slot is valid for the duration of the host
// function call (same scope as the invoker lifetime bound by 'a).
unsafe impl Send for HostContext<'_> {}

impl<'a> HostContext<'a> {
    /// Call a VM function reference from host code.
    /// Returns Value::Null if no invoker is available.
    pub fn invoke(&mut self, func_ref: &Value, args: &[Value]) -> Value {
        if let Some(ref mut invoker) = self.invoker {
            invoker(func_ref, args)
        } else {
            Value::Null
        }
    }

    /// Like invoke, but captures any exception thrown by the callback.
    /// Returns Ok(value) on normal return or Err(thrown_value) on throw.
    pub fn try_invoke(&mut self, func_ref: &Value, args: &[Value]) -> Result<Value, Value> {
        // Clear any stale exception before the call.
        unsafe {
            if !self.last_exception_slot.is_null() {
                *self.last_exception_slot = None;
            }
        }
        let result = self.invoke(func_ref, args);
        unsafe {
            if !self.last_exception_slot.is_null() {
                if let Some(exc) = (*self.last_exception_slot).take() {
                    return Err(exc);
                }
            }
        }
        Ok(result)
    }

    /// Raise a VM exception from host code so the surrounding JS/VM
    /// try/catch path observes it like a bytecode THROW.
    pub fn throw_value(&mut self, value: Value) {
        unsafe {
            if !self.last_exception_slot.is_null() {
                *self.last_exception_slot = Some(value);
            }
        }
    }

    /// Read the current JS receiver binding (`__js_this`) from globals.
    /// Returns `Undefined` when no binding exists.
    pub fn current_js_this(&self) -> Value {
        unsafe {
            if self.globals_slot.is_null() {
                Value::Undefined
            } else {
                (*self.globals_slot)
                    .get("__js_this")
                    .cloned()
                    .unwrap_or(Value::Undefined)
            }
        }
    }

    /// Update the current JS receiver binding (`__js_this`) in globals.
    pub fn set_js_this(&mut self, value: Value) {
        unsafe {
            if !self.globals_slot.is_null() {
                (*self.globals_slot).insert("__js_this".into(), value);
            }
        }
    }

    /// Queue a microtask (Promise reaction) to run after the current task.
    /// ECMA-262 §27.2.1.3 EnqueueJob("PromiseJobs", ...).
    pub fn queue_microtask(&mut self, callback: Value, value: Value) {
        if let Some(ref el) = self.event_loop {
            el.borrow_mut().queue_microtask(callback, value);
        }
    }

    /// Queue a timer macrotask and return its cancellable ID.
    /// HTML Living Standard §8.7 setTimeout semantics.
    pub fn queue_timer(&mut self, callback: Value, delay_ms: f64) -> u64 {
        if let Some(ref el) = self.event_loop {
            el.borrow_mut().queue_timer_id(callback, delay_ms)
        } else {
            0
        }
    }

    /// Cancel a previously scheduled timer by ID. Returns true if found.
    pub fn cancel_timer(&mut self, id: u64) -> bool {
        if let Some(ref el) = self.event_loop {
            el.borrow_mut().cancel_timer(id)
        } else {
            false
        }
    }

    /// Generate a unique promise ID.
    pub fn next_promise_id(&mut self) -> u64 {
        if let Some(ref el) = self.event_loop {
            el.borrow_mut().next_promise_id()
        } else {
            0
        }
    }

    /// Resolve a suspended promise fiber and queue it in the microtask queue.
    pub fn resolve_promise(&mut self, promise_id: u64, value: Value) {
        if let Some(ref el) = self.event_loop {
            let mut el_mut = el.borrow_mut();
            if let Some(fiber) = el_mut.resolve_promise(promise_id, value) {
                el_mut
                    .microtasks
                    .push_back(crate::event_loop::Task::ResumeFiber(fiber));
            }
        }
    }

    /// Queue a rejected-promise fiber resumption — the fiber will throw the reason.
    pub fn reject_promise(&mut self, promise_id: u64, reason: Value) {
        if let Some(ref el) = self.event_loop {
            let mut el_mut = el.borrow_mut();
            if let Some(fiber) = el_mut.reject_promise(promise_id, reason) {
                el_mut
                    .microtasks
                    .push_back(crate::event_loop::Task::ResumeFiber(fiber));
            }
        }
    }

    // ── CM3 / WASI 0.3 async ────────────────────────────────────────────────

    /// Create a future. Returns the future Value (for guest code) and its ID (for host resolve).
    pub fn create_future(&mut self) -> (Value, u64) {
        use crate::value::{Object, ObjectKind};
        if let Some(ref el) = self.event_loop {
            let id = el.borrow_mut().create_future();
            let obj = Object {
                properties: std::collections::HashMap::new(),
                kind: ObjectKind::Future { id },
                type_id: 0,
                fields: Vec::new(),
            };
            let val = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(obj)));
            (val, id)
        } else {
            (Value::Null, 0)
        }
    }

    /// Resolve a future — wakes the suspended fiber (if any) with the value.
    pub fn resolve_future(&mut self, future_id: u64, value: Value) {
        if let Some(ref el) = self.event_loop {
            let mut el_mut = el.borrow_mut();
            if let Some(fiber) = el_mut.resolve_future(future_id, value) {
                el_mut
                    .microtasks
                    .push_back(crate::event_loop::Task::ResumeFiber(fiber));
            }
        }
    }

    /// Reject a future — wakes the suspended fiber (if any) with an exception.
    pub fn reject_future(&mut self, future_id: u64, reason: Value) {
        if let Some(ref el) = self.event_loop {
            let mut el_mut = el.borrow_mut();
            if let Some(fiber) = el_mut.reject_future(future_id, reason) {
                el_mut
                    .microtasks
                    .push_back(crate::event_loop::Task::ResumeFiber(fiber));
            }
        }
    }

    /// Create a stream. Returns the stream Value (for guest code) and its ID (for host push/close).
    pub fn create_stream(&mut self) -> (Value, u64) {
        use crate::value::{Object, ObjectKind};
        if let Some(ref el) = self.event_loop {
            let id = el.borrow_mut().create_stream();
            let obj = Object {
                properties: std::collections::HashMap::new(),
                kind: ObjectKind::Stream { id },
                type_id: 0,
                fields: Vec::new(),
            };
            let val = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(obj)));
            (val, id)
        } else {
            (Value::Null, 0)
        }
    }

    /// Push one item to a stream. Wakes a waiting reader fiber if present.
    pub fn stream_push(&mut self, stream_id: u64, item: Value) {
        if let Some(ref el) = self.event_loop {
            let mut el_mut = el.borrow_mut();
            if let Some(fiber) = el_mut.stream_push(stream_id, item) {
                el_mut
                    .microtasks
                    .push_back(crate::event_loop::Task::ResumeFiber(fiber));
            }
        }
    }

    /// Close a stream (signal EOF). Wakes a waiting reader fiber if present.
    pub fn stream_close(&mut self, stream_id: u64) {
        if let Some(ref el) = self.event_loop {
            let mut el_mut = el.borrow_mut();
            if let Some(fiber) = el_mut.stream_close(stream_id) {
                el_mut
                    .microtasks
                    .push_back(crate::event_loop::Task::ResumeFiber(fiber));
            }
        }
    }

    /// Synchronously drain all buffered items from a `stream<u8>` Value into bytes.
    /// Used by host functions that receive a guest stream and need to read all buffered data.
    /// Items are converted: I32 → single byte, String → UTF-8 bytes, Array<I32> → byte sequence.
    /// Returns empty Vec if the value is not a Stream or the stream has no buffered data.
    pub fn stream_drain(&mut self, stream_val: &Value) -> Vec<u8> {
        use crate::value::ObjectKind;
        let stream_id = if let Value::Object(obj) = stream_val {
            let o = obj.lock().unwrap();
            if let ObjectKind::Stream { id } = o.kind {
                id
            } else {
                return Vec::new();
            }
        } else if let Value::I32(handle) = stream_val {
            // CM3 canonical lowering: a `stream<u8>` crosses the boundary as
            // an i32 readable-end handle (CanonicalABI §HandleTable). Resolve
            // it so spec-shaped imports like wasi:cli/stdout.write-via-stream
            // work when called from guest bytecode.
            match self.resolve_readable_stream_handle(*handle as u32) {
                Some(id) => id,
                None => return Vec::new(),
            }
        } else {
            return Vec::new();
        };
        let mut out: Vec<u8> = Vec::new();
        if let Some(ref el) = self.event_loop {
            let mut el_mut = el.borrow_mut();
            if let Some(rec) = el_mut.stream_buffers.get_mut(&stream_id) {
                while let Some(item) = rec.buffer.pop_front() {
                    match &item {
                        Value::I32(b) => out.push(*b as u8),
                        Value::String(s) => out.extend_from_slice(s.as_bytes()),
                        Value::Object(inner_obj) => {
                            let o = inner_obj.lock().unwrap();
                            if let ObjectKind::Array(arr) = &o.kind {
                                for v in arr {
                                    if let Value::I32(b) = v {
                                        out.push(*b as u8);
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
        }
        out
    }

    /// Resolve a CM3 readable-stream-end handle (i32 from the handle table)
    /// to its EventLoop stream id. None when the handle is absent, of the
    /// wrong kind, or no VM is attached.
    fn resolve_readable_stream_handle(&self, handle: u32) -> Option<u64> {
        unsafe {
            if self.handle_table_slot.is_null() {
                return None;
            }
            match (*self.handle_table_slot).get(handle) {
                Some(crate::handle_table::HandleEntry::ReadableStreamEnd(id)) => Some(*id),
                _ => None,
            }
        }
    }

    /// Create an empty context (for host functions that don't need callbacks).
    pub fn empty() -> Self {
        HostContext {
            invoker: None,
            memory: None,
            event_loop: None,
            last_exception_slot: std::ptr::null_mut(),
            globals_slot: std::ptr::null_mut(),
            stack_slot: std::ptr::null(),
            handle_table_slot: std::ptr::null(),
        }
    }
}

/// Host function signature. Receives restricted context + args, returns a value.
/// Host function signature.
pub type HostFn = Arc<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>;

/// WASM import resolution target. An import can resolve to:
/// - A host function (provided by the embedder)
/// - A component-exported function (another module's code)
/// - A stdlib redirect (global function registered at runtime)
#[derive(Clone)]
pub enum ImportTarget {
    /// Index into VM::host_fns
    Host(usize),
    /// Chunk index + arity — calls a function defined in another component
    ChunkFn { chunk_index: usize, arity: u8 },
    /// Runtime global lookup (stdlib functions registered via globals)
    StdlibRedirect(String),
    /// JSPI suspending import (`jspi`.`await`, a `WebAssembly.Suspending`
    /// import). `await x` lowers to a `call` to this import; the VM (acting as
    /// the engine) implements the suspension itself rather than dispatching to a
    /// host fn — fulfilled → unwrap, rejected → throw, pending → suspend the
    /// fiber on the event loop until the Promise settles.
    JspiSuspend,
}

#[derive(Debug, Clone)]
pub(crate) struct CallFrame {
    pub(crate) chunk_index: usize,
    pub(crate) ip: usize,
    pub(crate) base: usize,
    pub(crate) label_base: usize,
    pub(crate) upvalues: Vec<Arc<Mutex<Upvalue>>>,
}

/// Record of a live continuation on the VM's active-continuation
/// stack. Each entry owns the continuation `Value` plus the caller's
/// pre-RESUME `Fiber` — that's what we restore on suspend.
///
/// `mode` selects between the raw RESUME protocol (caller sees just
/// the yielded value on its stack) and the iterator protocol (caller
/// sees `[value, has_more_i32]` — so a loop can check `has_more`
/// without a second API call). `SUSPEND` and normal-`RETURN`-out-of-
/// a-continuation consult this flag to decide what to push.
#[derive(Debug)]
pub struct ActiveContinuation {
    pub cont: crate::value::Value,
    pub caller_fiber: crate::fiber::Fiber,
    pub mode: ResumeMode,
    pub handlers: Vec<crate::chunk::StackSwitchHandler>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ResumeMode {
    /// Bare `RESUME` — push only the yielded value on caller's stack.
    Raw,
    /// `GEN_NEXT` iterator protocol — push `(value, has_more_i32)`.
    Iterator,
}

/// Exception handler entry — pushed by try_start, popped by try_end or catch.
#[derive(Debug, Clone)]
pub(crate) struct ExceptionHandler {
    /// Instruction pointer to jump to on catch.
    pub(crate) catch_ip: usize,
    /// Chunk index the handler was registered in.
    pub(crate) _chunk_index: usize,
    /// Stack depth when try_start was executed (for unwinding).
    pub(crate) stack_depth: usize,
    /// Call frame depth when try_start was executed.
    pub(crate) frame_depth: usize,
    /// Structured-control label-stack depth when the handler was installed.
    /// On catch, the label stack is unwound to this — otherwise a throw from
    /// inside a nested `block`/`loop`/`if` skips that block's `end`, leaking a
    /// stale label entry that corrupts `br` depths on subsequent execution
    /// (e.g. a try/catch re-entered each loop iteration).
    pub(crate) label_depth: usize,
    /// Exception tag index (0 = catch-all, N = typed catch for tag N).
    /// References chunk.exception_tags[tag] for the type name.
    pub(crate) tag: u8,
}

/// A language-agnostic bytecode virtual machine.
///
/// The VM has no built-in functions or language-specific semantics.
/// The host (compiler runtime) registers native functions via `register_host_fn`
/// and sets up globals before calling `run`.
pub struct VM {
    pub chunks: Vec<Chunk>,
    pub(crate) frames: Vec<CallFrame>,
    pub(crate) stack: Vec<Value>,
    pub globals: HashMap<String, Value>,
    pub(crate) open_upvalues: Vec<Arc<Mutex<Upvalue>>>,
    pub(crate) host_fns: Vec<HostFn>,
    /// Registry: (module, name) → index into host_fns.
    ///
    /// Flat cache of module exports; lookup-optimized shadow of
    /// `modules[*].exports[*]`. Kept for hot-path dispatch; the
    /// authoritative per-module view lives in `modules`.
    pub host_registry: HashMap<(String, String), usize>,
    /// ECMA-262 §16.2.1 Abstract Module Records, keyed by canonical
    /// specifier (`"wasi:cli/environment"`, `"ecma:math"`, etc.).
    /// Populated by `register_host_fn` and the ESM-Integration
    /// `.wasm` loader. The Linker phase of the compiler reads from
    /// here. See `esmhostplan.md` for the migration plan.
    pub modules: HashMap<String, ModuleRecord>,
    /// Import resolution table: import_index → resolved target.
    /// WASM-aligned: a single `.wasm` module has one imports section shared
    /// by every function inside. Vybe represents one module as many chunks
    /// (one per function) but they all share the same imports list, which
    /// the compiler stores on `chunks[0]` by convention.
    pub(crate) import_table: Vec<ImportTarget>,
    /// Exception handler stack (WASM exception proposal).
    pub(crate) exception_handlers: Vec<ExceptionHandler>,
    /// Event loop for async operations (shared with host functions).
    pub event_loop: Rc<RefCell<EventLoop>>,
    /// WASM GC-style type definitions with vtable method dispatch.
    pub type_registry: crate::typedef::TypeRegistry,
    /// Linear memory (WASM MVP) — byte buffer for binary data.
    /// This is memory index 0 for backward compatibility.
    pub memory: SharedMemory,
    /// Additional memories for multi-memory support.
    /// memory index 0 = self.memory, index 1+ = extra_memories[i-1].
    pub(crate) extra_memories: Vec<Vec<u8>>,
    /// Maximum page limits for memories 1+ from decoded module types.
    pub(crate) extra_memory_max_pages: Vec<Option<usize>>,
    /// Bulk-memory data segments that have been dropped.
    pub(crate) dropped_data: HashSet<u32>,
    /// Reference-types element segments that have been dropped.
    pub(crate) dropped_elems: HashSet<u32>,
    /// GC/bulk-memory data segment payloads loaded for VM execution.
    pub(crate) data_segments: Vec<Vec<u8>>,
    /// GC/reference-types element segment payloads loaded for VM execution.
    pub(crate) elem_segments: Vec<Vec<Value>>,
    /// Currently selected memory index (for load/store ops). Default 0.
    pub(crate) active_memory: usize,
    /// Function table (WASM MVP) — for call_indirect. Also accessible
    /// as table index 0 via the reference-types table ops.
    pub func_table: Vec<Value>,
    /// Additional reference-typed tables for the reference-types /
    /// multi-table proposal. Index N (N >= 1) lives at
    /// `extra_tables[N-1]`; index 0 is `func_table`. Tables are lazily
    /// created: compilers that only use table 0 never allocate here.
    pub extra_tables: Vec<Vec<Value>>,
    /// Optional per-slot value-type recorder. `Some` enables the
    /// LOCAL_SET / LOCAL_GET hooks to tally which `Value` variants
    /// flow through each local, feeding the anyref/ABI migration with
    /// concrete measurements of how much typed-slot lowering can save.
    /// Off by default — zero dispatch cost when `None`.
    pub type_recorder: Option<crate::type_recorder::TypeRecorder>,
    /// Stack-switching: continuations currently running. Each `RESUME`
    /// pushes an entry (with the caller's pre-resume fiber); `SUSPEND`
    /// pops the topmost entry, captures the runnable fiber into the
    /// continuation's `saved` slot, and restores the caller. Empty when
    /// no coroutine is active.
    pub active_continuations: Vec<ActiveContinuation>,
    /// Identity of the fiber whose frames are currently live. Stack-switching
    /// (`resume`/`suspend`/`GEN_NEXT`) swaps whole frame stacks via `save_fiber`
    /// / `resume_fiber_with`; this id travels with each fiber so a nested
    /// `execute_until` (e.g. a host `invoke_callback`) only honours its
    /// `min_depth` boundary while running on the same fiber it was entered on.
    /// A continuation resumed from inside a callback runs on a different fiber,
    /// so its internal returns can't trip the callback's (now-stale) floor.
    pub(crate) cur_fiber_id: u64,
    /// Monotonic source of fresh fiber ids for newly-started continuations.
    pub(crate) next_fiber_id: u64,
    /// Block label stack for structured control flow.
    pub label_stack: Vec<LabelEntry>,
    /// Per-chunk block tables: chunk_index → (opcode_start → BlockTargets).
    /// Lazily populated on first BLOCK/LOOP/IF/ELSE dispatch in each chunk.
    pub(crate) block_tables: HashMap<usize, HashMap<usize, BlockTargets>>,
    /// Callback invoker for host functions (cached allocation).
    pub(crate) callback_invoker: Option<Box<dyn FnMut(&Value, &[Value]) -> Value>>,
    /// Last exception value thrown by a callback that had no handler.
    /// Populated by the THROW opcode when no ExceptionHandler matches.
    /// Consumed by HostContext::try_invoke to report the thrown value.
    pub last_exception: Option<Value>,
    /// When true, enforce strict WASM isolation:
    /// - Module-scoped globals (prefixed by module name)
    /// - Per-module memory (separate linear memory per component)
    /// Default false for trusted code (shared globals for cross-language interop).
    pub strict_isolation: bool,
    /// Module prefix for current execution context (used when strict_isolation=true).
    pub module_prefix: Option<String>,
    /// CLS case alias map: lowercase → canonical casing.
    /// When a global_get fails, tries lowercase lookup in this map to find
    /// the canonical-cased name. Enables cross-language name resolution
    /// (VB lowercase → C# PascalCase, COBOL UPPERCASE → VB lowercase).
    pub case_aliases: HashMap<String, String>,
    /// Finalizer registry: maps object identity to callback.
    /// When an object's strong count reaches the weak+finalizer threshold,
    /// the callback is queued for execution.
    pub(crate) finalizers: Vec<FinalizerEntry>,
    /// Active threads spawned by thread_spawn opcode.
    /// Maps thread_id → JoinHandle that returns the serialized result.
    pub(crate) thread_handles: HashMap<i32, std::thread::JoinHandle<Vec<u8>>>,
    /// Next thread ID to assign.
    pub(crate) next_thread_id: i32,
    /// Execution trace: when true, print every opcode + stack top.
    /// Enable via `vm.set_trace(true)` or `VYBE_TRACE=1` env var.
    pub(crate) trace: bool,
    /// Optional chunk-name filter for execution trace output.
    /// When set, only matching chunks emit trace lines.
    pub(crate) trace_chunk_filter: Option<String>,
    // ── CM3 Canonical ABI (Track A) ─────────────────────────────────────────
    /// Handle table — maps i32 indices to typed component resources.
    pub handle_table: crate::handle_table::HandleTable,
    /// Active CM3 tasks (keyed by task ID). Each async export invocation creates one.
    pub cm_tasks: Vec<crate::cm_task::CMTask>,
    /// Next CM3 task ID.
    #[allow(dead_code)]
    pub(crate) next_cm_task_id: u32,
    /// Waitable set registry.
    pub waitable_sets: crate::waitable::WaitableRegistry,
    /// CM3 context slots (canon context.get/set).
    pub context_slots: Vec<Value>,
}

/// A registered finalizer for an object.
#[derive(Clone)]
pub(crate) struct FinalizerEntry {
    /// Weak reference to the target object.
    pub(crate) target: ArcWeak<Mutex<crate::value::Object>>,
    /// Callback to invoke when the object is about to be collected.
    pub(crate) callback: Value,
}

/// Entry in the structured control flow label stack.
#[derive(Debug, Clone, Copy)]
pub struct LabelEntry {
    /// Instruction offset to jump to on `br` (end of block, or start of loop).
    pub target: usize,
    /// True if this is a loop (continue jumps to start), false if block/if (break jumps to end).
    pub is_loop: bool,
    /// Number of stack values the label carries when branched to.
    pub result_arity: u8,
    /// Value-stack height at label entry. Branches restore this height while
    /// preserving the top `result_arity` values.
    pub stack_height: usize,
}

/// Pre-scanned jump targets for one BLOCK / LOOP / IF / ELSE opcode.
/// Keyed by the opcode's position (first prefix byte) in chunk.code.
#[derive(Debug, Clone, Copy)]
pub struct BlockTargets {
    /// For IF: position of the matching ELSE opcode (None if no else branch).
    /// For ELSE: None.
    /// For BLOCK/LOOP: None.
    pub else_ip: Option<usize>,
    /// Position of the matching END opcode.
    pub end_ip: usize,
}

impl VM {
    /// Immutable borrow of the table at `tableidx`. Index 0 maps to
    /// `func_table`; indexes 1.. map to `extra_tables`.
    pub(crate) fn table_ref(&self, idx: usize) -> Option<&Vec<Value>> {
        if idx == 0 {
            Some(&self.func_table)
        } else {
            self.extra_tables.get(idx - 1)
        }
    }
    /// Mutable borrow of the table at `tableidx`.
    pub(crate) fn table_mut(&mut self, idx: usize) -> Option<&mut Vec<Value>> {
        if idx == 0 {
            Some(&mut self.func_table)
        } else {
            self.extra_tables.get_mut(idx - 1)
        }
    }

    /// Turn on per-slot value-type recording for the next `run`.
    /// Passing `false` disables and discards any existing recorder.
    pub fn record_types(&mut self, enabled: bool) {
        self.type_recorder = if enabled {
            Some(crate::type_recorder::TypeRecorder::new())
        } else {
            None
        };
    }

    /// Take ownership of the current type recorder (leaving `None` in
    /// place). Useful for producing a report after a run without
    /// holding a mutable borrow of the VM for the whole analysis.
    pub fn take_type_record(&mut self) -> Option<crate::type_recorder::TypeRecorder> {
        self.type_recorder.take()
    }

    pub fn new() -> Self {
        VM {
            chunks: Vec::new(),
            frames: Vec::new(),
            stack: Vec::with_capacity(256),
            globals: HashMap::new(),
            open_upvalues: Vec::new(),
            host_fns: Vec::new(),
            host_registry: HashMap::new(),
            modules: HashMap::new(),
            import_table: Vec::<ImportTarget>::new(),
            exception_handlers: Vec::new(),
            event_loop: Rc::new(RefCell::new(EventLoop::new())),
            type_registry: crate::typedef::TypeRegistry::new(),
            memory: SharedMemory::default(),
            extra_memories: Vec::new(),
            extra_memory_max_pages: Vec::new(),
            dropped_data: HashSet::new(),
            dropped_elems: HashSet::new(),
            data_segments: Vec::new(),
            elem_segments: Vec::new(),
            active_memory: 0,
            func_table: Vec::new(),
            extra_tables: Vec::new(),
            type_recorder: None,
            active_continuations: Vec::new(),
            cur_fiber_id: 0,
            next_fiber_id: 1,
            label_stack: Vec::new(),
            block_tables: HashMap::new(),
            callback_invoker: None,
            last_exception: None,
            strict_isolation: false,
            module_prefix: None,
            case_aliases: HashMap::new(),
            finalizers: Vec::new(),
            thread_handles: HashMap::new(),
            next_thread_id: 1,
            trace: std::env::var("VYBE_TRACE").map_or(false, |v| v == "1" || v == "true"),
            trace_chunk_filter: std::env::var("VYBE_TRACE_CHUNK").ok(),
            handle_table: crate::handle_table::HandleTable::new(),
            cm_tasks: Vec::new(),
            next_cm_task_id: 1,
            waitable_sets: crate::waitable::WaitableRegistry::new(),
            context_slots: Vec::new(),
        }
    }

    /// Enable or disable execution tracing. When enabled, every opcode
    /// execution prints the chunk name, offset, opcode, and stack top.
    /// Can also be enabled via `VYBE_TRACE=1` environment variable.
    pub fn set_trace(&mut self, enabled: bool) {
        self.trace = enabled;
    }

    /// Restrict execution trace output to a specific chunk name.
    pub fn set_trace_chunk_filter(&mut self, chunk_name: Option<String>) {
        self.trace_chunk_filter = chunk_name;
    }

    pub fn set_data_segment(&mut self, index: usize, bytes: Vec<u8>) {
        if self.data_segments.len() <= index {
            self.data_segments.resize_with(index + 1, Vec::new);
        }
        self.data_segments[index] = bytes;
    }

    pub fn set_elem_segment(&mut self, index: usize, values: Vec<Value>) {
        if self.elem_segments.len() <= index {
            self.elem_segments.resize_with(index + 1, Vec::new);
        }
        self.elem_segments[index] = values;
    }

    /// Capture the current call stack for error reporting.
    pub fn capture_call_stack(&self) -> Vec<crate::error::StackFrame> {
        self.frames
            .iter()
            .rev()
            .map(|f| {
                let chunk = &self.chunks[f.chunk_index];
                let line = chunk.get_line(f.ip.saturating_sub(1));
                crate::error::StackFrame {
                    chunk_name: chunk.name.clone(),
                    offset: f.ip,
                    line,
                }
            })
            .collect()
    }

    /// Dump disassembled bytecode for all chunks. Useful for debugging
    /// test failures — call after `compile()` to see what was emitted.
    /// Returns a formatted string, one chunk per section.
    pub fn dump_bytecode(&self) -> String {
        let mut out = String::new();
        for (i, chunk) in self.chunks.iter().enumerate() {
            out.push_str(&format!("\n── Chunk {} ──\n", i));
            out.push_str(&crate::debug::disassemble(chunk));
        }
        out
    }

    /// Dump bytecode for a specific chunk by index.
    pub fn dump_chunk(&self, index: usize) -> String {
        if index < self.chunks.len() {
            crate::debug::disassemble(&self.chunks[index])
        } else {
            format!("Chunk {} not found (have {})", index, self.chunks.len())
        }
    }

    /// Evaluate a constant expression (Extended Const Expressions).
    /// Used for global initialization at load time.
    pub(crate) fn eval_const_expr(&self, expr: &crate::chunk::ConstExpr) -> Value {
        use crate::chunk::ConstExpr;
        match expr {
            ConstExpr::Value(v) => v.clone(),
            ConstExpr::GlobalGet(name) => self.globals.get(name).cloned().unwrap_or(Value::Null),
            ConstExpr::Add(left, right) => {
                let l = self.eval_const_expr(left);
                let r = self.eval_const_expr(right);
                match (&l, &r) {
                    (Value::I32(a), Value::I32(b)) => Value::I32(a.wrapping_add(*b)),
                    (Value::I64(a), Value::I64(b)) => Value::I64(a.wrapping_add(*b)),
                    (Value::F64(a), Value::F64(b)) => Value::F64(a + b),
                    _ => Value::F64(l.as_f64() + r.as_f64()),
                }
            }
            ConstExpr::Mul(left, right) => {
                let l = self.eval_const_expr(left);
                let r = self.eval_const_expr(right);
                match (&l, &r) {
                    (Value::I32(a), Value::I32(b)) => Value::I32(a.wrapping_mul(*b)),
                    (Value::I64(a), Value::I64(b)) => Value::I64(a.wrapping_mul(*b)),
                    (Value::F64(a), Value::F64(b)) => Value::F64(a * b),
                    _ => Value::F64(l.as_f64() * r.as_f64()),
                }
            }
            ConstExpr::RefFunc(chunk_idx) => {
                if *chunk_idx < self.chunks.len() {
                    let chunk = &self.chunks[*chunk_idx];
                    let func = crate::value::Function {
                        name: Some(chunk.name.clone()),
                        arity: chunk.arity,
                        chunk_index: *chunk_idx,
                        upvalues: Vec::new(),
                    };
                    let mut obj = Object::new();
                    obj.kind = ObjectKind::Function(func);
                    Value::Object(Arc::new(Mutex::new(obj)))
                } else {
                    Value::Null
                }
            }
        }
    }

    /// Get the size (in bytes) of a memory by spec memory index.
    pub(crate) fn mem_len(&self, memidx: usize) -> usize {
        if memidx == 0 {
            self.memory.len()
        } else {
            self.extra_memories
                .get(memidx - 1)
                .map_or(0, |mem| mem.len())
        }
    }

    fn instantiate_declared_memories(
        &mut self,
        min_pages: &[u64],
        max_pages: &[Option<u64>],
    ) -> Result<(), crate::VMError> {
        for (idx, pages) in min_pages.iter().copied().enumerate() {
            let bytes_u64 = pages
                .checked_mul(65536)
                .ok_or_else(|| crate::VMError::new("memory declaration size overflow"))?;
            let bytes = usize::try_from(bytes_u64)
                .map_err(|_| crate::VMError::new("memory declaration size out of range"))?;
            let max = max_pages
                .get(idx)
                .copied()
                .flatten()
                .map(usize::try_from)
                .transpose()
                .map_err(|_| crate::VMError::new("memory maximum out of range"))?;
            if idx == 0 {
                self.memory.set_max_pages(max);
                if self.memory.len() < bytes {
                    self.memory.resize(bytes, 0);
                }
            } else {
                let extra_idx = idx - 1;
                if self.extra_memories.len() <= extra_idx {
                    self.extra_memories.resize_with(extra_idx + 1, Vec::new);
                }
                if self.extra_memory_max_pages.len() <= extra_idx {
                    self.extra_memory_max_pages.resize(extra_idx + 1, None);
                }
                self.extra_memory_max_pages[extra_idx] = max;
                if self.extra_memories[extra_idx].len() < bytes {
                    self.extra_memories[extra_idx].resize(bytes, 0);
                }
            }
        }
        Ok(())
    }

    fn instantiate_declared_tables(&mut self, min_sizes: &[u64]) -> Result<(), crate::VMError> {
        for (idx, size) in min_sizes.iter().copied().enumerate() {
            let size = usize::try_from(size)
                .map_err(|_| crate::VMError::new("table declaration size out of range"))?;
            if idx == 0 {
                if self.func_table.len() < size {
                    self.func_table.resize(size, Value::Null);
                }
            } else {
                let extra_idx = idx - 1;
                if self.extra_tables.len() <= extra_idx {
                    self.extra_tables.resize_with(extra_idx + 1, Vec::new);
                }
                if self.extra_tables[extra_idx].len() < size {
                    self.extra_tables[extra_idx].resize(size, Value::Null);
                }
            }
        }
        Ok(())
    }

    /// Grow a memory by spec memory index. Returns old page count or `usize::MAX` on failure.
    pub(crate) fn mem_grow(&mut self, memidx: usize, pages: usize) -> usize {
        if memidx == 0 {
            self.memory.grow(pages)
        } else {
            let idx = memidx - 1;
            if idx >= self.extra_memories.len() {
                self.extra_memories.resize_with(idx + 1, Vec::new);
            }
            if idx >= self.extra_memory_max_pages.len() {
                self.extra_memory_max_pages.resize(idx + 1, None);
            }
            let mem = &mut self.extra_memories[idx];
            let old_pages = mem.len() / 65536;
            let Some(new_pages) = old_pages.checked_add(pages) else {
                return usize::MAX;
            };
            if let Some(max_pages) = self.extra_memory_max_pages[idx] {
                if new_pages > max_pages {
                    return usize::MAX;
                }
            }
            if new_pages > 65536 {
                return usize::MAX;
            }
            let Some(new_len) = new_pages.checked_mul(65536) else {
                return usize::MAX;
            };
            mem.resize(new_len, 0);
            old_pages
        }
    }

    pub(crate) fn read_memory_bytes(
        &self,
        memidx: usize,
        addr: usize,
        size: usize,
    ) -> Result<Vec<u8>, crate::VMError> {
        if memidx == 0 {
            self.memory.with_buffer(|buf| {
                if addr.saturating_add(size) > buf.len() {
                    Err(crate::VMError::new(format!(
                        "trap: memory access out of bounds: addr={} size={} limit={}",
                        addr,
                        size,
                        buf.len()
                    )))
                } else {
                    Ok(buf[addr..addr + size].to_vec())
                }
            })
        } else {
            let mem = self.extra_mem(memidx);
            if addr.saturating_add(size) > mem.len() {
                Err(crate::VMError::new(format!(
                    "trap: memory access out of bounds: addr={} size={} limit={}",
                    addr,
                    size,
                    mem.len()
                )))
            } else {
                Ok(mem[addr..addr + size].to_vec())
            }
        }
    }

    pub(crate) fn write_memory_bytes(
        &mut self,
        memidx: usize,
        addr: usize,
        bytes: &[u8],
    ) -> Result<(), crate::VMError> {
        if memidx == 0 {
            self.memory.with_buffer_mut(|buf| {
                if addr.saturating_add(bytes.len()) > buf.len() {
                    Err(crate::VMError::new(format!(
                        "trap: memory access out of bounds: addr={} size={} limit={}",
                        addr,
                        bytes.len(),
                        buf.len()
                    )))
                } else {
                    buf[addr..addr + bytes.len()].copy_from_slice(bytes);
                    Ok(())
                }
            })
        } else {
            let mem = self.extra_mem_mut(memidx);
            if addr.saturating_add(bytes.len()) > mem.len() {
                Err(crate::VMError::new(format!(
                    "trap: memory access out of bounds: addr={} size={} limit={}",
                    addr,
                    bytes.len(),
                    mem.len()
                )))
            } else {
                mem[addr..addr + bytes.len()].copy_from_slice(bytes);
                Ok(())
            }
        }
    }

    pub(crate) fn branch_to_label(&mut self, depth: usize, entry: LabelEntry) {
        if let Some(frame) = self.frames.last_mut() {
            frame.ip = entry.target;
        }

        let arity = entry.result_arity as usize;
        let keep = if arity == 0 {
            Vec::new()
        } else {
            let split = self.stack.len().saturating_sub(arity);
            self.stack.split_off(split)
        };
        self.stack.truncate(entry.stack_height);
        self.stack.extend(keep);

        let len = self.label_stack.len();
        if entry.is_loop {
            self.label_stack.truncate(len - depth);
        } else {
            self.label_stack.truncate(len - depth - 1);
        }
    }

    /// Get a reference to a specific extra memory by index (index > 0 only).
    pub(crate) fn extra_mem(&self, idx: usize) -> &[u8] {
        if idx == 0 || idx - 1 >= self.extra_memories.len() {
            &[]
        } else {
            &self.extra_memories[idx - 1]
        }
    }

    /// Get a mutable reference to a specific extra memory by index (index > 0 only).
    pub(crate) fn extra_mem_mut(&mut self, idx: usize) -> &mut Vec<u8> {
        let i = idx - 1;
        if i >= self.extra_memories.len() {
            self.extra_memories.resize_with(i + 1, Vec::new);
        }
        &mut self.extra_memories[i]
    }

    /// Run any pending finalizers for objects whose strong count has dropped.
    /// Returns collected callbacks that should be invoked by the caller.
    pub fn collect_dead_finalizers(&mut self) -> Vec<Value> {
        let mut callbacks = Vec::new();
        let mut i = 0;
        while i < self.finalizers.len() {
            if self.finalizers[i].target.strong_count() == 0 {
                let entry = self.finalizers.remove(i);
                callbacks.push(entry.callback);
            } else {
                i += 1;
            }
        }
        callbacks
    }

    /// Register a host function with a (module, name) pair.
    /// Also adds it to the function table for call_indirect dispatch,
    /// and records the export in the per-module `ModuleRecord` so the
    /// Linker (ESM host-import resolver) can see it.
    pub fn register_host_fn(
        &mut self,
        module: &str,
        name: &str,
        f: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
    ) {
        let idx = self.host_fns.len();
        self.host_fns.push(Arc::from(f));
        self.host_registry
            .insert((module.to_string(), name.to_string()), idx);
        // Add to function table — func_table index == host_fns index for host functions
        while self.func_table.len() <= idx {
            self.func_table.push(Value::Null);
        }
        // Store as a lightweight marker — call_indirect will recognize host fn indices
        let mut obj = Object::new();
        obj.kind = ObjectKind::HostFunction(idx);
        self.func_table[idx] = Value::Object(Arc::new(Mutex::new(obj)));

        // Mirror the registration into the Module Records registry.
        // First registration under a given specifier auto-creates a
        // Synthetic ModuleRecord; subsequent registrations add exports.
        // `host_registry` remains the fast lookup path; `modules` is
        // the spec-aligned per-module view.
        self.insert_host_module_export(module, name, ExportEntry::Function { idx });
    }

    /// Authoritative host function export view used by both the ESM linker
    /// and the Component linker. This derives from Module Records rather than
    /// the legacy flat cache so both link paths observe the same synthetic
    /// modules and exports.
    pub fn iter_host_function_exports(&self) -> impl Iterator<Item = (String, String, usize)> + '_ {
        self.modules.iter().flat_map(move |(module, record)| {
            record.exports.keys().filter_map(move |name| {
                self.resolve_host_function_index(module, name)
                    .map(|idx| (module.clone(), name.clone(), idx))
            })
        })
    }

    /// Register an immutable host value in the authoritative Module Records
    /// registry so ESM validation and host-module adapters can observe it.
    pub fn register_host_value(&mut self, module: &str, name: &str, value: Value) {
        self.insert_host_module_export(module, name, ExportEntry::Value(value));
    }

    /// Register a host class type export backed by the shared TypeRegistry.
    pub fn register_host_class_export(&mut self, module: &str, name: &str, type_id: usize) {
        self.insert_host_module_export(module, name, ExportEntry::Class { type_id });
    }

    /// Register a host resource type export backed by the shared TypeRegistry.
    pub fn register_host_resource_type_export(&mut self, module: &str, name: &str, type_id: usize) {
        self.insert_host_module_export(module, name, ExportEntry::ResourceType { type_id });
    }

    /// Resolve a host function import through the authoritative Module Records,
    /// following any Indirect re-exports until a concrete host function is found.
    pub fn resolve_host_function_index(&self, module: &str, name: &str) -> Option<usize> {
        let mut visited = Vec::new();
        match self.resolve_host_export(module, name, &mut visited)? {
            ExportEntry::Function { idx } => Some(*idx),
            _ => None,
        }
    }

    /// Resolve host type exports (Class / ResourceType) through Module Records.
    pub fn iter_host_type_exports(
        &self,
    ) -> impl Iterator<Item = (String, String, crate::TypeDef)> + '_ {
        self.modules.iter().flat_map(move |(module, record)| {
            record.exports.keys().filter_map(move |name| {
                self.resolve_host_type_export(module, name)
                    .map(|typedef| (module.clone(), name.clone(), typedef))
            })
        })
    }

    fn insert_host_module_export(&mut self, module: &str, name: &str, export: ExportEntry) {
        self.insert_module_export(module, name, export.clone());
        if let Some((alias_module, alias_name)) = Self::canonical_subinterface_alias(module, name) {
            self.insert_module_export(&alias_module, &alias_name, export);
        }
    }

    fn insert_module_export(&mut self, module: &str, name: &str, export: ExportEntry) {
        let record = self
            .modules
            .entry(module.to_string())
            .or_insert_with(|| ModuleRecord::new_synthetic(module));
        record.exports.insert(name.to_string(), export);
    }

    fn canonical_subinterface_alias(module: &str, name: &str) -> Option<(String, String)> {
        if name.starts_with('[') {
            return None;
        }
        let (path, leaf) = name.rsplit_once('.')?;
        if path.is_empty() || leaf.is_empty() {
            return None;
        }
        let alias_module = format!("{}/{}", module, path.replace('.', "/"));
        Some((alias_module, leaf.to_string()))
    }

    fn resolve_host_export<'a>(
        &'a self,
        module: &str,
        name: &str,
        visited: &mut Vec<(String, String)>,
    ) -> Option<&'a ExportEntry> {
        let key = (module.to_string(), name.to_string());
        if visited.contains(&key) {
            return None;
        }
        visited.push(key);

        let record = self.modules.get(module)?;
        match record.exports.get(name)? {
            ExportEntry::Indirect {
                from,
                name: target_name,
            } => self.resolve_host_export(from, target_name, visited),
            export => Some(export),
        }
    }

    fn resolve_host_type_export(&self, module: &str, name: &str) -> Option<crate::TypeDef> {
        let mut visited = Vec::new();
        match self.resolve_host_export(module, name, &mut visited)? {
            ExportEntry::Class { type_id } | ExportEntry::ResourceType { type_id } => {
                self.type_registry.get(*type_id).cloned()
            }
            _ => None,
        }
    }

    /// Create a HostContext with callback capability for host functions.
    pub(crate) fn make_host_context(&mut self) -> HostContext<'_> {
        // We can't pass &mut self into the closure directly due to borrow rules.
        // Instead, we pass raw pointers — this is safe because the HostContext
        // lifetime is strictly scoped within the host function call.
        let vm_ptr = self as *mut VM;
        // Clone the Rc so host functions can queue microtasks/timers without
        // holding a mutable borrow of the VM.
        let el = self.event_loop.clone();
        // Raw pointer to last_exception — safe: valid for host call duration.
        let exc_ptr = &mut self.last_exception as *mut Option<Value>;
        let globals_ptr = &mut self.globals as *mut HashMap<String, Value>;
        HostContext {
            invoker: Some(unsafe {
                // SAFETY: vm_ptr is valid for the duration of the host function call.
                let vm_ref: &mut VM = &mut *vm_ptr;
                vm_ref.get_invoker()
            }),
            memory: None,
            event_loop: Some(el),
            last_exception_slot: exc_ptr,
            globals_slot: globals_ptr,
            stack_slot: &self.stack as *const Vec<Value>,
            handle_table_slot: &self.handle_table as *const crate::handle_table::HandleTable,
        }
    }

    /// Close open upvalues in a lambda value that escapes the current stack frame.
    /// When a closure is stored in a macrotask queue (setTimeout), it will run in
    /// a fresh execution context. Any `Open(slot)` upvalue referencing the current
    /// stack must be converted to `Closed(value)` so the slot index remains valid.
    #[allow(dead_code)]
    pub(crate) fn close_escaped_upvalues(&self, val: &Value) {
        use crate::value::ObjectKind;
        use crate::value::UpvalueLocation;
        if let Value::Object(obj) = val {
            let o = obj.lock().unwrap();
            if let ObjectKind::Function(ref func) = o.kind {
                for uv in &func.upvalues {
                    let mut u = uv.lock().unwrap();
                    if let UpvalueLocation::Open(slot) = u.location {
                        let captured = self.stack.get(slot).cloned().unwrap_or(Value::Null);
                        u.location = UpvalueLocation::Closed(captured);
                    }
                }
            }
        }
    }

    /// Get a mutable reference to the invoker closure.
    pub(crate) fn get_invoker(&mut self) -> &mut dyn FnMut(&Value, &[Value]) -> Value {
        // This is stored as a field to avoid repeated allocation
        if self.callback_invoker.is_none() {
            let vm_ptr = self as *mut VM;
            self.callback_invoker = Some(Box::new(move |func_ref: &Value, args: &[Value]| {
                // SAFETY: vm_ptr is valid during host function execution
                let vm = unsafe { &mut *vm_ptr };
                vm.invoke_callback(func_ref, args)
            }));
        }
        self.callback_invoker.as_mut().unwrap().as_mut()
    }

    /// Invoke a VM function reference from host code.
    /// This is the WASM-compliant callback mechanism: host functions
    /// can call exported/internal VM functions during execution.
    ///
    /// Usage from a host function:
    ///   let result = vm.invoke_callback(&predicate, &[element]);
    pub fn invoke_callback(&mut self, func_ref: &Value, args: &[Value]) -> Value {
        let saved_frame_depth = self.frames.len();
        // Save the stack height so we can restore it after the callback returns,
        // giving the callback an isolated value stack (WASM call-frame semantics).
        let saved_stack_len = self.stack.len();

        // Push function ref + args onto stack
        self.stack.push(func_ref.clone());
        for arg in args {
            self.stack.push(arg.clone());
        }

        // Call the function (pushes a new frame for compiled fns; inline for host fns)
        if self.call_value(args.len()).is_err() {
            self.stack.truncate(saved_stack_len);
            return Value::Null;
        }

        // Host functions run inline — no frame was pushed, result is already on the stack.
        // Calling execute_until would re-enter the dispatch loop at the wrong IP.
        if self.frames.len() == saved_frame_depth {
            let result = self.stack.pop().unwrap_or(Value::Null);
            self.stack.truncate(saved_stack_len);
            return result;
        }

        // Execute until the callback frame returns, then restore the stack to its
        // pre-call height so the caller's expression stack is not polluted.
        let result = match self.execute_until(saved_frame_depth + 1) {
            Ok(val) => val,
            Err(_) => {
                while self.frames.len() > saved_frame_depth {
                    self.frames.pop();
                }
                Value::Null
            }
        };
        self.stack.truncate(saved_stack_len);
        result
    }

    /// Convert a value to its string representation.
    /// For objects with a `toString` method, invokes it; otherwise falls back to Display.
    /// Used for string concatenation and template literals (JS coercion semantics).
    pub fn value_to_string(&mut self, value: &Value) -> String {
        if let Value::Object(obj) = value {
            // Look for toString or valueOf method on the object
            let to_str_fn = {
                let o = obj.lock().unwrap();
                o.properties.get("toString").cloned()
            };
            if let Some(fn_val) = to_str_fn {
                if matches!(fn_val, Value::Object(_)) {
                    let result = self.invoke_callback(&fn_val, &[value.clone()]);
                    // If result is a string-like primitive, use it; otherwise fall through
                    if !matches!(result, Value::Null | Value::Undefined) {
                        return format!("{}", result);
                    }
                }
            }
        }
        format!("{}", value)
    }

    /// Get a type_id by name from the TypeRegistry.
    pub fn get_type_id(&self, name: &str) -> usize {
        self.type_registry.get_id(name).unwrap_or(0)
    }

    /// Load chunks and execute the script chunk (first in the new set).
    /// Appends to existing chunks so cross-language calls work (functions reference chunk indices).
    /// Resolves the import table against registered host functions.
    /// Run linked components with WASM Component Model isolation.
    /// Each component gets its own global namespace (prefixed).
    /// Cross-component communication happens ONLY through declared exports/imports.
    /// Type metadata is shared read-only for inheritance.
    pub fn run_components(
        &mut self,
        link_result: &crate::component::LinkResult,
        components: &[crate::component::Component],
    ) -> Result<Value, VMError> {
        // Load all chunks
        let base_offset = self.chunks.len();
        self.chunks.extend(link_result.chunks.clone());

        // Load shared type table (read-only cross-module)
        for ((_, type_name), typedef) in &link_result.type_exports {
            // Types are shared — they enable cross-language inheritance
            let _ = (type_name, typedef);
        }

        // Load type tables from all chunks
        for chunk in &link_result.chunks {
            if !chunk.types.is_empty() {
                self.type_registry.load_type_table(&chunk.types);
            }
        }

        // Run each component's script chunk with module isolation
        let saved_isolation = self.strict_isolation;
        let saved_prefix = self.module_prefix.clone();
        self.strict_isolation = true;

        for (i, comp) in components.iter().enumerate() {
            let _chunk_offset = link_result.component_offsets[i] + base_offset;

            // Set module prefix for global isolation
            self.module_prefix = Some(comp.name.clone());

            // Inject imported function references into this module's scope
            for (iface, func_name) in &comp.imports {
                let key = (iface.clone(), func_name.clone());
                if let Some(export_impl) = link_result.exports.get(&key) {
                    let func_val = match export_impl {
                        crate::component::ExportImpl::ChunkFn(ci) => {
                            let adjusted_ci = ci + base_offset;
                            let chunk = &self.chunks[adjusted_ci];
                            let func = crate::value::Function {
                                name: Some(func_name.clone()),
                                arity: chunk.arity,
                                chunk_index: adjusted_ci,
                                upvalues: Vec::new(),
                            };
                            let obj = crate::value::Object {
                                properties: std::collections::HashMap::new(),
                                kind: crate::value::ObjectKind::Function(func),
                                type_id: 0,
                                fields: Vec::new(),
                            };
                            Value::Object(Arc::new(Mutex::new(obj)))
                        }
                        crate::component::ExportImpl::HostFn(idx) => {
                            let mut obj = crate::value::Object::new();
                            obj.kind = crate::value::ObjectKind::HostFunction(*idx);
                            Value::Object(Arc::new(Mutex::new(obj)))
                        }
                    };
                    // Store in module-scoped globals
                    let global_key = format!("{}::{}", comp.name, func_name.to_lowercase());
                    self.globals.insert(global_key, func_val.clone());
                    // Also store without prefix so the module's code can find it
                    // (the module emits global_get "func_name", which gets prefixed by strict_isolation)
                    let unprefixed = func_name.to_lowercase();
                    self.globals
                        .insert(format!("{}::{}", comp.name, unprefixed), func_val);
                }
            }

            // Also inject exported functions from OTHER modules that this module imports
            // by making them available under the importing module's prefix
            for other_comp in components {
                if other_comp.name == comp.name {
                    continue;
                }
                for ((_, func_name), export_impl) in &other_comp.exports {
                    let func_val = match export_impl {
                        crate::component::ExportImpl::ChunkFn(ci) => {
                            let other_offset = link_result.component_offsets[components
                                .iter()
                                .position(|c| c.name == other_comp.name)
                                .unwrap()]
                                + base_offset;
                            let adjusted_ci = ci + other_offset;
                            if adjusted_ci >= self.chunks.len() {
                                continue;
                            }
                            let chunk = &self.chunks[adjusted_ci];
                            let func = crate::value::Function {
                                name: Some(func_name.clone()),
                                arity: chunk.arity,
                                chunk_index: adjusted_ci,
                                upvalues: Vec::new(),
                            };
                            let obj = crate::value::Object {
                                properties: std::collections::HashMap::new(),
                                kind: crate::value::ObjectKind::Function(func),
                                type_id: 0,
                                fields: Vec::new(),
                            };
                            Value::Object(Arc::new(Mutex::new(obj)))
                        }
                        crate::component::ExportImpl::HostFn(idx) => {
                            let mut obj = crate::value::Object::new();
                            obj.kind = crate::value::ObjectKind::HostFunction(*idx);
                            Value::Object(Arc::new(Mutex::new(obj)))
                        }
                    };
                    // Available to this module via its prefix
                    let key = format!("{}::{}", comp.name, func_name.to_lowercase());
                    self.globals.entry(key).or_insert(func_val.clone());
                    // Also store class constructors so inheritance works
                    // (the child module does global_get "ClassName" which becomes "child::classname")
                    let ctor_key = format!("{}::{}", comp.name, func_name.to_lowercase());
                    self.globals.entry(ctor_key).or_insert(func_val);
                }
            }

            // Run this component's chunks through the standard run() path
            // which handles import resolution, type tables, and frame setup
            let comp_chunks = comp.chunks.clone();
            match self.run(comp_chunks) {
                Ok(_) => {}
                Err(e) => {
                    self.strict_isolation = saved_isolation;
                    self.module_prefix = saved_prefix;
                    return Err(e);
                }
            }
        }

        // Restore isolation state
        self.strict_isolation = saved_isolation;
        self.module_prefix = saved_prefix;

        Ok(Value::Null)
    }

    /// Run linked chunks with a pre-resolved import table from the Linker.
    /// Used for bootstrap: Linker resolves imports at link time, VM just loads them.
    pub fn run_linked(
        &mut self,
        chunks: Vec<Chunk>,
        resolved_imports: Vec<ImportTarget>,
    ) -> Result<Value, VMError> {
        if chunks.is_empty() {
            return Ok(Value::Null);
        }
        let script_idx = self.chunks.len();
        // Offset ref_func indices
        let mut adjusted = chunks;
        if script_idx > 0 {
            for chunk in &mut adjusted {
                let code = &mut chunk.code;
                let mut ip = 0;
                while ip < code.len() {
                    if ip + 1 >= code.len() {
                        break;
                    }
                    let prefix = code[ip];
                    let sub = code[ip + 1];
                    if let Some(op) = Op::decode(prefix, sub as u16) {
                        if op == Op::REF_FUNC {
                            if ip + 4 < code.len() {
                                let old_idx = ((code[ip + 2] as u16) << 8) | (code[ip + 3] as u16);
                                let new_idx = old_idx + script_idx as u16;
                                code[ip + 2] = (new_idx >> 8) as u8;
                                code[ip + 3] = (new_idx & 0xff) as u8;
                            }
                            ip += 4 + 1; // 2 opcode + 2 func_idx + 1 uv_count
                            if ip - 1 < code.len() {
                                let uv_count = code[ip - 1] as usize;
                                ip += uv_count * 2;
                            }
                            continue;
                        }
                        ip += 2; // all opcodes are 2 bytes
                        let fmt = op.operand_format();
                        ip += fmt.size_in(code, ip);
                    } else {
                        ip += 2; // skip unknown 2-byte opcode
                    }
                }
            }
        }
        self.chunks.extend(adjusted);

        // Use pre-resolved import table, adjusting ChunkFn indices by offset
        self.import_table.clear();
        for target in resolved_imports {
            match target {
                ImportTarget::ChunkFn { chunk_index, arity } => {
                    self.import_table.push(ImportTarget::ChunkFn {
                        chunk_index: chunk_index + script_idx,
                        arity,
                    });
                }
                ImportTarget::Host(idx) => {
                    self.import_table.push(ImportTarget::Host(idx));
                }
                ImportTarget::StdlibRedirect(name) => {
                    // Try to resolve against host registry first
                    let key_candidates = [("wasi:cli".to_string(), name.clone())];
                    let mut resolved = false;
                    for key in &key_candidates {
                        if let Some(idx) = self.resolve_host_function_index(&key.0, &key.1) {
                            self.import_table.push(ImportTarget::Host(idx));
                            resolved = true;
                            break;
                        }
                    }
                    if !resolved {
                        self.import_table.push(ImportTarget::StdlibRedirect(name));
                    }
                }
                ImportTarget::JspiSuspend => {
                    self.import_table.push(ImportTarget::JspiSuspend);
                }
            }
        }

        // Load type table
        {
            let types = self.chunks[script_idx].types.clone();
            if !types.is_empty() {
                let adjusted_types: Vec<_> = types
                    .iter()
                    .map(|t| {
                        let mut entry = t.clone();
                        entry.methods = t
                            .methods
                            .iter()
                            .map(|(name, idx)| (name.clone(), idx + script_idx))
                            .collect();
                        if let Some(ci) = entry.constructor_chunk {
                            entry.constructor_chunk = Some(ci + script_idx);
                        }
                        entry
                    })
                    .collect();
                self.type_registry.load_type_table(&adjusted_types);
            }
        }

        // Execute
        self.frames.push(CallFrame {
            chunk_index: script_idx,
            ip: 0,
            base: self.stack.len(),
            label_base: self.label_stack.len(),
            upvalues: Vec::new(),
        });
        self.stack.resize(
            self.stack.len() + self.chunks[script_idx].local_count as usize,
            Value::Null,
        );
        self.execute()
    }

    pub fn run(&mut self, chunks: Vec<Chunk>) -> Result<Value, VMError> {
        if chunks.is_empty() {
            return Ok(Value::Null);
        }
        // Preserve globals / chunks / type registry across runs, but discard
        // per-execution state (stale frames/stack from a previous run would
        // leave the next run's HALT stuck on an inner-frame path).
        self.close_upvalues(0);
        self.stack.clear();
        self.frames.clear();
        let script_idx = self.chunks.len(); // offset for new chunks
        // Offset ref_func indices in the new chunks so they point to correct positions
        let mut adjusted = chunks;
        if script_idx > 0 {
            for chunk in &mut adjusted {
                let code = &mut chunk.code;
                let mut ip = 0;
                while ip < code.len() {
                    if ip + 1 >= code.len() {
                        break;
                    }
                    let prefix = code[ip];
                    let sub = code[ip + 1];
                    if let Some(op) = Op::decode(prefix, sub as u16) {
                        if op == Op::REF_FUNC {
                            if ip + 4 < code.len() {
                                let old_idx = ((code[ip + 2] as u16) << 8) | (code[ip + 3] as u16);
                                let new_idx = old_idx + script_idx as u16;
                                code[ip + 2] = (new_idx >> 8) as u8;
                                code[ip + 3] = (new_idx & 0xff) as u8;
                            }
                            ip += 4 + 1;
                            if ip - 1 < code.len() {
                                let uv_count = code[ip - 1] as usize;
                                ip += uv_count * 2;
                            }
                            continue;
                        }
                        ip += 2; // all opcodes are 2 bytes
                        let fmt = op.operand_format();
                        ip += fmt.size_in(code, ip);
                    } else {
                        ip += 2; // skip unknown 2-byte opcode
                    }
                }
            }
        }
        self.chunks.extend(adjusted);

        let declared_memories = self.chunks[script_idx].memory_min_pages.clone();
        let declared_memory_maxes = self.chunks[script_idx].memory_max_pages.clone();
        self.instantiate_declared_memories(&declared_memories, &declared_memory_maxes)?;
        let declared_tables = self.chunks[script_idx].table_min_sizes.clone();
        self.instantiate_declared_tables(&declared_tables)?;
        let data_segments = self.chunks[script_idx].data_segments.clone();
        if !data_segments.is_empty() {
            self.data_segments = data_segments;
            self.dropped_data.clear();
        }
        let elem_segments = self.chunks[script_idx].elem_segments.clone();
        if !elem_segments.is_empty() {
            self.elem_segments = elem_segments
                .into_iter()
                .map(|segment| {
                    segment
                        .into_iter()
                        .map(|value| match value {
                            Value::I32(func_idx) if func_idx >= 0 => {
                                let defined_func_base = if self.chunks[script_idx].name
                                    == "<script>"
                                    && self.chunks.len() > script_idx + 1
                                {
                                    script_idx + 1
                                } else {
                                    script_idx
                                };
                                let chunk_idx = defined_func_base + func_idx as usize;
                                if chunk_idx < self.chunks.len() {
                                    let chunk = &self.chunks[chunk_idx];
                                    let func = crate::value::Function {
                                        name: Some(chunk.name.clone()),
                                        arity: chunk.arity,
                                        chunk_index: chunk_idx,
                                        upvalues: Vec::new(),
                                    };
                                    let mut obj = Object::new();
                                    obj.kind = ObjectKind::Function(func);
                                    Value::Object(Arc::new(Mutex::new(obj)))
                                } else {
                                    Value::Null
                                }
                            }
                            other => other,
                        })
                        .collect()
                })
                .collect();
            self.dropped_elems.clear();
        }
        let active_data_segments = self.chunks[script_idx].active_data_segments.clone();
        for init in active_data_segments {
            let bytes = self
                .data_segments
                .get(init.data_index as usize)
                .ok_or_else(|| crate::VMError::new("active data segment payload missing"))?
                .clone();
            let offset = usize::try_from(init.offset)
                .map_err(|_| crate::VMError::new("active data segment offset out of range"))?;
            self.write_memory_bytes(init.memory_index as usize, offset, &bytes)?;
        }
        let active_elem_segments = self.chunks[script_idx].active_elem_segments.clone();
        for init in active_elem_segments {
            let values = self
                .elem_segments
                .get(init.elem_index as usize)
                .ok_or_else(|| crate::VMError::new("active element segment payload missing"))?
                .clone();
            let offset = usize::try_from(init.offset)
                .map_err(|_| crate::VMError::new("active element segment offset out of range"))?;
            let table = self
                .table_mut(init.table_index as usize)
                .ok_or_else(|| crate::VMError::new("active element segment table missing"))?;
            if offset.saturating_add(values.len()) > table.len() {
                return Err(crate::VMError::new(
                    "trap: active element segment out of bounds",
                ));
            }
            table[offset..offset + values.len()].clone_from_slice(&values);
        }

        // Resolve imports for ALL new chunks (not just script chunk).
        // Each chunk has its own import list. We build one unified import table
        // by scanning all chunks and mapping their import indices to host functions.
        // The trick: all chunks compiled by the same compiler share the same import list
        // (imports are added to chunks[0] by all compilers). For multi-module programs,
        // different modules may have different imports. We resolve the union.
        self.import_table.clear();
        for (_i, import) in self.chunks[script_idx].imports.iter().enumerate() {
            // 0. JSPI suspending import (`await`): handled by the VM itself.
            if import.module == "jspi" && import.name == "await" {
                self.import_table.push(ImportTarget::JspiSuspend);
                continue;
            }
            // 1. Try host function registry (exact module:name match)
            if let Some(idx) = self.resolve_host_function_index(&import.module, &import.name) {
                self.import_table.push(ImportTarget::Host(idx));
                continue;
            }
            // 2. Wildcard module "*" — resolve from globals (cross-language or same-language)
            if import.module == "*" {
                // Check lowercase and original case
                let candidates = [import.name.clone(), import.name.to_lowercase()];
                let found = candidates
                    .iter()
                    .find(|g| self.globals.contains_key(g.as_str()));
                if let Some(global_name) = found {
                    self.import_table
                        .push(ImportTarget::StdlibRedirect(global_name.clone()));
                    continue;
                }
            }
            // 3. Check for stdlib global
            let candidates = [
                format!("__vybe_{}", import.name),
                format!("__vybe_{}", import.name.to_lowercase()),
            ];
            let found = candidates
                .iter()
                .find(|g| self.globals.contains_key(g.as_str()));
            if let Some(global_name) = found {
                self.import_table
                    .push(ImportTarget::StdlibRedirect(global_name.clone()));
            } else {
                return Err(VMError::new(format!(
                    "Unresolved import: \"{}\" \"{}\"",
                    import.module, import.name
                )));
            }
        }

        // Load type table from the script chunk (WASM GC type section).
        // Registers user-defined class types and their vtable methods.
        // Sets __tid_<name> globals so constructors can stamp type_id.
        {
            let types = self.chunks[script_idx].types.clone();
            if !types.is_empty() {
                // Adjust chunk indices in methods (same offset as ref_func)
                let adjusted_types: Vec<_> = types
                    .iter()
                    .map(|t| {
                        let mut entry = t.clone();
                        entry.methods = t
                            .methods
                            .iter()
                            .map(|(name, idx)| (name.clone(), idx + script_idx))
                            .collect();
                        entry
                    })
                    .collect();
                self.type_registry.load_type_table(&adjusted_types);
                // Set __tid_<name> globals for each registered type. The
                // name is what the compiler canonicalised on registration —
                // the constructor will look up the same `__tid_<canon>`
                // global, so don't re-transform here.
                for entry in &adjusted_types {
                    if let Some(tid) = self.type_registry.get_id(&entry.name) {
                        let key = format!("__tid_{}", entry.name);
                        self.globals.insert(key, Value::I32(tid as i32));
                    }
                }
            }
        }

        // Evaluate global initializers (Extended Const Expressions).
        // These are computed at load time before any code runs.
        //
        // Polyfill rule: if the host has already populated a global before
        // run() is called (e.g. Vybe overriding `__vybe_pow` with native f64
        // pow), do NOT clobber it with the bundled stdlib fallback. The
        // stdlib chunks installed via global_inits are the portable
        // implementation that runs on standard WASM VMs; on Vybe they yield
        // to the optimized native version.
        {
            let inits = self.chunks[script_idx].global_inits.clone();
            for gi in &inits {
                if matches!(self.globals.get(&gi.name), None | Some(Value::Null)) {
                    let val = self.eval_const_expr(&gi.init);
                    self.globals.insert(gi.name.clone(), val);
                }
            }
        }

        self.frames.push(CallFrame {
            chunk_index: script_idx,
            ip: 0,
            base: 0,
            label_base: self.label_stack.len(),
            upvalues: Vec::new(),
        });

        let local_count = self.chunks[script_idx].local_count as usize;
        for _ in 0..local_count {
            self.stack.push(Value::Null);
        }

        // Run synchronous code
        let result = self.execute_with_async()?;

        // Event loop — process async tasks until all done
        match result {
            ExecResult::Done(val) => {
                self.run_event_loop()?;
                Ok(val)
            }
            ExecResult::Suspended {
                kind: SuspensionKind::Jspi,
                id,
            } => {
                self.run_event_loop()?;
                if self.has_pending_jspi() {
                    Err(VMError::new(format!("__jspi__:{}", id)))
                } else {
                    Ok(Value::Null)
                }
            }
            ExecResult::Suspended { .. } => {
                self.run_event_loop()?;
                Ok(Value::Null)
            }
        }
    }

    /// Run the event loop until all pending tasks are processed.
    /// Used for event callbacks — the VM state (globals, chunks) is preserved.
    pub fn invoke(&mut self, callee: &Value, args: &[Value]) -> Result<Value, VMError> {
        // Close any remaining open upvalues before clearing the stack,
        // so closures retain their captured values.
        self.close_upvalues(0);
        self.stack.clear();
        self.frames.clear();

        self.push(callee.clone())?;
        for arg in args {
            self.push(arg.clone())?;
        }

        self.call_value(args.len())?;

        // HostFunction calls are handled inline by call_value (no frame pushed).
        // If no frame was pushed, the result is already on the stack.
        if self.frames.is_empty() {
            return Ok(self.stack.pop().unwrap_or(Value::Null));
        }

        self.execute()
    }

    // -- Stack --

    pub(crate) fn push(&mut self, value: Value) -> Result<(), VMError> {
        if self.stack.len() >= MAX_STACK {
            return Err(VMError::new("Stack overflow"));
        }
        self.stack.push(value);
        Ok(())
    }

    pub(crate) fn stack_floor(&self) -> usize {
        self.frames
            .last()
            .map(|frame| {
                let chunk = &self.chunks[frame.chunk_index];
                frame.base + (chunk.local_count as usize).max(chunk.arity as usize)
            })
            .unwrap_or(0)
    }

    pub(crate) fn pop(&mut self) -> Value {
        self.stack.pop().expect("stack underflow")
    }

    pub(crate) fn peek(&self, distance: usize) -> &Value {
        &self.stack[self.stack.len() - 1 - distance]
    }

    // -- Frame --

    pub(crate) fn frame(&self) -> &CallFrame {
        self.frames.last().expect("no frame")
    }

    pub(crate) fn frame_mut(&mut self) -> &mut CallFrame {
        self.frames.last_mut().expect("no frame")
    }

    pub(crate) fn read_byte(&mut self) -> u8 {
        let f = self.frame();
        let byte = self.chunks[f.chunk_index].code[f.ip];
        self.frame_mut().ip += 1;
        byte
    }

    pub(crate) fn read_u16(&mut self) -> u16 {
        let hi = self.read_byte() as u16;
        let lo = self.read_byte() as u16;
        (hi << 8) | lo
    }

    pub(crate) fn read_i16(&mut self) -> i16 {
        self.read_u16() as i16
    }

    pub(crate) fn get_constant(&self, index: u16) -> Value {
        let f = self.frame();
        self.chunks[f.chunk_index].constants[index as usize].clone()
    }

    pub(crate) fn resolve_chunk_import(
        &self,
        chunk_index: usize,
        import_idx: usize,
    ) -> Result<Option<ImportTarget>, VMError> {
        let Some(import) = self
            .chunks
            .get(chunk_index)
            .and_then(|chunk| chunk.imports.get(import_idx))
        else {
            return Ok(None);
        };

        if import.module == "jspi" && import.name == "await" {
            return Ok(Some(ImportTarget::JspiSuspend));
        }

        if let Some(idx) = self.resolve_host_function_index(&import.module, &import.name) {
            return Ok(Some(ImportTarget::Host(idx)));
        }

        if import.module == "*" {
            let candidates = [import.name.clone(), import.name.to_lowercase()];
            if let Some(global_name) = candidates
                .iter()
                .find(|name| self.globals.contains_key(name.as_str()))
            {
                return Ok(Some(ImportTarget::StdlibRedirect(global_name.clone())));
            }
        }

        let candidates = [
            format!("__vybe_{}", import.name),
            format!("__vybe_{}", import.name.to_lowercase()),
        ];
        if let Some(global_name) = candidates
            .iter()
            .find(|name| self.globals.contains_key(name.as_str()))
        {
            return Ok(Some(ImportTarget::StdlibRedirect(global_name.clone())));
        }

        Err(VMError::new(format!(
            "Unresolved import: \"{}\" \"{}\"",
            import.module, import.name
        )))
    }

    pub(crate) fn constant_str(&self, index: u16) -> String {
        match &self.get_constant(index) {
            Value::String(s) => s.to_string(),
            v => format!("{}", v),
        }
    }

    // -- SIMD helpers --
    pub(crate) fn execute_with_async(&mut self) -> Result<ExecResult, VMError> {
        match self.execute() {
            Ok(val) => Ok(ExecResult::Done(val)),
            Err(e) if e.message.starts_with("__await__:") => {
                // Await suspension — extract promise ID
                let id: u64 = e.message["__await__:".len()..].parse().unwrap_or(0);
                Ok(ExecResult::Suspended {
                    kind: SuspensionKind::Await,
                    id,
                })
            }
            Err(e) if e.message.starts_with("__jspi__:") => {
                let id: u64 = e.message["__jspi__:".len()..].parse().unwrap_or(0);
                Ok(ExecResult::Suspended {
                    kind: SuspensionKind::Jspi,
                    id,
                })
            }
            Err(e) if e.message.starts_with("__future__:") => {
                let id: u64 = e.message["__future__:".len()..].parse().unwrap_or(0);
                Ok(ExecResult::Suspended {
                    kind: SuspensionKind::Future,
                    id,
                })
            }
            Err(e) if e.message.starts_with("__stream_read__:") => {
                let id: u64 = e.message["__stream_read__:".len()..].parse().unwrap_or(0);
                Ok(ExecResult::Suspended {
                    kind: SuspensionKind::StreamRead,
                    id,
                })
            }
            Err(e) => Err(e),
        }
    }

    pub(crate) fn execute(&mut self) -> Result<Value, VMError> {
        self.execute_until(0).map_err(|e| {
            if e.call_stack.is_empty() {
                let stack = self.capture_call_stack();
                e.with_stack(stack)
            } else {
                e
            }
        })
    }
}

pub(crate) fn dyn_truthy(v: &Value) -> bool {
    match v {
        Value::Null | Value::Undefined => false,
        Value::Bool(b) => *b,
        Value::F64(n) => *n != 0.0 && !n.is_nan(),
        Value::I32(n) => *n != 0,
        Value::I64(n) => *n != 0,
        Value::String(s) => !s.is_empty(),
        Value::Object(o) => {
            let ob = o.lock().unwrap();
            // Check __bool__ property (set by Python classes as a bool value).
            // If __bool__ is a Bool value, use it. If it's a function, we can't
            // call it from here (no VM access) — the compiler should call it
            // and set the result as a bool property.
            if let Some(Value::Bool(b)) = ob.properties.get("__bool__") {
                return *b;
            }
            // JS semantics: all objects are truthy (including empty arrays/objects).
            // Python's bool() builtin handles empty-container checks at the compiler level.
            true
        }
        Value::WeakRef(w) => w.upgrade().is_some(),
        Value::V128(b) => b.iter().any(|&x| x != 0),
        Value::Symbol(_) => true,
        Value::BigInt(n) => *n != 0,
    }
}
