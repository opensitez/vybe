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

// Heap-OOM guard for the iterative frame Vec (vybe's WASM recursion is
// iterative — `call` pushes a CallFrame and the single `run` loop continues —
// so this bounds heap growth, NOT the native OS stack). WASM imposes no
// call-depth limit (stack exhaustion is implementation-defined); ECMA-262
// §6.2.3 surfaces it as a catchable `RangeError`, which the cap does via
// `make_stack_overflow_error`. 256 was far too low for legitimate deep
// recursion; 16_384 is comparable to a JS engine's frame budget while a
// CallFrame is only tens of bytes (well under memory pressure).
pub(crate) const MAX_FRAMES: usize = 16_384;
pub(crate) const MAX_STACK: usize = 65536;

/// Result of VM execution — may complete or suspend for async.
pub enum ExecResult {
    /// Execution completed with a value.
    Done(Value),
    /// Execution suspended — waiting for host/runtime resolution.
    Suspended { kind: SuspensionKind, id: u64 } }

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SuspensionKind {
    Await,
    Jspi,
    Future,
    StreamRead }

/// Restricted context passed to host functions.
/// Provides only the capabilities a host function needs:
/// - Invoke VM callbacks (for LINQ, event handlers, etc.)
/// - Access linear memory (for WASI filesystem, network, etc.)
/// - Enqueue ready callbacks through the event loop
///
/// Does NOT expose: globals, stack, frames, bytecode, type registry.
/// This matches the WASM security model (Wasmtime Caller<State>).
pub struct HostContext<'a> {
    /// Invoke a VM function reference with arguments.
    /// This is the ONLY way host functions can call back into the VM.
    invoker: Option<&'a mut dyn FnMut(&Value, &[Value]) -> Value>,
    /// Linear memory access (WASM MVP memory[0]).
    pub memory: Option<&'a mut [u8]>,
    /// Event loop reference for enqueueing ready callbacks.
    /// Cloned from VM.event_loop — valid for the lifetime of the host call.
    event_loop: Option<Rc<RefCell<EventLoop>>>,
    /// Raw pointer to VM.last_exception — set by THROW when no handler matches.
    /// Null when no VM is attached (HostContext::empty()).
    last_exception_slot: *mut Option<Value>,
    /// Raw pointer to VM.pending_exit — set by `wasi:cli/exit` to end the run.
    /// Null when no VM is attached (HostContext::empty()).
    exit_slot: *mut bool,
    /// Raw pointer to VM.pending_exit_code — the status `wasi:cli/exit` was
    /// given. Null when no VM is attached (`HostContext::empty()`).
    exit_code_slot: *mut i32,
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
    /// Raw pointer to the VM's shared memory, for the `wasm:threads`
    /// scheduler intrinsics (`all_parked`). Null when no VM is attached.
    shared_memory_slot: *const crate::shared_memory::SharedMemory }

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

    /// Request clean termination of the current run (WASI `cli/exit`). The VM
    /// ends `run()` at the next host-call return, unwinding all frames and
    /// handing control back to the embedder — not `std::process::exit`.
    /// End the run with `code` as the process status. Sibling of
    /// [`Self::request_exit`], which is the same thing with status 0.
    pub fn request_exit_with_code(&mut self, code: i32) {
        unsafe {
            if !self.exit_code_slot.is_null() {
                *self.exit_code_slot = code;
            }
        }
        self.request_exit();
    }

    pub fn request_exit(&mut self) {
        unsafe {
            if !self.exit_slot.is_null() {
                *self.exit_slot = true;
            }
        }
    }

    /// Read a VM global by name (`Undefined` when absent). Lets host
    /// modules reach canonical per-VM anchors (e.g. `__ctor_Error`'s
    /// `prototype`, wired by a language prelude) so host-minted values
    /// stay identical to compiled ones.
    pub fn get_global(&self, name: &str) -> Value {
        unsafe {
            if self.globals_slot.is_null() {
                Value::Undefined
            } else {
                (*self.globals_slot)
                    .get(name)
                    .cloned()
                    .unwrap_or(Value::Undefined)
            }
        }
    }

    /// Write a VM global by name — counterpart of [`Self::get_global`],
    /// used by host modules that must bind calling-convention globals
    /// (e.g. `__js_new_target` around [[Construct]] dispatch).
    pub fn set_global(&mut self, name: &str, value: Value) {
        unsafe {
            if !self.globals_slot.is_null() {
                (*self.globals_slot).insert(name.to_string(), value);
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

    /// Enqueue a callback on the ready queue. What the entry MEANS is the
    /// caller's spec (`platforms/ecma` enqueues §27.2.1.3 PromiseJobs here);
    /// the VM only preserves arrival order.
    pub fn queue_ready(&mut self, callback: Value, value: Value) {
        if let Some(ref el) = self.event_loop {
            el.borrow_mut().queue_immediate(callback, value);
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

    /// Resolve a suspended promise fiber and queue its resumption as ready.
    pub fn resolve_promise(&mut self, promise_id: u64, value: Value) {
        if let Some(ref el) = self.event_loop {
            let mut el_mut = el.borrow_mut();
            if let Some(fiber) = el_mut.resolve_promise(promise_id, value) {
                el_mut
                    .immediate
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
                    .immediate
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
                properties: indexmap::IndexMap::new(),
                kind: ObjectKind::Future { id },
                type_id: 0,
                fields: Vec::new() };
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
                    .immediate
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
                    .immediate
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
                properties: indexmap::IndexMap::new(),
                kind: ObjectKind::Stream { id },
                type_id: 0,
                fields: Vec::new() };
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
                    .immediate
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
                    .immediate
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
                None => return Vec::new() }
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
                _ => None }
        }
    }

    /// Create an empty context (for host functions that don't need callbacks).
    pub fn empty() -> Self {
        HostContext {
            invoker: None,
            memory: None,
            event_loop: None,
            last_exception_slot: std::ptr::null_mut(),
            exit_slot: std::ptr::null_mut(),
            exit_code_slot: std::ptr::null_mut(),
            globals_slot: std::ptr::null_mut(),
            stack_slot: std::ptr::null(),
            handle_table_slot: std::ptr::null(),
            shared_memory_slot: std::ptr::null() }
    }

    /// From an awake host-fn caller's view: is every OTHER VM thread parked
    /// in `wait32`? False when no VM is attached.
    pub fn all_other_threads_parked(&self) -> bool {
        if self.shared_memory_slot.is_null() {
            return false;
        }
        // SAFETY: set from &self.memory in make_host_context; the VM outlives
        // the host call, same contract as the other slots.
        unsafe { (*self.shared_memory_slot).all_others_parked() }
    }
}

/// Host function signature. Receives restricted context + args, returns a value.
/// Host function signature.
pub type HostFn = Arc<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>;

/// WASM import resolution target. An import can resolve to:
/// - A host function (provided by the embedder)
/// - A component-exported function (another module's code)
/// - A stdlib redirect (global function registered at runtime)
/// - A string constant (js-string-builtins imported string)
#[derive(Clone)]
pub enum ImportTarget {
    /// Index into VM::host_fns
    Host(usize),
    /// Chunk index + arity — calls a function defined in another component
    ChunkFn { chunk_index: usize, arity: u8 },
    /// Runtime global lookup (stdlib functions registered via globals)
    StdlibRedirect(String),
    /// js-string-builtins imported string constant — returns the string value.
    StringConst(Arc<str>),
    /// JSPI suspending import (`jspi`.`await`, a `WebAssembly.Suspending`
    /// import). `await x` lowers to a `call` to this import; the VM (acting as
    /// the engine) implements the suspension itself rather than dispatching to a
    /// host fn — fulfilled → unwrap, rejected → throw, pending → suspend the
    /// fiber on the event loop until the Promise settles.
    JspiSuspend,
    /// `jspi.await_eager` — the eager-continuation await (`AsyncOp::AwaitEager`):
    /// settled antecedents continue synchronously, pending ones suspend. The
    /// instruction chose the semantics; the VM just implements both.
    JspiSuspendEager,
    /// `jspi`.`yield` — one full turn of the ready queue: save the fiber
    /// and requeue it at the BACK, so every already-queued job (Task.Run
    /// bodies, reactions) runs first. C# `Task.Yield`, the polling tick of
    /// the async channel surface. Distinct from `await`: yield NEVER
    /// continues synchronously, even under eager-await semantics.
    JspiYield,
    /// `wasi:threads`.`thread-spawn` (wasi-threads proposal:
    /// `thread-spawn(start_arg: i32) -> i32`, tid or negative error). The VM
    /// implements the import natively — exactly as wasmtime does — spawning
    /// an OS thread that invokes the module's `__wasi_thread_start` chunk
    /// with `(tid, start_arg)`. `start_arg` points at a record in shared
    /// linear memory: `{fn_table_index: i32, status_word: i32}` (the
    /// wasi-libc pthread_create pattern). No thread OPCODE exists — this
    /// import is the whole surface.
    WasiThreadSpawn }

#[derive(Debug, Clone)]
pub(crate) struct CallFrame {
    pub(crate) chunk_index: usize,
    pub(crate) ip: usize,
    pub(crate) base: usize,
    pub(crate) label_base: usize,
    pub(crate) upvalues: Vec<Arc<Mutex<Upvalue>>> }

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
    pub handlers: Vec<crate::chunk::StackSwitchHandler> }

/// Order-insensitive equality of two import tables (compile emits imports in
/// HashMap order, which varies run to run). A genuine edit changes the SET.
fn imports_equal_as_set(a: &[crate::chunk::Import], b: &[crate::chunk::Import]) -> bool {
    if a.len() != b.len() {
        return false;
    }
    let mut av: Vec<(&str, &str)> = a.iter().map(|i| (i.module.as_str(), i.name.as_str())).collect();
    let mut bv: Vec<(&str, &str)> = b.iter().map(|i| (i.module.as_str(), i.name.as_str())).collect();
    av.sort_unstable();
    bv.sort_unstable();
    av == bv
}

/// Structural equality for a chunk's exception-tag declarations (`TagDecl` has
/// no derived `PartialEq`). Used by hot reload to require identical tag layout
/// before swapping a body — so the per-chunk tag→entity maps stay valid.
fn tags_equal(a: &[crate::chunk::TagDecl], b: &[crate::chunk::TagDecl]) -> bool {
    a.len() == b.len()
        && a.iter().zip(b.iter()).all(|(x, y)| {
            x.debug_name == y.debug_name && x.arity == y.arity && x.imported == y.imported
        })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ResumeMode {
    /// Bare `RESUME` — push only the yielded value on caller's stack.
    Raw,
    /// `GEN_NEXT` iterator protocol — push `(value, has_more_i32)`.
    Iterator }

/// Spec EH catch-clause kinds (exception-handling proposal `try_table`).
pub(crate) const CATCH_KIND_CATCH: u8 = 0;
pub(crate) const CATCH_KIND_CATCH_REF: u8 = 1;
pub(crate) const CATCH_KIND_CATCH_ALL: u8 = 2;
pub(crate) const CATCH_KIND_CATCH_ALL_REF: u8 = 3;

/// A resolved tag ENTITY — spec EH tag identity. Entity 0 is always the
/// host-provided `vybe:exception` tag (arity 1: the exception object),
/// which every legacy `raise_exception_value` throw uses; chunk-local
/// declarations create fresh entities at load, imports resolve by name.
#[derive(Debug, Clone)]
pub(crate) struct TagEntity {
    pub(crate) debug_name: String,
    pub(crate) arity: u8 }

/// Exception handler entry — pushed per catch clause by `try_table`,
/// popped (as a group) by TRY_END or on catch.
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
    /// Clause kind: CATCH_KIND_* (catch / catch_ref / catch_all / catch_all_ref).
    pub(crate) kind: u8,
    /// Resolved tag ENTITY id — matching is `thrown_entity == tag_entity`,
    /// nothing else (spec: tag identity; the payload is never inspected).
    /// Unused for the catch_all kinds.
    pub(crate) tag_entity: usize,
    /// All clauses of one `try_table` share a group id, so a catch or
    /// TRY_END removes the whole table's clauses together.
    pub(crate) group: u64 }

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
    /// Resolved tag ENTITIES (spec EH): identity is the index. Entity 0 is
    /// the host-provided `vybe:exception` tag every legacy raise uses.
    pub(crate) tag_entities: Vec<TagEntity>,
    /// Per-chunk resolution: chunk tag index → entity id. Built at load:
    /// imports resolve by name to a shared entity, local declarations are
    /// fresh ("created fresh each time" — spec tag section).
    pub(crate) chunk_tag_maps: Vec<Vec<usize>>,
    /// Monotonic group id for try_table clause groups.
    pub(crate) try_group_counter: u64,
    /// Name → entity for IMPORTED tags only (spec: imports resolve by
    /// name; local declarations never enter this registry).
    pub(crate) imported_tag_registry: HashMap<String, usize>,
    /// Event loop for async operations (shared with host functions).
    pub event_loop: Rc<RefCell<EventLoop>>,
    /// The installed host scheduler (see [`crate::scheduler::Scheduler`]).
    /// `None` = the VM's own mechanism-preserving fallback loop — bare-VM
    /// tests run without any platform. Installed at plugin registration,
    /// like host functions; deliberately NOT part of snapshots.
    pub scheduler: Option<std::sync::Arc<dyn crate::scheduler::Scheduler>>,
    /// Host-owned sources of time-deferred work (see
    /// [`crate::scheduler::DeferredSource`]) — e.g. `platforms/web`'s timer
    /// wheel. Registered at plugin init; the VM polls readiness, the host
    /// owns the storage. Deliberately NOT part of snapshots.
    pub deferred_sources: Vec<std::sync::Arc<dyn crate::scheduler::DeferredSource>>,
    /// WASM GC-style type definitions with vtable method dispatch.
    pub type_registry: crate::typedef::TypeRegistry,
    /// Names of the running module's own defined types, in `chunk.types` order.
    /// Diagnostics only — nothing resolves a type by name at run time; see
    /// `module_type_ids`.
    pub(crate) module_type_names: Vec<String>,
    /// The module's type index space: slot `i` holds the **registry id** of
    /// the module's `i`-th defined type, so a `struct.new` / `array.new`
    /// immediate resolves by INDEX, as the spec addresses types. Names are
    /// bound once here at load; the running instruction never sees one.
    /// (Kept 1-based at the instruction — `0` means "no GC type", rtt `0`.)
    pub(crate) module_type_ids: Vec<usize>,
    /// Where each chunk's OWN module starts inside `module_type_ids`, parallel
    /// to `chunks`. A type immediate is relative to the module that emitted
    /// it, so several modules — components, a dynamically eval'd program —
    /// can each number their types from 1 without colliding. Without this the
    /// second module's `$1` reads the first module's type, and because
    /// `test_type` early-returns on `type_id > 0`, a WRONG rtt is worse than
    /// none: it suppresses every fallback.
    pub(crate) chunk_type_base: Vec<usize>,
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
    /// Per-memory index type: `true` = 64-bit (memory64). Populated at
    /// instantiation from the script chunk's `memory_is_64`. Read by every
    /// load/store to pick the address width (memory64 adds no new opcodes).
    pub(crate) memory_is_64: Vec<bool>,
    /// Per-table 64-bit index type (table64), populated at instantiation.
    pub(crate) table_is_64: Vec<bool>,
    /// The module's function index space (imports then defined funcs) —
    /// populated by `ref.func` and host-fn registration, read by call-by-index
    /// and `return_call_indirect`. This is NOT a WASM table: a `(table …)` is a
    /// separate funcref array (see `wasm_tables`). Keeping them apart is what
    /// stops the ~2000 registered host fns from swamping a module's table 0.
    pub func_table: Vec<Value>,
    /// Interned capture-free funcrefs, keyed by `chunk_index`. A `ref.func $f`
    /// (`REF_FUNC`) with no captured upvalues always yields the SAME object, so
    /// two tear-offs of one function are reference-identical — this is what
    /// makes `ref.eq` (and `is`/`===` on functions) a pure `Arc::ptr_eq` without
    /// any funcref-dedup extension in the comparison op. Closures WITH captures
    /// are never interned (different captures = distinct objects).
    pub(crate) funcref_cache: std::collections::HashMap<usize, Value>,
    /// WASM reference tables (reference-types / multi-table proposal), indexed
    /// directly: table N is `wasm_tables[N]`, table 0 included. Declared by
    /// `(table …)`, populated by elem segments / `table.set` / `ref.func`
    /// values, and read by `table.get`/`call_indirect`.
    pub wasm_tables: Vec<Vec<Value>>,
    /// Optional maximum element count per table (spec table limits). `table.grow`
    /// past the max returns -1; `None` = unbounded. Aligns with `wasm_tables`.
    pub(crate) wasm_table_maxes: Vec<Option<usize>>,
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
    /// JSPI promising boundary of the fiber currently running: the pending
    /// result Promise handed out when an async-function call suspended. Settled
    /// with the body's outcome when the fiber runs to completion. Travels with
    /// the fiber across save/resume (`Fiber::result_promise`).
    pub(crate) cur_fiber_result_promise: Option<Value>,
    /// Frame-depth floors of async-function calls currently on THIS fiber's
    /// stack (JSPI promising boundaries). While non-empty, a pending `await`
    /// suspends only the innermost async call's frames (bounded capture in
    /// `call_async`) instead of the whole program — the caller keeps running
    /// with a pending Promise, per JSPI `WebAssembly.promising` semantics.
    pub(crate) async_floors: Vec<usize>,
    /// `await` on an ALREADY-SETTLED promise inside an async boundary: JSPI
    /// resumes "by the event queue task runner" even when the promise is
    /// resolved, so the suspension still happens (bounded) and the resume is
    /// queued as immediately ready. This side-channel carries the settled
    /// (id, value, is_exception) from `do_await` to `call_async`, which wakes
    /// the just-registered fiber via the ready queue.
    pub(crate) pending_settled_await: Option<(u64, Value, bool)>,
    /// Completion value of the most recent fiber the event loop resumed.
    /// A top-level await suspends the script fiber; its eventual RETURN
    /// happens inside `run_event_loop`, and `run()`'s Suspended path
    /// returns this so the program's final value isn't dropped.
    pub(crate) last_fiber_completion: Option<Value>,
    /// TEMP diagnostics (VYBE_DEBUG_AC): last host import invoked.
    pub(crate) dbg_last_import: Option<String>,
    /// Frame-depth floors of the active (nested) `execute_until` loops.
    /// Exception unwinding must NOT cross a floor: a handler in a frame
    /// below the innermost floor belongs to an OUTER dispatch loop, so the
    /// raise defers (last_exception + Err) and re-raises at the outer
    /// loop's host-call site with clean stack discipline.
    pub(crate) exec_floors: Vec<usize>,
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
    /// Set by a host function (`wasi:cli/exit`) to request clean termination of
    /// the current run. Checked after each host call in CALL_IMPORT: it ends
    /// `run()` by returning control to the embedder from any frame depth — NOT a
    /// `std::process::exit`, so it never tears down the host process.
    pub pending_exit: bool,
    /// Status the guest asked to exit with — `sys.exit(3)`, `System.exit(3)`,
    /// `halt(2)`, `STOP RUN`, `exit(255)`. Every one of those accepted a status
    /// and produced 0, because `request_exit` had nowhere to put it: the slot
    /// was a `*mut bool`. Read by the embedder after `run()` returns.
    pub pending_exit_code: i32,
    /// When true, enforce strict WASM isolation:
    /// - Module-scoped globals (prefixed by module name)
    /// - Per-module memory (separate linear memory per component)
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
    /// `VYBE_DEBUG_AC=1`, read ONCE at construction. The dispatch loop's
    /// AC diagnostics must never call `env::var` per instruction: `getenv`
    /// takes libc's process-global lock, and an ungated read on every host
    /// import call made a goroutine's empty counting loop ~1000x slower
    /// under a concurrently polling main thread (measured via profiler
    /// sample — all child time in `__findenv_locked`).
    pub(crate) dbg_ac: bool,
    /// Attached step debugger (see `debugger.rs`). `None` in normal runs.
    pub(crate) debugger: Option<crate::debugger::Debugger>,
    /// Compiler-backed expression evaluator for the debugger. Installed by the
    /// shell (`vybex`) since compilation lives above this crate. Given the live
    /// VM (read-only), an expression string, and the paused frame's locals as
    /// (name, value) pairs, it compiles + evaluates the expression in an
    /// isolated mini-VM (never perturbing this VM's stack/frames) and returns
    /// the value. `None` → expression eval reports as unavailable.
    #[allow(clippy::type_complexity)]
    pub(crate) eval_hook:
        Option<Box<dyn FnMut(&VM, &str, &[(String, Value)]) -> Result<Value, String>>>,
    /// Debugger hot-reload recompiler. Installed by the shell: given the live VM
    /// (for aligned import/module resolution), re-reads + recompiles the source
    /// and returns the fresh chunk set. `apply_reload` decides what is safe to
    /// swap. `None` → hot reload reports as unavailable.
    #[allow(clippy::type_complexity)]
    pub(crate) reload_hook: Option<Box<dyn FnMut(&mut VM) -> Result<Vec<Chunk>, String>>>,
    /// Debugger event simulator. Installed by the shell (captures the live GUI
    /// state): given `(control, event)` it looks up the registered handler and
    /// invokes it through this VM, so a click/close can be fired from the
    /// debugger without an OS window. `None` → simulating reports as unavailable.
    #[allow(clippy::type_complexity)]
    pub(crate) event_fire_hook:
        Option<Box<dyn FnMut(&mut VM, &str, &str) -> Result<Value, String>>>,
    /// Single hot-path gate = `trace || debugger.is_some()`. Kept in sync by
    /// `set_trace` / `attach_debugger` / `detach_debugger` so the dispatch loop
    /// tests one bool, not two.
    pub(crate) instrumented: bool,
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
    pub context_slots: Vec<Value> }

/// A restorable post-boot baseline for [`VM::snapshot`] / [`VM::reset_to`].
///
/// Captures every script-mutable ("bucket B") field's value right after boot,
/// so a reset returns the VM to a pristine warm state without re-running the
/// prelude. Code / host-fn / type-registry fields ("bucket A") are never
/// captured — they don't change per run. The heap portion requires
/// `heap::enable_tracking()` BEFORE boot to reclaim the script generation
/// (including cycles); without it the heap restore is a no-op and only the
/// VM-owned fields reset.
///
/// Every field here is captured by VALUE from the live post-boot VM (not
/// assumed empty), so `reset_to` is a structural restore, not a per-field
/// "clear on the belief it was empty at boot" guess.
pub struct VmSnapshot {
    heap: crate::heap::HeapSnapshot,
    globals: HashMap<String, Value>,
    memory: Vec<u8>,
    extra_memories: Vec<Vec<u8>>,
    wasm_tables: Vec<Vec<Value>>,
    dropped_data: HashSet<u32>,
    dropped_elems: HashSet<u32>,
    active_memory: usize,
    handle_table: crate::handle_table::HandleTable,
    waitable_sets: crate::waitable::WaitableRegistry,
    cm_tasks: Vec<crate::cm_task::CMTask>,
    context_slots: Vec<Value>,
    try_group_counter: u64,
    cur_fiber_id: u64,
    next_fiber_id: u64,
    next_thread_id: i32,
    next_cm_task_id: u32,
    // ── Chunk-parallel structures that ACCUMULATE per `run()` (append, not
    // replace). Truncating them back to boot length on reset drops the prior
    // run's code — and, security-critically, its embedded string/data CONSTANTS
    // — so no earlier tenant's script bytes survive in a reused VM.
    chunks_len: usize,
    chunk_tag_maps_len: usize,
    tag_entities_len: usize,
    // Data-carrying "code-adjacent" fields a script run can extend: the funcref
    // index space (script closures), cross-language name aliases, and the import
    // table. Restored by value so a reset leaves them byte-identical to boot.
    func_table: Vec<Value>,
    case_aliases: HashMap<String, String>,
    import_table: Vec<ImportTarget>,
    // Per-run module payloads that `run()` overwrites only when the new script
    // HAS them (`if !empty`) — so a later run without segments would otherwise
    // inherit the prior tenant's embedded data/element bytes. Security: restore
    // to boot so no earlier script's bytes survive.
    data_segments: Vec<Vec<u8>>,
    elem_segments: Vec<Vec<Value>>,
    module_type_names: Vec<String>,
    module_type_ids: Vec<usize>,
    chunk_type_base: Vec<usize>,
    // Coupled with `tag_entities` (maps imported-tag name → index into it). Must
    // restore together: truncating tag_entities without this would leave a
    // dangling index a later lookup could read out of bounds.
    imported_tag_registry: HashMap<String, usize> }

/// A registered finalizer for an object.
#[derive(Clone)]
pub(crate) struct FinalizerEntry {
    /// Weak reference to the target object.
    pub(crate) target: ArcWeak<Mutex<crate::value::Object>>,
    /// Callback to invoke when the object is about to be collected.
    pub(crate) callback: Value }

/// Entry in the structured control flow label stack.
#[derive(Debug, Clone, Copy)]
pub struct LabelEntry {
    /// Instruction offset to jump to on `br` (end of block, or start of loop).
    pub target: usize,
    /// True if this is a loop (continue jumps to start), false if block/if (break jumps to end).
    pub is_loop: bool,
    /// True if this label closes a try_table. Normal END must also pop
    /// the active exception handler for the protected region.
    pub is_try: bool,
    /// Number of stack values the label carries when branched to.
    pub result_arity: u8,
    /// Value-stack height at label entry. Branches restore this height while
    /// preserving the top `result_arity` values.
    pub stack_height: usize }

/// Pre-scanned jump targets for one BLOCK / LOOP / IF / ELSE opcode.
/// Keyed by the opcode's position (first prefix byte) in chunk.code.
#[derive(Debug, Clone, Copy)]
pub struct BlockTargets {
    /// For IF: position of the matching ELSE opcode (None if no else branch).
    /// For ELSE: None.
    /// For BLOCK/LOOP: None.
    pub else_ip: Option<usize>,
    /// Position of the matching END opcode.
    pub end_ip: usize }

impl VM {
    /// Immutable borrow of the table at `tableidx`. Index 0 maps to
    /// WASM tables in `wasm_tables`, indexed directly (table 0 = `wasm_tables[0]`).
    pub(crate) fn table_ref(&self, idx: usize) -> Option<&Vec<Value>> {
        self.wasm_tables.get(idx)
    }
    /// Mutable borrow of the WASM table at `tableidx`.
    pub(crate) fn table_mut(&mut self, idx: usize) -> Option<&mut Vec<Value>> {
        self.wasm_tables.get_mut(idx)
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
            globals: {
                let mut g = HashMap::new();
                g.insert("undefined".to_string(), Value::Undefined);
                g
            },
            open_upvalues: Vec::new(),
            host_fns: Vec::new(),
            host_registry: HashMap::new(),
            modules: HashMap::new(),
            import_table: Vec::<ImportTarget>::new(),
            exception_handlers: Vec::new(),
            tag_entities: vec![TagEntity {
                debug_name: "vybe:exception".into(),
                arity: 1 }],
            chunk_tag_maps: Vec::new(),
            try_group_counter: 0,
            imported_tag_registry: HashMap::from([("vybe:exception".to_string(), 0usize)]),
            event_loop: Rc::new(RefCell::new(EventLoop::new())),
            scheduler: None,
            deferred_sources: Vec::new(),
            type_registry: crate::typedef::TypeRegistry::new(),
            module_type_names: Vec::new(),
            module_type_ids: Vec::new(),
            chunk_type_base: Vec::new(),
            memory: SharedMemory::default(),
            extra_memories: Vec::new(),
            extra_memory_max_pages: Vec::new(),
            dropped_data: HashSet::new(),
            dropped_elems: HashSet::new(),
            data_segments: Vec::new(),
            elem_segments: Vec::new(),
            active_memory: 0,
            memory_is_64: Vec::new(),
            table_is_64: Vec::new(),
            func_table: Vec::new(),
            funcref_cache: std::collections::HashMap::new(),
            wasm_tables: Vec::new(),
            wasm_table_maxes: Vec::new(),
            type_recorder: None,
            active_continuations: Vec::new(),
            cur_fiber_id: 0,
            cur_fiber_result_promise: None,
            async_floors: Vec::new(),
            pending_settled_await: None,
            last_fiber_completion: None,
            dbg_last_import: None,
            exec_floors: Vec::new(),
            next_fiber_id: 1,
            label_stack: Vec::new(),
            block_tables: HashMap::new(),
            callback_invoker: None,
            last_exception: None,
            pending_exit: false,
            pending_exit_code: 0,
            case_aliases: HashMap::new(),
            finalizers: Vec::new(),
            thread_handles: HashMap::new(),
            next_thread_id: 1,
            trace: std::env::var("VYBE_TRACE").map_or(false, |v| v == "1" || v == "true"),
            trace_chunk_filter: std::env::var("VYBE_TRACE_CHUNK").ok(),
            dbg_ac: std::env::var("VYBE_DEBUG_AC").is_ok(),
            debugger: None,
            eval_hook: None,
            reload_hook: None,
            event_fire_hook: None,
            instrumented: std::env::var("VYBE_TRACE").map_or(false, |v| v == "1" || v == "true"),
            handle_table: crate::handle_table::HandleTable::new(),
            cm_tasks: Vec::new(),
            next_cm_task_id: 1,
            waitable_sets: crate::waitable::WaitableRegistry::new(),
            context_slots: Vec::new() }
    }

    /// Capture the current state as a restorable warm baseline (VM hot-reset,
    /// Tier 1). Call once right after boot (register_all + prelude run) with
    /// `heap::enable_tracking()` already on. Cheap: clones `globals` + the small
    /// component/data fields and snapshots the heap registry (a shallow clone of
    /// each baseline object's contents). See `vmhotresetplan.md`.
    pub fn snapshot(&self) -> VmSnapshot {
        VmSnapshot {
            heap: crate::heap::snapshot(),
            globals: self.globals.clone(),
            memory: self.memory.with_buffer(|b| b.to_vec()),
            extra_memories: self.extra_memories.clone(),
            wasm_tables: self.wasm_tables.clone(),
            dropped_data: self.dropped_data.clone(),
            dropped_elems: self.dropped_elems.clone(),
            active_memory: self.active_memory,
            handle_table: self.handle_table.clone(),
            waitable_sets: self.waitable_sets.clone(),
            cm_tasks: self.cm_tasks.clone(),
            context_slots: self.context_slots.clone(),
            try_group_counter: self.try_group_counter,
            cur_fiber_id: self.cur_fiber_id,
            next_fiber_id: self.next_fiber_id,
            next_thread_id: self.next_thread_id,
            next_cm_task_id: self.next_cm_task_id,
            chunks_len: self.chunks.len(),
            chunk_tag_maps_len: self.chunk_tag_maps.len(),
            tag_entities_len: self.tag_entities.len(),
            func_table: self.func_table.clone(),
            case_aliases: self.case_aliases.clone(),
            import_table: self.import_table.clone(),
            data_segments: self.data_segments.clone(),
            elem_segments: self.elem_segments.clone(),
            module_type_names: self.module_type_names.clone(),
            module_type_ids: self.module_type_ids.clone(),
            chunk_type_base: self.chunk_type_base.clone(),
            imported_tag_registry: self.imported_tag_registry.clone() }
    }

    /// Restore the VM to a [`snapshot`](VM::snapshot) baseline: free the whole
    /// post-snapshot script generation (objects + cycles, via `heap::restore`),
    /// drop script-added globals / restore reassigned ones, reset wasm memory &
    /// tables to boot bytes, and clear all transient execution state. Leaves the
    /// VM byte-indistinguishable from a freshly-booted one. Code, host fns, type
    /// registry, modules, and the debugger/eval/reload hooks are untouched.
    pub fn reset_to(&mut self, snap: &VmSnapshot) {
        // 1. Heap: force-clear the script generation (breaks cycles so refcounts
        //    collapse to 0) and rewire baseline objects to their boot contents.
        //    Runs FIRST: collect_since clears contents regardless of live roots,
        //    so cycles break here; steps 2/5 then drop the roots.
        crate::heap::restore(&snap.heap);
        // 2. Globals: script-added keys vanish; reassigned baseline keys restored.
        self.globals = snap.globals.clone();
        // 3. Wasm linear memory + tables + segment-drop state → boot.
        self.memory.with_buffer_mut(|b| {
            b.clear();
            b.extend_from_slice(&snap.memory);
        });
        self.extra_memories = snap.extra_memories.clone();
        self.wasm_tables = snap.wasm_tables.clone();
        self.dropped_data = snap.dropped_data.clone();
        self.dropped_elems = snap.dropped_elems.clone();
        self.active_memory = snap.active_memory;
        // 4. Component-model data state — restored from the captured baseline,
        //    not assumed empty (handle_table can root Values).
        self.handle_table = snap.handle_table.clone();
        self.waitable_sets = snap.waitable_sets.clone();
        self.cm_tasks = snap.cm_tasks.clone();
        self.context_slots = snap.context_slots.clone();
        // 4b. Drop the prior run's appended CODE (and its embedded string/data
        //     constants — security: no earlier tenant's bytes survive) + the
        //     chunk-parallel structures that grow with it. Everything below the
        //     boot length is baseline (prelude) and stays. Other per-chunk caches
        //     keyed by index (block_tables, funcref_cache) are cleared in step 5.
        self.chunks.truncate(snap.chunks_len);
        self.chunk_tag_maps.truncate(snap.chunk_tag_maps_len);
        self.tag_entities.truncate(snap.tag_entities_len);
        // Restore code-adjacent data fields a run can extend (script funcrefs,
        // name aliases, resolved imports) to their exact boot value.
        self.func_table = snap.func_table.clone();
        self.case_aliases = snap.case_aliases.clone();
        self.import_table = snap.import_table.clone();
        // Security: restore per-run module payloads to boot (else a prior
        // script's data/element bytes or module identity could survive a reset
        // whose next script happens not to declare its own).
        self.data_segments = snap.data_segments.clone();
        self.elem_segments = snap.elem_segments.clone();
        self.module_type_names = snap.module_type_names.clone();
        self.module_type_ids = snap.module_type_ids.clone();
        self.chunk_type_base = snap.chunk_type_base.clone();
        self.imported_tag_registry = snap.imported_tag_registry.clone();
        // 5. Transient execution state — always empty between top-level runs.
        self.stack.clear();
        self.frames.clear();
        self.open_upvalues.clear();
        self.exception_handlers.clear();
        self.exec_floors.clear();
        self.async_floors.clear();
        self.label_stack.clear();
        self.active_continuations.clear();
        self.finalizers.clear();
        // Detaches (does NOT join) any threads the script spawned — acceptable
        // for reset-between-runs; a hung script thread is the embedder's concern.
        self.thread_handles.clear();
        self.funcref_cache.clear();
        self.block_tables.clear(); // code-derived cache; rebuilds lazily.
        // 6. Event loop: reset the SHARED RefCell contents in place so host fns
        //    holding an `Rc` clone see the drained loop (reassigning the Rc would
        //    desync them). Drops all queued ready work + pending fibers.
        *self.event_loop.borrow_mut() = EventLoop::new();
        // 7. Fiber / async scalars back to baseline; per-run flags cleared.
        self.cur_fiber_id = snap.cur_fiber_id;
        self.next_fiber_id = snap.next_fiber_id;
        self.cur_fiber_result_promise = None;
        self.pending_settled_await = None;
        self.last_fiber_completion = None;
        self.last_exception = None;
        self.pending_exit = false;
        self.pending_exit_code = 0;
        self.dbg_last_import = None;
        // 8. A stale callback invoker can root Values across a reset — drop it
        //    (re-installed on demand by the next host callback).
        self.callback_invoker = None;
        // 9. Counters.
        self.try_group_counter = snap.try_group_counter;
        self.next_thread_id = snap.next_thread_id;
        self.next_cm_task_id = snap.next_cm_task_id;
    }

    /// Enable or disable execution tracing. When enabled, every opcode
    /// execution prints the chunk name, offset, opcode, and stack top.
    /// Can also be enabled via `VYBE_TRACE=1` environment variable.
    pub fn set_trace(&mut self, enabled: bool) {
        self.trace = enabled;
        self.instrumented = self.trace || self.debugger.is_some();
    }

    /// Attach a step debugger. The dispatch loop will call into it at every
    /// instruction boundary until detached. `cmd_rx`/`evt_tx` are the VM-side
    /// ends of the channels whose other ends the transport (in `vybex`) holds.
    pub fn attach_debugger(
        &mut self,
        cmd_rx: std::sync::mpsc::Receiver<crate::debugger::DebugRequest>,
        evt_tx: std::sync::mpsc::Sender<crate::debugger::DebugEvent>,
        pause_on_entry: bool,
    ) {
        self.debugger = Some(crate::debugger::Debugger::new(
            cmd_rx,
            evt_tx,
            pause_on_entry,
        ));
        self.instrumented = true;
    }

    /// Detach the debugger and (unless tracing) leave the hot path uninstrumented.
    pub fn detach_debugger(&mut self) {
        self.debugger = None;
        self.instrumented = self.trace;
    }

    /// Install the debugger's compiler-backed expression evaluator (see the
    /// `eval_hook` field). Called by the shell once, before running.
    #[allow(clippy::type_complexity)]
    pub fn set_eval_hook(
        &mut self,
        hook: Box<dyn FnMut(&VM, &str, &[(String, Value)]) -> Result<Value, String>>,
    ) {
        self.eval_hook = Some(hook);
    }

    /// Evaluate a debugger expression against the live VM with the paused
    /// frame's `locals` in scope. Faithful semantics (real compiler, isolated
    /// mini-VM); this VM's execution state is never touched. Errors if no hook
    /// is installed.
    pub fn debug_eval(&mut self, expr: &str, locals: &[(String, Value)]) -> Result<Value, String> {
        let mut hook = self.eval_hook.take();
        let result = match hook.as_mut() {
            Some(h) => h(self, expr, locals),
            None => Err("expression eval unavailable (no compiler hook attached)".to_string()) };
        self.eval_hook = hook;
        result
    }

    /// Install the debugger's event simulator (see the `event_fire_hook` field).
    #[allow(clippy::type_complexity)]
    pub fn set_event_fire_hook(
        &mut self,
        hook: Box<dyn FnMut(&mut VM, &str, &str) -> Result<Value, String>>,
    ) {
        self.event_fire_hook = Some(hook);
    }

    /// Fire a GUI event (`control`.`event`) by invoking its registered handler
    /// through this live VM — lets the debugger simulate a click / window-close
    /// without an OS window. Breakpoints inside the handler fire normally
    /// (`invoke_callback` re-enters the instrumented dispatch loop). Returns the
    /// handler's result, or an error if no simulator is attached / no handler is
    /// registered for that control+event.
    pub fn fire_event(&mut self, control: &str, event: &str) -> Result<Value, String> {
        let mut hook = self.event_fire_hook.take();
        let result = match hook.as_mut() {
            Some(h) => h(self, control, event),
            None => Err("event simulation unavailable (no gui hook attached)".to_string()) };
        self.event_fire_hook = hook;
        result
    }

    /// Install the debugger's hot-reload recompiler (see the `reload_hook` field).
    #[allow(clippy::type_complexity)]
    pub fn set_reload_hook(&mut self, hook: Box<dyn FnMut(&mut VM) -> Result<Vec<Chunk>, String>>) {
        self.reload_hook = Some(hook);
    }

    /// Dart-style stateful hot reload (stage 1): recompile the source and swap
    /// the bodies of *changed* functions IN PLACE — heap, globals, and the
    /// current call stack are preserved (`main` is NOT re-run). Only body-only
    /// edits to functions that are not currently executing are applied; anything
    /// structural, or a change to a live function, is rejected with a reason so
    /// old state is never left half-updated. Returns a human-readable report.
    pub fn debug_reload(&mut self) -> Result<String, String> {
        let mut hook = self.reload_hook.take();
        let compiled = match hook.as_mut() {
            Some(h) => h(self),
            None => Err("hot reload unavailable (no compiler hook attached)".to_string()) };
        self.reload_hook = hook;
        let new_chunks = compiled?;
        self.apply_reload(new_chunks)
    }

    fn apply_reload(&mut self, mut new_chunks: Vec<Chunk>) -> Result<String, String> {
        // 1. Structural identity: same count, names in order, imports, and tags.
        // A shifted/renamed/added/removed function invalidates the chunk-index
        // identity every funcref depends on → reject, don't corrupt.
        if new_chunks.len() != self.chunks.len() {
            return Err(format!(
                "structure changed ({} → {} functions) — restart needed",
                self.chunks.len(),
                new_chunks.len()
            ));
        }
        for i in 0..new_chunks.len() {
            if new_chunks[i].name != self.chunks[i].name {
                return Err(format!(
                    "structure changed (function #{i}: '{}' → '{}') — restart needed",
                    self.chunks[i].name, new_chunks[i].name
                ));
            }
        }
        // 2. Content diff: bodies whose code, constants, OR imports changed.
        // Imports matter because JS string literals are `wasm:string-constants`
        // imports (the import name IS the string) — editing a string changes the
        // import table, not the code. Swapping the WHOLE chunk (code+imports
        // together) stays consistent since import resolution reads the chunk's
        // own table lazily per call. (Drops the unchanged runtime prelude.)
        // Imports compared as a SORTED SET, not by order: two compiles can emit
        // the same imports in a different order (HashMap iteration), and since
        // identical code references identical import indices, only a set
        // difference (a string added/removed) is a real change.
        let changed: Vec<usize> = (0..new_chunks.len())
            .filter(|&i| {
                new_chunks[i].code != self.chunks[i].code
                    || !imports_equal_as_set(&new_chunks[i].imports, &self.chunks[i].imports)
                    || format!("{:?}", new_chunks[i].constants)
                        != format!("{:?}", self.chunks[i].constants)
            })
            .collect();
        if changed.is_empty() {
            return Ok("no changes — nothing to reload".to_string());
        }
        // 2b. Exception tags are the one structure cached by chunk index (the
        // per-chunk tag→entity maps), so a body swap must keep them identical;
        // a tag change is beyond a body swap → reject.
        for &i in &changed {
            if !tags_equal(&new_chunks[i].tags, &self.chunks[i].tags) {
                return Err(format!(
                    "exception tags changed in '{}' — restart needed",
                    self.chunks[i].name
                ));
            }
        }
        // 3. Liveness. A changed function that is live in a SUSPENDED async task
        // (a fiber not on the current stack) can't be safely relocated (its saved
        // frames live in fiber storage we don't rewrite) → reject, no corruption.
        // A changed function live on the CURRENT stack is handled by relocation
        // (step 4) — its old body finishes; the next call uses the new body.
        let all_live = self.live_chunk_indices();
        let mut current_live = HashSet::new();
        for f in &self.frames {
            current_live.insert(f.chunk_index);
        }
        let suspended_live: Vec<usize> = changed
            .iter()
            .copied()
            .filter(|i| all_live.contains(i) && !current_live.contains(i))
            .collect();
        if !suspended_live.is_empty() {
            let names: Vec<String> = suspended_live
                .iter()
                .map(|&i| self.chunks[i].name.clone())
                .collect();
            return Err(format!(
                "cannot reload {} — live in a suspended async task (restart)",
                names.join(", ")
            ));
        }

        // 4. Apply. For each changed function:
        //   • not on the stack  → swap the new body in place at its index.
        //   • live on the stack → RELOCATE: copy the old body to a fresh chunk
        //     index, repoint the live frame(s) there so the current activation
        //     keeps running the old (self-consistent) body to completion, then
        //     install the new body at the original index for the next call.
        // Either way the index identity funcrefs depend on now resolves to the
        // new body, and heap/globals/stack are untouched (Dart hot reload).
        let mut reloaded = Vec::new();
        for &i in &changed {
            reloaded.push(self.chunks[i].name.clone());
            if current_live.contains(&i) {
                let old = self.chunks[i].clone();
                let relocated = self.chunks.len();
                self.chunks.push(old);
                // The relocated copy is the SAME module, so it keeps the same
                // type index base — `chunk_type_base` is parallel to `chunks`.
                let base = self.chunk_type_base.get(i).copied().unwrap_or(0);
                self.chunk_type_base.resize(relocated, 0);
                self.chunk_type_base.push(base);
                for f in self.frames.iter_mut() {
                    if f.chunk_index == i {
                        f.chunk_index = relocated;
                    }
                }
            }
            std::mem::swap(&mut self.chunks[i], &mut new_chunks[i]);
            // Pre-scanned BLOCK/LOOP/IF jump targets are keyed by chunk index and
            // derived from code bytes — drop this index's so they rebuild for the
            // new body. (The relocated old index builds its own lazily.)
            self.block_tables.remove(&i);
        }
        // A reloaded body may reference string constants the old one never
        // did, and chunks arrive here by SWAP rather than by the paths that
        // bind on entry. Without this, a literal added by the edit reads
        // `undefined` instead of its text.
        self.bind_imported_globals();

        // 5. Drop cached funcref values whose chunk bodies changed.
        self.funcref_cache.clear();
        let relocated = changed.iter().any(|i| current_live.contains(i));
        let relocated_note = if relocated {
            " (live frame kept on old body until it returns)"
        } else {
            ""
        };
        // `new_chunks.len()` is the original chunk count (relocation appended to
        // `self.chunks`, so use it — not the now-grown live length).
        Ok(format!(
            "reloaded {} function(s): {} · {} unchanged (heap/globals preserved){}",
            reloaded.len(),
            reloaded.join(", "),
            new_chunks.len() - changed.len(),
            relocated_note
        ))
    }

    /// One `(frame chunk-indices, label)` per suspended fiber — for the debugger
    /// `fibers` command. Bottom-to-top frame order per fiber.
    pub fn debug_suspended_fibers(&self) -> Vec<(Vec<usize>, String)> {
        let mut out = Vec::new();
        for ac in &self.active_continuations {
            let idxs: Vec<usize> = ac.caller_fiber.frames.iter().map(|f| f.chunk_index).collect();
            out.push((idxs, "continuation".to_string()));
        }
        let el = self.event_loop.borrow();
        for (pid, fib) in el.waiting_fibers.iter() {
            let idxs: Vec<usize> = fib.frames.iter().map(|f| f.chunk_index).collect();
            out.push((idxs, format!("await promise {pid}")));
        }
        for (fid, fib) in el.future_waiting_fibers.iter() {
            let idxs: Vec<usize> = fib.frames.iter().map(|f| f.chunk_index).collect();
            out.push((idxs, format!("await future {fid}")));
        }
        for (sid, fib) in el.stream_waiting_fibers.iter() {
            let idxs: Vec<usize> = fib.frames.iter().map(|f| f.chunk_index).collect();
            out.push((idxs, format!("await stream {sid}")));
        }
        for task in el.immediate.iter() {
            if let crate::event_loop::Task::ResumeFiber(fib) = task {
                let idxs: Vec<usize> = fib.frames.iter().map(|f| f.chunk_index).collect();
                out.push((idxs, "queued resume".to_string()));
            }
        }
        out
    }

    /// Every chunk index referenced by a LIVE frame — the current call stack plus
    /// all suspended fibers (continuations + event-loop waiters + queued resume
    /// tasks). Current-stack live chunks are relocated on reload; suspended-fiber
    /// live chunks (in this set but not on the current stack) are rejected.
    fn live_chunk_indices(&self) -> HashSet<usize> {
        let mut set = HashSet::new();
        for f in &self.frames {
            set.insert(f.chunk_index);
        }
        for ac in &self.active_continuations {
            for sf in &ac.caller_fiber.frames {
                set.insert(sf.chunk_index);
            }
        }
        let el = self.event_loop.borrow();
        for (_, fib) in el.waiting_fibers.iter() {
            for sf in &fib.frames {
                set.insert(sf.chunk_index);
            }
        }
        for (_, fib) in el.future_waiting_fibers.iter() {
            for sf in &fib.frames {
                set.insert(sf.chunk_index);
            }
        }
        for (_, fib) in el.stream_waiting_fibers.iter() {
            for sf in &fib.frames {
                set.insert(sf.chunk_index);
            }
        }
        for task in el.immediate.iter() {
            if let crate::event_loop::Task::ResumeFiber(fib) = task {
                for sf in &fib.frames {
                    set.insert(sf.chunk_index);
                }
            }
        }
        set
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
                    line }
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
                    _ => Value::F64(l.as_f64() + r.as_f64()) }
            }
            ConstExpr::Mul(left, right) => {
                let l = self.eval_const_expr(left);
                let r = self.eval_const_expr(right);
                match (&l, &r) {
                    (Value::I32(a), Value::I32(b)) => Value::I32(a.wrapping_mul(*b)),
                    (Value::I64(a), Value::I64(b)) => Value::I64(a.wrapping_mul(*b)),
                    (Value::F64(a), Value::F64(b)) => Value::F64(a * b),
                    _ => Value::F64(l.as_f64() * r.as_f64()) }
            }
            ConstExpr::RefFunc(chunk_idx) => {
                if *chunk_idx < self.chunks.len() {
                    let chunk = &self.chunks[*chunk_idx];
                    let func = crate::value::Function {
                        name: Some(chunk.name.clone()),
                        arity: chunk.arity,
                        chunk_index: *chunk_idx,
                        upvalues: Vec::new() };
                    let mut obj = Object::new();
                    obj.kind = ObjectKind::Function(func);
                    Value::Object(crate::heap::alloc(obj))
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
            if self.wasm_tables.len() <= idx {
                self.wasm_tables.resize_with(idx + 1, Vec::new);
            }
            if self.wasm_tables[idx].len() < size {
                self.wasm_tables[idx].resize(size, Value::Null);
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
        let new_len = if entry.is_loop {
            len - depth
        } else {
            len - depth - 1
        };
        let exited_try_count = self.label_stack[new_len..]
            .iter()
            .filter(|label| label.is_try)
            .count();
        for _ in 0..exited_try_count {
            self.exception_handlers.pop();
        }
        self.label_stack.truncate(new_len);
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
        self.func_table[idx] = Value::Object(crate::heap::alloc(obj));

        // Mirror the registration into the Module Records registry.
        // First registration under a given specifier auto-creates a
        // Synthetic ModuleRecord; subsequent registrations add exports.
        // `host_registry` remains the fast lookup path; `modules` is
        // the spec-aligned per-module view.
        self.insert_host_module_export(module, name, ExportEntry::Function { idx });
    }

    /// Debugger-eval support: replace this VM's host-function closures with the
    /// ones registered under the same `(module, name)` in `live`, so host calls
    /// evaluated in this VM hit the LIVE program's captured host state (e.g. the
    /// shared `GuiState` Arc) instead of this VM's fresh, empty state.
    ///
    /// Matched **by name**, not by index: this VM keeps its own index scheme, so
    /// the eval fragment's imports (linked against this VM's `host_registry`) and
    /// any copied namespace refs still resolve. Only the closures' captured
    /// *side-state* is shared — execution and exception state stay isolated in
    /// this separate VM, which is what keeps eval from perturbing the paused
    /// program (a throw is contained here, never on the live handler stack).
    ///
    /// This shares live host state for **reads**; a host fn whose effect is on
    /// `HostContext.vm` (allocation, callbacks, the event loop) still runs
    /// against *this* VM, so it won't mutate the live program. That's the correct
    /// boundary for a debugger.
    pub fn overlay_host_fns_from(&mut self, live: &VM) {
        for ((module, name), &live_idx) in &live.host_registry {
            if let Some(&idx) = self.host_registry.get(&(module.clone(), name.clone())) {
                if let (Some(dst), Some(src)) =
                    (self.host_fns.get_mut(idx), live.host_fns.get(live_idx))
                {
                    *dst = src.clone();
                }
            }
        }
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
            _ => None }
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
                name: target_name } => self.resolve_host_export(from, target_name, visited),
            export => Some(export) }
    }

    fn resolve_host_type_export(&self, module: &str, name: &str) -> Option<crate::TypeDef> {
        let mut visited = Vec::new();
        match self.resolve_host_export(module, name, &mut visited)? {
            ExportEntry::Class { type_id } | ExportEntry::ResourceType { type_id } => {
                self.type_registry.get(*type_id).cloned()
            }
            _ => None }
    }

    /// Create a HostContext with callback capability for host functions.
    pub(crate) fn make_host_context(&mut self) -> HostContext<'_> {
        // We can't pass &mut self into the closure directly due to borrow rules.
        // Instead, we pass raw pointers — this is safe because the HostContext
        // lifetime is strictly scoped within the host function call.
        let vm_ptr = self as *mut VM;
        // Clone the Rc so host functions can enqueue ready work without
        // holding a mutable borrow of the VM.
        let el = self.event_loop.clone();
        // Raw pointer to last_exception — safe: valid for host call duration.
        let exc_ptr = &mut self.last_exception as *mut Option<Value>;
        let exit_ptr = &mut self.pending_exit as *mut bool;
        let exit_code_ptr = &mut self.pending_exit_code as *mut i32;
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
            exit_slot: exit_ptr,
            exit_code_slot: exit_code_ptr,
            globals_slot: globals_ptr,
            stack_slot: &self.stack as *const Vec<Value>,
            handle_table_slot: &self.handle_table as *const crate::handle_table::HandleTable,
            shared_memory_slot: &self.memory as *const crate::shared_memory::SharedMemory }
    }

    /// Close open upvalues in a lambda value that escapes the current stack frame.
    /// When a closure is stored in a host timer wheel (setTimeout), it will run in
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
    /// Run linked chunks with a pre-resolved import table from the Linker.
    /// Used for bootstrap: Linker resolves imports at link time, VM just loads them.
    /// Runs to completion (drains the frame stack) — for top-level entry.
    pub fn run_linked(
        &mut self,
        chunks: Vec<Chunk>,
        resolved_imports: Vec<ImportTarget>,
    ) -> Result<Value, VMError> {
        self.run_linked_impl(chunks, resolved_imports, false)
    }

    /// Re-entrant variant of [`run_linked`] for running dynamically-compiled
    /// chunks (e.g. `eval`) from INSIDE a host function while the VM is mid-
    /// execution. Runs only the newly-pushed script frame to its return via
    /// `execute_until`, leaving the caller's frames intact, then restores the
    /// value stack — the same discipline as `invoke_callback`. Returns the
    /// script's top-level `return` value (else null). Definitions and global
    /// writes persist in this VM, so they escape to the caller's scope.
    pub fn run_linked_nested(
        &mut self,
        chunks: Vec<Chunk>,
        resolved_imports: Vec<ImportTarget>,
    ) -> Result<Value, VMError> {
        self.run_linked_impl(chunks, resolved_imports, true)
    }

    fn run_linked_impl(
        &mut self,
        chunks: Vec<Chunk>,
        resolved_imports: Vec<ImportTarget>,
        nested: bool,
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
                    if ip + 3 >= code.len() {
                        break;
                    }
                    let group = ((code[ip] as u16) << 8) | code[ip + 1] as u16;
                    let sub = ((code[ip + 2] as u16) << 8) | code[ip + 3] as u16;
                    if let Some(op) = Op::decode(group, sub) {
                        if op == Op::REF_FUNC {
                            if ip + 5 < code.len() {
                                let old_idx = ((code[ip + 4] as u16) << 8) | (code[ip + 5] as u16);
                                let new_idx = old_idx + script_idx as u16;
                                code[ip + 4] = (new_idx >> 8) as u8;
                                code[ip + 5] = (new_idx & 0xff) as u8;
                            }
                            ip += 4 + 2 + 1; // 4 opcode + 2 func_idx + 1 uv_count
                            if ip - 1 < code.len() {
                                let uv_count = (code[ip - 1] & 0x7f) as usize;
                                ip += uv_count * 3; // u8 is_local + u16 index
                            }
                            continue;
                        }
                        ip += 4;
                        let fmt = op.operand_format();
                        ip += fmt.size_in(code, ip);
                    } else {
                        ip += 4;
                    }
                }
            }
        }
        self.chunks.extend(adjusted);
        self.bind_imported_globals();

        // Use pre-resolved import table, adjusting ChunkFn indices by offset.
        // Nested dynamic runs temporarily replace this table; restore the
        // caller's table before returning to its dispatch loop.
        let saved_import_table = if nested {
            Some(self.import_table.clone())
        } else {
            None
        };
        self.import_table.clear();
        for target in resolved_imports {
            match target {
                ImportTarget::ChunkFn { chunk_index, arity } => {
                    self.import_table.push(ImportTarget::ChunkFn {
                        chunk_index: chunk_index + script_idx,
                        arity });
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
                ImportTarget::JspiSuspendEager => {
                    self.import_table.push(ImportTarget::JspiSuspendEager);
                }
                ImportTarget::WasiThreadSpawn => {
                    self.import_table.push(ImportTarget::WasiThreadSpawn);
                }
                ImportTarget::JspiYield => {
                    self.import_table.push(ImportTarget::JspiYield);
                }
                ImportTarget::StringConst(s) => {
                    self.import_table.push(ImportTarget::StringConst(s));
                }
            }
        }

        // Load type table. This program is its OWN module — a dynamically
        // compiled one numbers its types from 1 just like the host program,
        // so it gets its own base rather than continuing the caller's space.
        {
            let type_base = self.module_type_ids.len();
            self.set_chunk_type_base(script_idx, type_base);
            let types = self.chunks[script_idx].types.clone();
            if !types.is_empty() {
                // `array.new` immediates index the module's own types in this
                // order; `bind_module_type_ids` turns that order into registry
                // ids once, below, so no name is resolved at run time.
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
                self.bind_module_type_ids(&adjusted_types);
            }
        }

        // Execute
        let saved_frame_depth = self.frames.len();
        let saved_stack_len = self.stack.len();
        self.frames.push(CallFrame {
            chunk_index: script_idx,
            ip: 0,
            base: self.stack.len(),
            label_base: self.label_stack.len(),
            upvalues: Vec::new() });
        self.stack.resize(
            self.stack.len() + self.chunks[script_idx].local_count as usize,
            Value::Null,
        );
        if nested {
            // Run only the script frame we just pushed; leave the caller's
            // frames untouched, then restore its value stack (WASM call-frame
            // semantics — see `invoke_callback`).
            let result = self.execute_until(saved_frame_depth + 1);
            self.stack.truncate(saved_stack_len);
            if let Some(import_table) = saved_import_table {
                self.import_table = import_table;
            }
            result
        } else {
            self.execute()
        }
    }

    pub fn run(&mut self, chunks: Vec<Chunk>) -> Result<Value, VMError> {
        if chunks.is_empty() {
            return Ok(Value::Null);
        }
        // `VYBE_VERIFY=1` — check the structural invariants before running:
        // every instruction on the 4-byte opcode grid, every jump landing on an
        // instruction start. These defects surface far from the emitter that
        // caused them (a nonsense opcode mid-execution), and only for input
        // shapes that change a body's length, which is what makes them
        // expensive to find by hand.
        if std::env::var_os("VYBE_VERIFY").is_some() {
            for (i, chunk) in chunks.iter().enumerate() {
                for issue in crate::debug::verify_chunk(chunk) {
                    eprintln!(
                        "[verify] chunk {} '{}' @{}: {}",
                        i, chunk.name, issue.offset, issue.what
                    );
                }
            }
        }
        // Embedder default: funcref table 0 always exists (spec modules
        // declare their tables; our bundles use table 0 as the thread-start
        // transport for `wasi:threads/thread-spawn`, per the wasi-libc
        // `pthread_create` pattern).
        if self.wasm_tables.is_empty() {
            self.wasm_tables.push(Vec::new());
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
                    if ip + 3 >= code.len() {
                        break;
                    }
                    let group = ((code[ip] as u16) << 8) | code[ip + 1] as u16;
                    let sub = ((code[ip + 2] as u16) << 8) | code[ip + 3] as u16;
                    if let Some(op) = Op::decode(group, sub) {
                        if op == Op::REF_FUNC {
                            // Operand layout after the 4-byte opcode: func_idx
                            // (u16) at ip+4, then uv_count (u8) at ip+6, then
                            // upvalues. Relocate the func_idx by script_idx.
                            if ip + 6 <= code.len() {
                                let old_idx = ((code[ip + 4] as u16) << 8) | (code[ip + 5] as u16);
                                let new_idx = old_idx + script_idx as u16;
                                code[ip + 4] = (new_idx >> 8) as u8;
                                code[ip + 5] = (new_idx & 0xff) as u8;
                            }
                            ip += 4 + 2 + 1; // opcode + func_idx + uv_count
                            if ip - 1 < code.len() {
                                let uv_count = (code[ip - 1] & 0x7f) as usize;
                                ip += uv_count * 3; // per upvalue: u8 is_local + u16 index
                            }
                            continue;
                        }
                        ip += 4; // opcodes are 4 bytes
                        let fmt = op.operand_format();
                        ip += fmt.size_in(code, ip);
                    } else {
                        ip += 4; // skip unknown 4-byte opcode
                    }
                }
            }
        }
        self.chunks.extend(adjusted);
        self.bind_imported_globals();

        let declared_memories = self.chunks[script_idx].memory_min_pages.clone();
        let declared_memory_maxes = self.chunks[script_idx].memory_max_pages.clone();
        self.memory_is_64 = self.chunks[script_idx].memory_is_64.clone();
        self.instantiate_declared_memories(&declared_memories, &declared_memory_maxes)?;
        self.table_is_64 = self.chunks[script_idx].table_is_64.clone();
        let declared_tables = self.chunks[script_idx].table_min_sizes.clone();
        self.wasm_table_maxes = self.chunks[script_idx]
            .table_max_sizes
            .iter()
            .map(|m| m.map(|v| v as usize))
            .collect();
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
                                        upvalues: Vec::new() };
                                    let mut obj = Object::new();
                                    obj.kind = ObjectKind::Function(func);
                                    Value::Object(crate::heap::alloc(obj))
                                } else {
                                    Value::Null
                                }
                            }
                            other => other })
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
        // Passive element segments compiled from source: resolve each function
        // chunk index to its canonical funcref and populate `elem_segments`, so
        // `table.init`/`array.new_elem` copy real funcrefs (WASM passive elems).
        let passive_elem_funcs = self.chunks[script_idx].passive_elem_funcs.clone();
        for (seg_idx, funcs) in passive_elem_funcs.iter().enumerate() {
            if funcs.is_empty() {
                continue;
            }
            let vals: Vec<crate::value::Value> =
                funcs.iter().map(|&fi| self.make_funcref(fi)).collect();
            self.set_elem_segment(seg_idx, vals);
        }

        // Resolve imports for ALL new chunks (not just script chunk).
        // Each chunk has its own import list. We build one unified import table
        // by scanning all chunks and mapping their import indices to host functions.
        // The trick: all chunks compiled by the same compiler share the same import list
        // (imports are added to chunks[0] by all compilers). For multi-module programs,
        // different modules may have different imports. We resolve the union.
        self.import_table.clear();
        let script_imports = self.chunks[script_idx].imports.clone();
        for import in &script_imports {
            let target = self.resolve_import_target(&import.module, &import.name)?;
            self.import_table.push(target);
        }

        // Load type table from the script chunk (WASM GC type section).
        // Registers user-defined class types and their vtable methods.
        // Sets __tid_<name> globals so constructors can stamp type_id.
        {
            // This run's chunks are their own module: their type immediates
            // are numbered from 1 against the table loaded just below.
            let type_base = self.module_type_ids.len();
            self.set_chunk_type_base(script_idx, type_base);
            let types = self.chunks[script_idx].types.clone();
            if !types.is_empty() {
                // `array.new` immediates index the module's own types in this
                // order; `bind_module_type_ids` turns that order into registry
                // ids once, below, so no name is resolved at run time.
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
                self.bind_module_type_ids(&adjusted_types);
                // The `__tid_<name>` globals that used to be written here are
                // gone: they existed so a constructor could GLOBAL_GET a
                // registry id and stamp it after allocating. Allocation
                // carries the type now (`struct.new_default $T`), so nothing
                // read them — measured, zero live readers — and writing one
                // per type per program was a name-keyed type lookup kept
                // alive by a dead consumer.
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
            upvalues: Vec::new() });

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
                id } => {
                self.run_event_loop()?;
                if self.has_pending_jspi() {
                    Err(VMError::new(format!("__jspi__:{}", id)))
                } else {
                    // The suspended script fiber completed inside the event
                    // loop — surface its final value (top-level await).
                    Ok(self.last_fiber_completion.take().unwrap_or(Value::Null))
                }
            }
            ExecResult::Suspended { .. } => {
                self.run_event_loop()?;
                Ok(self.last_fiber_completion.take().unwrap_or(Value::Null))
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

    pub(crate) fn read_leb_i32(&mut self) -> i32 {
        let mut result: u32 = 0;
        let mut shift = 0u32;
        loop {
            let byte = self.read_byte();
            result |= ((byte & 0x7f) as u32) << shift;
            shift += 7;
            if byte & 0x80 == 0 {
                if shift < 32 && (byte & 0x40) != 0 {
                    result |= !0u32 << shift;
                }
                break;
            }
        }
        result as i32
    }

    pub(crate) fn read_leb_i64(&mut self) -> i64 {
        let mut result: u64 = 0;
        let mut shift = 0u32;
        loop {
            let byte = self.read_byte();
            result |= ((byte & 0x7f) as u64) << shift;
            shift += 7;
            if byte & 0x80 == 0 {
                if shift < 64 && (byte & 0x40) != 0 {
                    result |= !0u64 << shift;
                }
                break;
            }
        }
        result as i64
    }

    pub(crate) fn read_f32(&mut self) -> f32 {
        let mut bytes = [0u8; 4];
        for b in &mut bytes {
            *b = self.read_byte();
        }
        f32::from_le_bytes(bytes)
    }

    pub(crate) fn read_f64(&mut self) -> f64 {
        let mut bytes = [0u8; 8];
        for b in &mut bytes {
            *b = self.read_byte();
        }
        f64::from_le_bytes(bytes)
    }

    pub(crate) fn get_constant(&self, index: u16) -> Value {
        let f = self.frame();
        self.chunks[f.chunk_index].constants[index as usize].clone()
    }

    /// Resolve a chunk-level tag index to its tag ENTITY (spec EH identity).
    /// Maps are built lazily so every chunk-installation path is covered.
    pub(crate) fn resolve_chunk_tag(
        &mut self,
        chunk_index: usize,
        tag_idx: u16,
    ) -> Result<usize, VMError> {
        if self.chunk_tag_maps.len() < self.chunks.len() {
            self.resolve_chunk_tags();
        }
        self.chunk_tag_maps
            .get(chunk_index)
            .and_then(|m| m.get(tag_idx as usize))
            .copied()
            .ok_or_else(|| {
                VMError::new(format!(
                    "unknown exception tag index {tag_idx} in chunk {chunk_index}"
                ))
            })
    }

    /// Build tag→entity maps for chunks that don't have one yet (spec EH
    /// instantiation): LOCAL declarations mint fresh entities — "created
    /// fresh each time" — while IMPORTS resolve by name to a shared entity
    /// (entity 0 is the host-provided `vybe:exception` tag).
    pub(crate) fn resolve_chunk_tags(&mut self) {
        while self.chunk_tag_maps.len() < self.chunks.len() {
            let ci = self.chunk_tag_maps.len();
            let decls = self.chunks[ci].tags.clone();
            let mut map = Vec::with_capacity(decls.len());
            for decl in decls {
                let entity = if decl.imported {
                    if let Some(&id) = self.imported_tag_registry.get(&decl.debug_name) {
                        id
                    } else {
                        self.tag_entities.push(TagEntity {
                            debug_name: decl.debug_name.clone(),
                            arity: decl.arity });
                        let id = self.tag_entities.len() - 1;
                        self.imported_tag_registry.insert(decl.debug_name, id);
                        id
                    }
                } else {
                    self.tag_entities.push(TagEntity {
                        debug_name: decl.debug_name,
                        arity: decl.arity });
                    self.tag_entities.len() - 1
                };
                map.push(entity);
            }
            self.chunk_tag_maps.push(map);
        }
    }

    /// Record that chunks `first..` belong to a module whose type index space
    /// starts at `base`.
    ///
    /// **Invariant: every site that grows `self.chunks` must call this.** A
    /// chunk with no entry reads base `0` — right for the first module, wrong
    /// for every later one, and wrong here means a valid-but-foreign rtt,
    /// which `test_type` prefers over the fallbacks it would otherwise use.
    /// The four growth sites are `run`, the dynamic path, `run_components`,
    /// and hot-reload relocation.
    pub(crate) fn set_chunk_type_base(&mut self, first: usize, base: usize) {
        if self.chunk_type_base.len() < self.chunks.len() {
            self.chunk_type_base.resize(self.chunks.len(), 0);
        }
        for slot in self.chunk_type_base.iter_mut().skip(first) {
            *slot = base;
        }
    }

    /// Append `types` to the module's type index space, resolving each name to
    /// its registry id **once, at load**.
    ///
    /// Must run AFTER `load_type_table`, or the lookup finds nothing and the
    /// slot silently becomes rtt `0`. The registry id is not the compile-time
    /// table position — the host pre-registers its builtin types ahead of the
    /// module's — so the mapping is what the index space is FOR.
    pub(crate) fn bind_module_type_ids(&mut self, types: &[crate::chunk::TypeEntry]) {
        for entry in types {
            let id = self.type_registry.get_id(&entry.name).unwrap_or(0);
            self.module_type_names.push(entry.name.clone());
            self.module_type_ids.push(id);
        }
    }

    /// Install the host scheduler — the policy half of async. Called at
    /// plugin registration (the ecma platform installs the ECMA-262 §9.5 job
    /// discipline); the VM itself never decides which callback runs next.
    /// Register a host-owned deferred-work source (a timer wheel). Plugin
    /// init calls this, exactly like `register_host_fn` / `set_scheduler`.
    pub fn register_deferred_source(
        &mut self,
        source: std::sync::Arc<dyn crate::scheduler::DeferredSource>,
    ) {
        self.deferred_sources.push(source);
    }

    /// Any host-deferred work registered (due or not)?
    pub fn deferred_pending(&self) -> bool {
        self.deferred_sources.iter().any(|s| s.has_pending())
    }

    /// Pop ONE due deferred callback across the registered sources, in
    /// registration order. The one-task-per-turn contract lives in the
    /// caller (the drain); this is only the mechanism.
    pub fn next_due_deferred(&self) -> Option<Value> {
        self.deferred_sources.iter().find_map(|s| s.pop_due())
    }

    /// Sleep until the earliest deferred deadline (`wasi:clocks` 
    /// subscribe-duration shape). Returns immediately if ready work exists.
    pub fn wait_for_deferred(&self) {
        if self.event_loop.borrow().has_pending() {
            return;
        }
        let earliest = self
            .deferred_sources
            .iter()
            .filter_map(|s| s.earliest_deadline_ms())
            .reduce(f64::min);
        if let Some(earliest) = earliest {
            let now = crate::event_loop::monotonic_now_ms();
            if earliest > now {
                std::thread::sleep(std::time::Duration::from_millis((earliest - now) as u64));
            }
        }
    }

    pub fn set_scheduler(&mut self, scheduler: std::sync::Arc<dyn crate::scheduler::Scheduler>) {
        self.scheduler = Some(scheduler);
    }

    /// Mechanism surface for host schedulers: run one scheduled callback.
    pub fn run_scheduled_callback(
        &mut self,
        callback: &Value,
        args: &[Value],
    ) -> Result<Value, VMError> {
        self.invoke(callback, args)
    }

    /// Mechanism surface for host schedulers: resume a suspended fiber and
    /// record its completion (top-level await surfaces the program's final
    /// value through this).
    pub fn resume_scheduled_fiber(&mut self, fiber: crate::fiber::Fiber) -> Result<(), VMError> {
        let completion = self.resume_fiber(fiber)?;
        self.last_fiber_completion = Some(completion);
        Ok(())
    }

    /// Create the globals declared by the module's **global imports**.
    ///
    /// js-string-builtins § String constants: when a namespace is designated
    /// for string constants, *"every import that refers to this namespace has
    /// a global created to hold the string constant specified in the import
    /// field"*. Creating it is the host's job, and here the VM is the host.
    ///
    /// This runs wherever chunks enter the VM rather than in the load-once
    /// path, because that path only sees the script chunk while every chunk
    /// declares its own imports. A chunk that is never bound fails silently —
    /// `GLOBAL_GET` yields `Undefined`, so every string literal in it would
    /// read as `undefined` rather than trap.
    pub(crate) fn bind_imported_globals(&mut self) {
        let bindings: Vec<(String, Arc<str>)> = self
            .chunks
            .iter()
            .flat_map(|chunk| chunk.global_imports.iter())
            .filter(|import| import.module == crate::chunk::STRING_CONSTANTS_MODULE)
            .map(|import| {
                (
                    crate::chunk::imported_global_key(&import.module, &import.name),
                    Arc::from(import.name.as_str()),
                )
            })
            .collect();
        for (key, value) in bindings {
            self.globals.insert(key, Value::String(value));
        }
    }

    /// THE import-resolution policy — the single copy. Every path that maps
    /// an `(module, name)` import to an `ImportTarget` goes through here:
    /// `run`'s link loop, per-chunk lazy resolution (`resolve_chunk_import`),
    /// and the dynamic compiler service's `resolve_imports`.
    ///
    /// Order: VM-implemented imports (jspi, wasi:threads) → string constants
    /// (the import name IS the value, per js-string-builtins) → host
    /// functions through Module Records → `"*"` wildcard against globals →
    /// `__vybe_<name>` stdlib redirect → loud error.
    pub fn resolve_import_target(
        &self,
        module: &str,
        name: &str,
    ) -> Result<ImportTarget, VMError> {
        if module == "jspi" && name == "await" {
            return Ok(ImportTarget::JspiSuspend);
        }
        if module == "jspi" && name == "await_eager" {
            return Ok(ImportTarget::JspiSuspendEager);
        }
        if module == "jspi" && name == "yield" {
            return Ok(ImportTarget::JspiYield);
        }
        if module == "wasi:threads" && name == "thread-spawn" {
            return Ok(ImportTarget::WasiThreadSpawn);
        }
        if module == "wasm:string-constants" {
            return Ok(ImportTarget::StringConst(Arc::from(name)));
        }
        if let Some(idx) = self.resolve_host_function_index(module, name) {
            return Ok(ImportTarget::Host(idx));
        }
        if module == "*" {
            let candidates = [name.to_string(), name.to_lowercase()];
            if let Some(global_name) = candidates
                .iter()
                .find(|g| self.globals.contains_key(g.as_str()))
            {
                return Ok(ImportTarget::StdlibRedirect(global_name.clone()));
            }
        }
        let candidates = [
            format!("__vybe_{}", name),
            format!("__vybe_{}", name.to_lowercase()),
        ];
        if let Some(global_name) = candidates
            .iter()
            .find(|g| self.globals.contains_key(g.as_str()))
        {
            return Ok(ImportTarget::StdlibRedirect(global_name.clone()));
        }
        Err(VMError::new(format!(
            "Unresolved import: \"{}\" \"{}\"",
            module, name
        )))
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
        self.resolve_import_target(&import.module, &import.name)
            .map(Some)
    }

    pub(crate) fn constant_str(&self, index: u16) -> String {
        match &self.get_constant(index) {
            Value::String(s) => s.to_string(),
            v => format!("{}", v) }
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
                    id })
            }
            Err(e) if e.message.starts_with("__jspi__:") => {
                let id: u64 = e.message["__jspi__:".len()..].parse().unwrap_or(0);
                Ok(ExecResult::Suspended {
                    kind: SuspensionKind::Jspi,
                    id })
            }
            Err(e) if e.message.starts_with("__future__:") => {
                let id: u64 = e.message["__future__:".len()..].parse().unwrap_or(0);
                Ok(ExecResult::Suspended {
                    kind: SuspensionKind::Future,
                    id })
            }
            Err(e) if e.message.starts_with("__stream_read__:") => {
                let id: u64 = e.message["__stream_read__:".len()..].parse().unwrap_or(0);
                Ok(ExecResult::Suspended {
                    kind: SuspensionKind::StreamRead,
                    id })
            }
            Err(e) => Err(e) }
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

#[cfg(test)]
mod reset_tests {
    use super::*;
    use crate::value::Object;

    /// End-to-end proof of VM hot-reset: after `reset_to`, a baseline global
    /// survives with its boot contents, a script-added global is gone, a
    /// script mutation to a baseline object is undone, and cyclic script
    /// garbage (which pure Arc refcounting can NEVER free) is reclaimed.
    #[test]
    fn reset_to_frees_script_state_including_cycles_and_restores_baseline() {
        crate::heap::enable_tracking();
        let mut vm = VM::new();

        // Baseline: a global bound to a tracked heap object with one property.
        let base_obj = crate::heap::alloc(Object::new());
        base_obj
            .lock()
            .unwrap()
            .properties
            .insert("boot".into(), Value::I32(7));
        vm.globals
            .insert("baseline".into(), Value::Object(base_obj.clone()));

        let snap = vm.snapshot();

        // Script mutations: (a) a new global, (b) a mutation of the baseline
        // object, (c) a reference cycle rooted in a new global.
        let a = crate::heap::alloc(Object::new());
        let b = crate::heap::alloc(Object::new());
        a.lock()
            .unwrap()
            .properties
            .insert("b".into(), Value::Object(b.clone()));
        b.lock()
            .unwrap()
            .properties
            .insert("a".into(), Value::Object(a.clone()));
        let (wa, wb) = (Arc::downgrade(&a), Arc::downgrade(&b));
        vm.globals.insert("script".into(), Value::Object(a.clone()));
        base_obj
            .lock()
            .unwrap()
            .properties
            .insert("mutated".into(), Value::I32(1));
        drop(a);
        drop(b);
        // Under pure refcounting the cycle is still alive here.
        assert!(wa.upgrade().is_some() && wb.upgrade().is_some());

        vm.reset_to(&snap);

        // Baseline global survives, boot contents intact, script mutation gone.
        assert!(vm.globals.contains_key("baseline"));
        let base = base_obj.lock().unwrap();
        assert_eq!(base.properties.get("boot"), Some(&Value::I32(7)));
        assert!(!base.properties.contains_key("mutated"));
        drop(base);
        // Script-added global gone.
        assert!(!vm.globals.contains_key("script"));
        // Cyclic script garbage reclaimed.
        assert!(
            wa.upgrade().is_none() && wb.upgrade().is_none(),
            "reset_to must free cyclic script garbage"
        );
        // Transient state clean.
        assert!(vm.stack.is_empty() && vm.frames.is_empty());
    }

    /// Cycle-leak proof: N cyclic structures rooted in script globals are fully
    /// reclaimed by `reset_to` — the live-object count returns to the post-boot
    /// baseline, not N-leaked (pure refcounting would leak every cycle forever).
    #[test]
    fn reset_to_returns_live_count_to_baseline_after_cycles() {
        crate::heap::enable_tracking();
        let mut vm = VM::new();
        // A baseline object so the baseline count is a real, non-trivial number.
        let _base = crate::heap::alloc(Object::new());
        vm.globals
            .insert("keep".into(), Value::Object(_base.clone()));
        let snap = vm.snapshot();
        let baseline = crate::heap::live_count();

        // Allocate 500 independent a↔b cycles, each rooted in its own global.
        for i in 0..500 {
            let a = crate::heap::alloc(Object::new());
            let b = crate::heap::alloc(Object::new());
            a.lock()
                .unwrap()
                .properties
                .insert("b".into(), Value::Object(b.clone()));
            b.lock()
                .unwrap()
                .properties
                .insert("a".into(), Value::Object(a.clone()));
            vm.globals.insert(format!("cyc{i}"), Value::Object(a));
        }
        assert!(
            crate::heap::live_count() >= baseline + 1000,
            "the 1000 cycle objects must be live before reset"
        );

        vm.reset_to(&snap);

        assert_eq!(
            crate::heap::live_count(),
            baseline,
            "reset_to must reclaim every cyclic structure — live count back to baseline"
        );
    }
}
