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

/// Spec element-count bound for a 32-bit table: a table32's size must fit the
/// index type, i.e. `2^32 - 1` elements (WASM 3.0, limits of `i32` tabletypes).
/// `table.grow` past it MUST report `-1` — table_grow.wast:49 grows a `0x10`
/// table by `0xffff_fff0`, which lands on `2^32` exactly and must fail.
/// (table64's `2^64 - 1` is unreachable here — `TABLE_ALLOC_LIMIT` bites first.)
pub(crate) const MAX_TABLE32_ELEMS: u64 = 0xffff_ffff;

/// RESOURCE ceiling — the largest table this host will actually ALLOCATE.
///
/// This is deliberately NOT the spec bound above, and must not be "corrected"
/// to it. `2^32 - 1` elements is a legal table32 size to declare and to grow
/// to, but every element is a `Value`, so honouring it means a multi-gigabyte
/// `Vec` — `(table 0xffff_ffff funcref)` is a ~70GB `resize` that gets the
/// process OOM-killed instead of failing (table.wast exits 137).
///
/// The spec permits refusing: `table.grow` may report `-1` for ANY reason
/// including allocation failure, and instantiation may fail on resource
/// exhaustion. So the bound is a policy choice, and 2^26 is chosen to sit far
/// above anything a real program tables (the whole official suite never grows
/// past 800 elements) while keeping the worst case to a couple of gigabytes.
///
/// The same reasoning already governs memory: see `SharedMemory::grow`, where
/// widening the ceiling to memory64's spec limit would reinstate that hang.
pub(crate) const TABLE_ALLOC_LIMIT: u64 = 1 << 26;

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
    /// Raw pointer to `VM::host_receiver` — the thisArgument for the NEXT
    /// callback this host function makes.
    ///
    /// The compiler emits no ambient receiver global, so the host needs a
    /// channel of its own — `slot_set` deliberately never creates a global
    /// (a module's global index space is fixed at instantiation;
    /// a global conjured at runtime cannot exist on a stock engine). Writing it
    /// therefore became a SILENT NO-OP, which is what made `f.call(o)` and
    /// `map(fn, thisArg)` lose their receiver while plain `map(fn)` was fine.
    ///
    /// This slot is the host's own channel for that value. It is VM state, not
    /// a module global, so it costs the emitted module nothing and disappears
    /// from the WASM surface entirely — which is the point of M5.
    host_receiver_slot: *mut Value,
    /// The running module's receiver ABI, copied from its module chunk.
    ///
    /// It is what decides which of the two channels above is live, and it has
    /// to be carried explicitly: a `HostContext` holds raw slots, not the
    /// chunk table, so it cannot ask the module itself. Defaults to `Ambient`
    /// in `empty()`, the behaviour every host had before the second channel
    /// existed.
    module_receiver_abi: crate::chunk::ReceiverAbi,
    /// How many leading argument slots THIS call handed to the receiver.
    ///
    /// Set per invocation by `call_value_inner`, because it is a property of
    /// the CALL: a JS call from bytecode under `ReceiverAbi::Parameter`
    /// carries a receiver, and host plumbing invoking the same function with
    /// arguments it built itself does not.
    /// Raw pointer to VM.pending_exit_code — the status `wasi:cli/exit` was
    /// given. Null when no VM is attached (`HostContext::empty()`).
    exit_code_slot: *mut i32,
    /// Raw pointer to VM globals for host-managed JS receiver binding.
    /// Null when no VM is attached (HostContext::empty()).
    /// Globals are indexed by globalidx; host interop still asks BY NAME, so
    /// the slot carries both halves — the storage and the resolver.
    globals_slot: *mut Vec<Value>,
    global_index_slot: *mut HashMap<String, u32>,
    /// Raw pointer to VM.stack for closing escaped upvalues in timer callbacks.
    /// Null when no VM is attached (HostContext::empty()).
    #[allow(dead_code)]
    stack_slot: *const Vec<Value>,
    /// Raw pointer to the CM3 handle table, so host functions receiving a
    /// canon `stream<u8>` / `future<T>` i32 handle (CanonicalABI lowering)
    /// can resolve it to the EventLoop stream/future id — and so a host
    /// function PRODUCING a `stream<u8>` can LIFT it to a handle, which is
    /// why this is `*mut` rather than `*const`. Lowering without lifting was
    /// only half the ABI: it let the guest hand us a stream but gave the guest
    /// no conforming way to read one back.
    /// Null when no VM is attached (HostContext::empty()).
    handle_table_slot: *mut crate::handle_table::HandleTable,
    /// Raw pointer to the VM's shared memory, for the `wasm:threads`
    /// scheduler intrinsics (`all_parked`). Null when no VM is attached.
    shared_memory_slot: *const crate::shared_memory::SharedMemory,
    /// Raw pointers to the call-tag tables, so a host function can ask which
    /// CONVENTION a funcref answers to instead of inferring it from arity.
    ///
    /// Property dispatch needs exactly this: `__set_x(self, v)` and a JS
    /// `defineProperty` `set(v)` are the same wasm signature and opposite
    /// conventions. Declarations are resolved at load, before any host call, so
    /// a read-only view is enough.
    /// The VM's `TypeRegistry`, for reading a typed object's DECLARED FIELD
    /// NAMES.
    ///
    /// ⛔ `[[OwnPropertyKeys]]` COULD NOT SEE A TYPED OBJECT'S FIELDS. The
    /// host's key walk reads `ObjectKind` and the `properties` map, so a class
    /// whose fields live in indexed GC storage enumerated as EMPTY —
    /// `Object.keys`, `getOwnPropertyNames`, `in`, `for…in` and
    /// `JSON.stringify` all missed them, while `o.a` read back fine. That is
    /// why `instance_fields_are_own_properties` has to WITHHOLD indexed
    /// storage from js today (`classes.rs:2548`): the optimization is correct
    /// and unobservable at the same time.
    ///
    /// A raw slot, not a `&VM` — same contract as the neighbours above: valid
    /// for the duration of one host call, read-only, and null when no VM is
    /// attached (`HostContext::empty()`).
    type_registry_slot: *const crate::typedef::TypeRegistry,
    call_tag_registry_slot: *const HashMap<String, u32>,
    func_call_tags_slot: *const HashMap<usize, Vec<u32>>,
    /// Raw pointer to the chunk table, so the receiver query can fall back to
    /// `Chunk.is_method` for a chunk that declares no tag.
    chunks_slot: *const Vec<crate::chunk::Chunk>,
}

// SAFETY: HostContext is always created and used on the VM's owning thread.
// The raw pointer to last_exception_slot is valid for the duration of the host
// function call (same scope as the invoker lifetime bound by 'a).
unsafe impl Send for HostContext<'_> {}

impl<'a> HostContext<'a> {
    /// Does this funcref declare that it handles `tag`?
    ///
    /// The question property dispatch actually wants — "does parameter 0 mean
    /// the receiver?" — asked rather than inferred. A func that declares
    /// nothing answers `false`, which is the ambient-receiver case: its
    /// receiver comes from the call, not from an argument.
    /// The DECLARED fields of a typed object's class as `(name, enumerable)`,
    /// in declaration order — the half of `[[OwnPropertyKeys]]` that lives in
    /// the type rather than in the object.
    ///
    /// ⚠ ENUMERABILITY IS RETURNED, NOT APPLIED. `Object.keys` and `for…in`
    /// take only the enumerable ones (§20.1.2.17); `getOwnPropertyNames` takes
    /// all of them (§20.1.2.10). Filtering here would make one of the two
    /// wrong, and the descriptor is the type's to state — `field_descriptors`
    /// already carries it (WASM Annotations, `@ecma262` namespace).
    ///
    /// Empty for `type_id == 0` (an untyped/dynamic object, whose keys are all
    /// in the `properties` map already) and when no VM is attached. Order is
    /// the type's own field order, which is the order ECMA-262 §10.2.11
    /// created them in, so a caller can concatenate without sorting.
    pub fn declared_fields(&self, type_id: usize) -> Vec<(String, bool)> {
        if type_id == 0 || self.type_registry_slot.is_null() {
            return Vec::new();
        }
        // SAFETY: valid for the duration of this host call — same contract as
        // every other slot on this struct.
        let registry = unsafe { &*self.type_registry_slot };
        let Some(t) = registry.types.get(type_id) else {
            return Vec::new();
        };
        t.fields()
            .into_iter()
            .map(|name| {
                let enumerable = t.get_field_descriptor(&name).enumerable;
                (name, enumerable)
            })
            .collect()
    }

    pub fn func_handles_call_tag(&self, func: &Value, tag: &str) -> bool {
        if self.call_tag_registry_slot.is_null() || self.func_call_tags_slot.is_null() {
            return false;
        }
        let Value::Object(obj) = func else {
            return false;
        };
        let chunk_index = match &obj.lock().unwrap().kind {
            ObjectKind::Function(f) => f.chunk_index,
            _ => return false,
        };
        // SAFETY: same contract as every other slot here — the pointers are the
        // VM's own tables, valid for the duration of the host call.
        let (registry, func_tags) = unsafe {
            (&*self.call_tag_registry_slot, &*self.func_call_tags_slot)
        };
        if let Some(&id) = registry.get(tag)
            && let Some(tags) = func_tags.get(&chunk_index)
        {
            return tags.contains(&id);
        }
        // No tag declared. Fall back to `Chunk.is_method`, which states the
        // same fact for compiler-generated accessors: "arity includes an
        // implicit leading receiver". It is a DECLARATION too, not a guess —
        // the compiler sets it where it builds the accessor — and it covers the
        // shapes the tag has not reached yet.
        if self.chunks_slot.is_null() {
            return false;
        }
        // SAFETY: as above — the VM's own table, valid for this host call.
        let chunks = unsafe { &*self.chunks_slot };
        // `is_method` alone is NOT sufficient evidence: a JS object-literal
        // accessor is compiled as a plain lambda whose chunk may carry it for
        // unrelated reasons, and treating that as receiver-first hands a
        // one-parameter setter the receiver as its VALUE. Require the arity to
        // agree with a receiver-first setter as well.
        // ⛔ `arity >= 2` IS A GUESS, AND UNDER `ReceiverAbi::Parameter` THERE
        // IS NO NEED TO GUESS. The heuristic was written when an `is_method`
        // chunk had NO receiver parameter, so "at least two parameters" stood
        // in for "receiver-first" and excluded a plain one-value setter. Under
        // `Parameter` the chunk STATES the answer in `takes_receiver`, so ask
        // it — a declaration always beats an arity inference.
        //
        // ⛔ And do NOT raise the threshold to 3 there instead: a receiver-first
        // setter is `(receiver, value)`, arity 2, so `>= 3` rejects exactly the
        // accessor this must accept. Measured — `defineProperty(o,"v",{set})`
        // then `o.v = 9` wrote nowhere and `this._x` stayed undefined.
        let chunk = chunks.get(chunk_index);
        if self.module_uses_host_receiver_channel() {
            return chunk.is_some_and(|c| c.takes_receiver);
        }
        chunk.is_some_and(|c| c.is_method && c.arity >= 2)
    }


    /// Does THIS MODULE pass receivers as parameters?
    ///
    /// ⛔ NOT the same question as [`Self::receiver_argc`], and mixing them up
    /// is measurable. `receiver_argc` asks "did the call that reached ME carry
    /// a receiver slot" — a property of one invocation, and 0 whenever host
    /// plumbing built the arguments. This asks "will a callee I invoke expect
    /// one", which is a property of the module and true regardless of how the
    /// current host function was itself reached.
    ///
    /// Using `receiver_argc` for the second question left an object-literal
    /// setter being handed the receiver TWICE — once explicitly by the call
    /// site and once prepended by the VM — so `this` was `undefined` and the
    /// value parameter got the receiver.
    pub fn receiver_is_parameter(&self) -> bool {
        self.module_receiver_abi == crate::chunk::ReceiverAbi::Parameter
    }

    /// The USER arguments of a host call: the leading `captures` bound values
    /// and the receiver slot both skipped.
    ///
    /// A host function reached as a value (`d.resolve(42)`) is handed
    /// `[captures…, receiver, args…]`. `captures` is the function's OWN fact —
    /// it built them — while the receiver slot is the module's ABI, which the
    /// function cannot know. Asking here keeps the second half in one place
    /// rather than in every settler's index arithmetic.
    /// Invoke `target` with an EXPLICIT receiver, exactly once.
    ///
    /// ⛔ THE PRIMITIVE HOST CODE WAS MISSING. `invoke` prepends a receiver it
    /// INFERS, which is right when the host has none — a `map` callback, a
    /// microtask — and wrong wherever the host knows it. An accessor knows the
    /// object it is reading; passing it through the argument list on top of the
    /// inferred one handed the callee TWO, so `this` took the inferred value
    /// and the first declared parameter took the real receiver.
    ///
    /// Works under BOTH bindings without a branch at the call site: setting the
    /// receiver is what the ambient binding needs, and under `Parameter` it is
    /// also the channel `invoke` prepends FROM — so the callee is handed
    /// `[receiver, args…]` either way.
    pub fn invoke_with_receiver(
        &mut self,
        target: &Value,
        receiver: Value,
        args: &[Value],
    ) -> Value {
        // ⛔ A HOST callee reads its receiver as ARGUMENT 0 and never consults
        // the ambient channel — `ecma:array.map` opens `array_of(args, 0)`.
        // Setting the receiver and invoking would hand a host setter only its
        // value: measured as `u.pathname = "/z"` on a `URL` silently doing
        // nothing while the getters were fine. A BYTECODE callee is the
        // opposite: `invoke` prepends from the channel, so setting it is
        // exactly right and prepending here as well would give it two.
        // ⛔ A BOUND WRAPPER IS A HOST FUNCTION THAT MUST *NOT* BE PREPENDED
        // TO. Its `__bound_args` captures are prepended by the dispatch and it
        // reads them at fixed offsets, so an extra leading value shifts them —
        // and §20.2.3.2 says a bound function IGNORES the thisArg of a later
        // call anyway, because it already closed over one. Measured:
        // `f.bind(null, 1).apply(null, [2])` called `f(1, null)` instead of
        // `f(1, 2)`, the prepended receiver landing between the partial and the
        // real argument.
        let host_callee = Self::is_unbound_host_fn(target);
        // ⛔ DO NOT PREPEND HERE. The receiver slot is filled in exactly ONE
        // place — `call_value_inner`, which puts it at argument 0 of every host
        // callee — so prepending again hands the callee two and shifts every
        // real argument: measured, `u.pathname = "/z"` on a `URL` silently did
        // nothing. This function's job is to say WHICH receiver, not to place
        // it; binding the channel is the whole of it, for host and bytecode
        // callees alike.
        let previous = self.current_js_this();
        self.set_js_this(receiver.clone());
        // ⛔ WHO FILLS THE SLOT DEPENDS ON THE ABI, AND ONLY ON THE ABI.
        //
        // Under `ReceiverAbi::Parameter` `call_value_inner` puts the receiver
        // at argument 0 of EVERY host callee, so prepending here as well hands
        // it two and shifts every real argument (`u.pathname = "/z"` on a
        // `URL` silently did nothing).
        //
        // Under the AMBIENT ABI that site pushes nothing at all, so this is
        // the only thing that places a receiver for a host callee — removing
        // it dropped the receiver for every ambient language and cost 156
        // csharp regressions, an expression-bodied property answering `0`.
        //
        // This is a placement decision made by the CALLER from the module's
        // ABI, which is uniform for the whole module. It is NOT the per-call
        // arity flag that was removed: under either ABI a host callee's
        // argument list has one shape, and the callee is never asked which
        // kind of call it was in.
        let placed = self.place_receiver(host_callee, receiver, args);
        let out = match &placed {
            Some(all) => self.invoke(target, all),
            None => self.invoke(target, args),
        };
        self.set_js_this(previous);
        out
    }

    /// Like [`Self::invoke_with_receiver`], but surfaces a thrown value
    /// instead of swallowing it.
    ///
    /// The receiver placement is [`Self::place_receiver`] in both, so the two
    /// entry points can never disagree about who fills the slot.
    pub fn try_invoke_with_receiver(
        &mut self,
        target: &Value,
        receiver: Value,
        args: &[Value],
    ) -> Result<Value, Value> {
        let host_callee = Self::is_unbound_host_fn(target);
        let previous = self.current_js_this();
        self.set_js_this(receiver.clone());
        let placed = self.place_receiver(host_callee, receiver, args);
        let out = match &placed {
            Some(all) => self.try_invoke(target, all),
            None => self.try_invoke(target, args),
        };
        self.set_js_this(previous);
        out
    }

    /// A host callee that is not a bound wrapper — the only kind whose
    /// receiver slot the AMBIENT ABI leaves for the caller to fill.
    fn is_unbound_host_fn(target: &Value) -> bool {
        matches!(
            target,
            Value::Object(o)
                if matches!(
                    o.lock().map(|g| matches!(g.kind, crate::value::ObjectKind::HostFunction(_))
                        && !g.properties.contains_key("__bound_args")),
                    Ok(true)
                )
        )
    }

    /// `Some(args)` when the CALLER must place the receiver, `None` when the
    /// VM already does.
    ///
    /// Under `ReceiverAbi::Parameter` the answer is always `None`:
    /// `call_value_inner` puts the receiver at argument 0 of every host callee,
    /// so one signature means one thing. Under the AMBIENT ABI that site pushes
    /// nothing, and only a host callee that declares a receiver gets one —
    /// handing it to the rest shifts their arguments.
    fn place_receiver(
        &self,
        host_callee: bool,
        receiver: Value,
        args: &[Value],
    ) -> Option<Vec<Value>> {
        if !host_callee || self.receiver_is_parameter() {
            return None;
        }
        let mut all = Vec::with_capacity(args.len() + 1);
        all.push(receiver);
        all.extend_from_slice(args);
        Some(all)
    }

    /// Capture `i` of a BOUND host function (`__bound_args`).
    ///
    /// The captures sit after the receiver slot, which the VM places first for
    /// a host callee under [`crate::chunk::ReceiverAbi::Parameter`]. Reading
    /// them at a fixed index is correct only under the ambient binding, and
    /// silently reads the receiver as capture 0 otherwise.
    pub fn capture(&self, args: &[Value], i: usize) -> Value {
        args.get(i).cloned().unwrap_or(Value::Undefined)
    }

    pub fn user_args<'v>(&self, args: &'v [Value], captures: usize) -> &'v [Value] {
        &args[captures.min(args.len())..]
    }

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
            self.slot_get(name)
        }
    }

    /// Resolve a name through the index slot, then read the value slot.
    unsafe fn slot_get(&self, name: &str) -> Value {
        if self.globals_slot.is_null() || self.global_index_slot.is_null() {
            return Value::Undefined;
        }
        unsafe {
        let vals: &Vec<Value> = &*self.globals_slot;
        match (*self.global_index_slot).get(name) {
            Some(&i) => vals.get(i as usize).cloned().unwrap_or(Value::Undefined),
            None => Value::Undefined,
        }
        }
    }

    /// Write by name, allocating an index if the name is new.
    unsafe fn slot_set(&self, name: &str, value: Value) {
        if self.globals_slot.is_null() || self.global_index_slot.is_null() {
            return;
        }
        unsafe {
        let vals: &mut Vec<Value> = &mut *self.globals_slot;
        let index: &mut HashMap<String, u32> = &mut *self.global_index_slot;
        // ⚠ A GLOBAL IS NEVER CREATED HERE.
        //
        // This used to mint an entry — `index.insert(name); vals.push(Null)` —
        // for any unknown name. That is not something WASM can express: a
        // module's global index space is fixed at instantiation, so a global
        // conjured by name at runtime cannot exist on a stock engine, and any
        // program depending on one would not run there.
        //
        // Every global a module actually uses is assigned a `globalidx` by the
        // compiler's global-table normalisation pass before the module is
        // handed to the VM. A name that misses here was never declared, so
        // writing it would be writing to a global the module does not have.
        let Some(&i) = index.get(name) else { return };
        let idx = i as usize;
        if idx >= vals.len() {
            vals.resize(idx + 1, Value::Null);
        }
        vals[idx] = value;
        }
    }

    /// Write a VM global by name — counterpart of [`Self::get_global`],
    /// used by host modules that must bind calling-convention globals
    /// (e.g. `__js_new_target` around [[Construct]] dispatch).
    pub fn set_global(&mut self, name: &str, value: Value) {
        unsafe {
            self.slot_set(name, value);
        }
    }

    /// Read the receiver the host will hand the next callback.
    /// Returns `Undefined` when no binding exists.
    pub fn current_js_this(&self) -> Value {
        if self.host_receiver_slot.is_null() {
            return Value::Undefined;
        }
        unsafe { (*self.host_receiver_slot).clone() }
    }

    /// Update the current JS receiver binding in whichever channel this module
    /// has. See [`HostContext::host_receiver_slot`] for why there are two.
    pub fn set_js_this(&mut self, value: Value) {
        if !self.host_receiver_slot.is_null() {
            unsafe {
                *self.host_receiver_slot = value;
            }
        }
    }

    /// Does this module place a receiver at argument 0?
    ///
    /// ⛔ THIS ASKS THE MODULE'S ABI. It is a property of how the module was
    /// compiled, never of what its global index space happens to contain.
    fn module_uses_host_receiver_channel(&self) -> bool {
        self.module_receiver_abi == crate::chunk::ReceiverAbi::Parameter
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

    /// Create a stream. Returns the stream Value (for guest code) and its ID
    /// (for host push/close).
    ///
    /// The Value is an i32 READABLE-END HANDLE, per `CanonicalABI.md`
    /// §HandleTable: a `stream<u8>` in a function signature is lifted and
    /// lowered as an index into the component instance's handle table, in BOTH
    /// directions. The guest already lowered that way — it passes an i32 to
    /// `wasi:cli/stdout.write-via-stream` — but every host producer handed back
    /// an `ObjectKind::Stream` marker instead, which `canon stream.read` cannot
    /// accept: it resolves its argument through the handle table and demands a
    /// `ReadableStreamEnd`. So `read-via-stream`, both `consume-body`s and
    /// socket receive were all returning a stream no conforming guest could
    /// read. That was the whole read side of the ABI.
    ///
    /// Only the READABLE end gets a handle, unlike `canon stream.new` which
    /// mints both. That is not an omission: the host keeps the raw
    /// `stream_id` and writes through [`stream_push`]/[`stream_close`], so
    /// there is no writable end for the guest to name.
    ///
    /// Falls back to the object form when no event loop is attached, which is
    /// also the only case where there is no handle table to lift into.
    pub fn create_stream(&mut self) -> (Value, u64) {
        self.create_stream_of(None)
    }

    /// Mint an `own<T>` handle for a host-held value — the host counterpart of
    /// `canon resource.new`.
    ///
    /// Without this a host could not hand a guest a RESOURCE at all: `resource.new`
    /// is a bytecode built-in, so only guest code could mint one, and every host
    /// that wanted to return a resource returned a plain object pretending to be
    /// one. That works right up until the value has to cross the canonical ABI —
    /// `own<T>` lowers as an i32 index into the handle table, and an object
    /// lowers as `as_i32()` of nothing. It is why `tcp-socket.listen`'s
    /// `stream<tcp-socket>` could not declare its element type.
    ///
    /// The `Value` is the resource's REPRESENTATION and stays private to the
    /// component that owns it; the handle is the only thing that crosses.
    ///
    /// `None` when there is no handle table to lift into, which is the same
    /// condition under which [`create_stream`] falls back to its object form.
    pub fn create_own_resource(&mut self, type_id: u32, rep: Value) -> Option<Value> {
        unsafe {
            if self.handle_table_slot.is_null() {
                return None;
            }
            let handle = (*self.handle_table_slot)
                .insert(crate::handle_table::HandleEntry::OwnedResource {
                    type_id,
                    value: rep,
                });
            Some(Value::I32(handle as i32))
        }
    }

    /// [`create_stream`] for a `stream<T>` whose `T` is not `u8`.
    ///
    /// The element type has to be recorded HERE, at creation, because
    /// `canon stream.read`'s signature carries only `(handle, ptr, n)` — the
    /// `$stream_t` immediate lives on the canon definition, not in the call.
    /// A stream that did not record its own element type could only ever be
    /// read as bytes, which is precisely why `stream<directory-entry>` and
    /// `stream<tcp-socket>` were unreadable.
    pub fn create_stream_of(
        &mut self,
        elem: Option<crate::component::ValType>,
    ) -> (Value, u64) {
        use crate::value::{Object, ObjectKind};
        if let Some(ref el) = self.event_loop {
            let id = el.borrow_mut().create_stream_of(elem);
            unsafe {
                if !self.handle_table_slot.is_null() {
                    let handle = (*self.handle_table_slot).insert(
                        crate::handle_table::HandleEntry::ReadableStreamEnd(
                            crate::handle_table::StreamEnd::new(id),
                        ),
                    );
                    return (Value::I32(handle as i32), id);
                }
            }
            let obj = Object {
                properties: indexmap::IndexMap::new(),
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
                    .immediate
                    .push_back(crate::event_loop::Task::ResumeFiber(fiber));
            }
        }
    }

    /// Name the host function that produces more elements for `stream_id`.
    ///
    /// For a stream whose elements arrive over TIME — `wasi:sockets`'
    /// `listen()`, one element per inbound connection — the call that created
    /// the stream cannot fill it, because a host function returns once. The
    /// producer is called by a reader that is about to park, so the host gets
    /// to `accept` at the moment the guest asks. Without it, such a stream can
    /// only ever hold what was already pending when it was created, which for
    /// a listener is reliably nothing.
    ///
    /// The producer receives the stream id as its only argument and pushes
    /// through `stream_push` / `stream_close` like any other host code. It is
    /// dropped automatically when the stream closes.
    pub fn set_stream_producer(&mut self, stream_id: u64, module: &str, name: &str) {
        if let Some(ref el) = self.event_loop {
            el.borrow_mut().set_stream_producer(stream_id, module, name);
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
                Some(crate::handle_table::HandleEntry::ReadableStreamEnd(e)) => Some(e.id),
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
            host_receiver_slot: std::ptr::null_mut(),
            module_receiver_abi: crate::chunk::ReceiverAbi::Ambient,
            exit_slot: std::ptr::null_mut(),
            exit_code_slot: std::ptr::null_mut(),
            globals_slot: std::ptr::null_mut(),
            global_index_slot: std::ptr::null_mut(),
            stack_slot: std::ptr::null(),
            handle_table_slot: std::ptr::null_mut(),
            shared_memory_slot: std::ptr::null(),
            type_registry_slot: std::ptr::null(),
            call_tag_registry_slot: std::ptr::null(),
            func_call_tags_slot: std::ptr::null(),
            chunks_slot: std::ptr::null(),
        }
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

/// Length of a thread's `context.get`/`context.set` storage array.
///
/// `CanonicalABI.md class Thread`: `storage: tuple[int,int]` /
/// `self.storage = [0,0]`. `Explainer.md:1679` states the matching validation
/// rule: "Validation currently restricts `i` to be less than 2 and `T` to be
/// `i32`". Named rather than inlined so the bound and the initial length can
/// never drift apart — they are the same spec fact.
pub const CONTEXT_STORAGE_SLOTS: usize = 2;

/// Component Model canonical built-ins the VM implements natively. The CM
/// defines these as `(core func)` DEFINITIONS a component wires into a core
/// instance's imports — functions, not instructions — so a core module
/// reaches them via spec `call` on an import under the "canon" module,
/// exactly like the jspi/wasi-threads VM-implemented imports. The 0xF0
/// instruction prefix that used to carry them is being retired.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum CanonBuiltin {
    /// `(canon lift ...)` — typeidx rides the stack (top), above the value.
    Lift,
    /// `(canon lower ...)` — typeidx rides the stack (top), above the value.
    Lower,
    TaskReturn,
    TaskCancel,
    SubtaskCancel,
    SubtaskDrop,
    WaitableSetNew,
    WaitableSetWait,
    WaitableSetPoll,
    WaitableJoin,
    StreamNew,
    StreamRead,
    StreamWrite,
    StreamCancelRead,
    StreamCancelWrite,
    StreamDropReadable,
    StreamDropWritable,
    FutureNew,
    FutureRead,
    FutureWrite,
    FutureCancelRead,
    FutureCancelWrite,
    FutureDropReadable,
    FutureDropWritable,
    WaitableSetDrop,
    ThreadYield,
    /// 🧵 `thread.index` — the current thread's index in the instance table.
    ThreadIndex,
    /// 🧵 `thread.resume-later` — mark a suspended thread ready to run.
    ThreadResumeLater,
    /// 🧵 `thread.suspend-then-resume` — park me, run them.
    ThreadSuspendThenResume,
    /// 🧵 `thread.yield-then-resume` — stay runnable, run them.
    ThreadYieldThenResume,
    /// 🧵 `thread.suspend-then-promote` — park me, run them IF ready.
    ThreadSuspendThenPromote,
    /// 🧵 `thread.yield-then-promote` — stay runnable, run them IF ready.
    ThreadYieldThenPromote,
    /// 🧵 `thread.new-indirect` — create a suspended thread over a funcref.
    ThreadNewIndirect,
    /// 🧵 `thread.suspend` — block the current thread with no `switch_to`.
    ThreadSuspend,
    /// 🧵② `thread.spawn-ref` — `new-ref` + `resume-later`, fused.
    ThreadSpawnRef,
    /// 🧵② `thread.spawn-indirect` — `new-indirect` + `resume-later`, fused.
    ThreadSpawnIndirect,
    /// 🧵② `thread.available-parallelism`.
    ThreadAvailableParallelism,
    ResourceNew,
    ResourceRep,
    ResourceDrop,
    /// 📝 `error-context.new` — `CanonicalABI.md:5147`.
    ErrorContextNew,
    /// 📝 `error-context.debug-message` — `CanonicalABI.md:5189`.
    ErrorContextDebugMessage,
    /// 📝 `error-context.drop` — `CanonicalABI.md:5215`.
    ErrorContextDrop,
    BackpressureInc,
    BackpressureDec,
    ContextGet,
    ContextSet,
}

impl CanonBuiltin {
    /// CM canonical names (Binary.md spellings) → builtin.
    /// Split `"future.read@2"` into `("future.read", Some(2))`.
    ///
    /// `@` and not `:` — a colon already separates module from function in the
    /// `host:<module>:<fn>` callee spelling, so `host:canon:future.read:1`
    /// would rsplit into module `canon:future.read` and function `1`. Two
    /// encodings sharing a delimiter is how that kind of bug hides.
    ///
    /// A bare name keeps `None`, which means "no type immediate was declared".
    /// That is honest rather than defaulted: a built-in that NEEDS the type
    /// (`future.{read,write}`, or a `stream<T>` where T is not a byte) refuses
    /// instead of moving a guessed number of bytes, because a canonical ABI
    /// that is quietly wrong about layout corrupts a peer's memory.
    pub fn split_type_immediate(name: &str) -> (&str, Option<u32>) {
        match name.rsplit_once('@') {
            Some((bare, idx)) => match idx.parse::<u32>() {
                Ok(n) => (bare, Some(n)),
                // Not a number — the colon belongs to the name itself.
                Err(_) => (name, None),
            },
            None => (name, None),
        }
    }


    /// The Binary.md spelling of this built-in — the inverse of [`Self::by_name`].
    ///
    /// Exists so a trap can name the ROW it belongs to (`canon task.return:
    /// ...`) rather than a Rust identifier. Derived from the same 33 rows as
    /// `by_name`, so a new built-in that forgets one side fails to compile.
    pub fn spec_name(self) -> &'static str {
        match self {
            Self::Lift => "lift",
            Self::Lower => "lower",
            Self::TaskReturn => "task.return",
            Self::TaskCancel => "task.cancel",
            Self::SubtaskCancel => "subtask.cancel",
            Self::SubtaskDrop => "subtask.drop",
            Self::WaitableSetNew => "waitable-set.new",
            Self::WaitableSetWait => "waitable-set.wait",
            Self::WaitableSetPoll => "waitable-set.poll",
            Self::WaitableJoin => "waitable.join",
            Self::StreamNew => "stream.new",
            Self::StreamRead => "stream.read",
            Self::StreamWrite => "stream.write",
            Self::StreamCancelRead => "stream.cancel-read",
            Self::StreamCancelWrite => "stream.cancel-write",
            Self::StreamDropReadable => "stream.drop-readable",
            Self::StreamDropWritable => "stream.drop-writable",
            Self::FutureNew => "future.new",
            Self::FutureRead => "future.read",
            Self::FutureWrite => "future.write",
            Self::FutureCancelRead => "future.cancel-read",
            Self::FutureCancelWrite => "future.cancel-write",
            Self::FutureDropReadable => "future.drop-readable",
            Self::FutureDropWritable => "future.drop-writable",
            Self::WaitableSetDrop => "waitable-set.drop",
            Self::ThreadYield => "thread.yield",
            Self::ThreadIndex => "thread.index",
            Self::ThreadResumeLater => "thread.resume-later",
            Self::ThreadSuspendThenResume => "thread.suspend-then-resume",
            Self::ThreadYieldThenResume => "thread.yield-then-resume",
            Self::ThreadSuspendThenPromote => "thread.suspend-then-promote",
            Self::ThreadYieldThenPromote => "thread.yield-then-promote",
            Self::ThreadNewIndirect => "thread.new-indirect",
            Self::ThreadSuspend => "thread.suspend",
            Self::ThreadSpawnRef => "thread.spawn-ref",
            Self::ThreadSpawnIndirect => "thread.spawn-indirect",
            Self::ThreadAvailableParallelism => "thread.available-parallelism",
            Self::ErrorContextNew => "error-context.new",
            Self::ErrorContextDebugMessage => "error-context.debug-message",
            Self::ErrorContextDrop => "error-context.drop",
            Self::ResourceNew => "resource.new",
            Self::ResourceRep => "resource.rep",
            Self::ResourceDrop => "resource.drop",
            Self::BackpressureInc => "backpressure.inc",
            Self::BackpressureDec => "backpressure.dec",
            Self::ContextGet => "context.get",
            Self::ContextSet => "context.set",
        }
    }

    pub fn by_name(name: &str) -> Option<Self> {
        Some(match name {
            "lift" => Self::Lift,
            "lower" => Self::Lower,
            "task.return" => Self::TaskReturn,
            "task.cancel" => Self::TaskCancel,
            "subtask.cancel" => Self::SubtaskCancel,
            "subtask.drop" => Self::SubtaskDrop,
            "waitable-set.new" => Self::WaitableSetNew,
            "waitable-set.wait" => Self::WaitableSetWait,
            "waitable-set.poll" => Self::WaitableSetPoll,
            "waitable.join" => Self::WaitableJoin,
            "stream.new" => Self::StreamNew,
            "stream.read" => Self::StreamRead,
            "stream.write" => Self::StreamWrite,
            "stream.cancel-read" => Self::StreamCancelRead,
            "stream.cancel-write" => Self::StreamCancelWrite,
            "stream.drop-readable" => Self::StreamDropReadable,
            "stream.drop-writable" => Self::StreamDropWritable,
            "future.new" => Self::FutureNew,
            "future.read" => Self::FutureRead,
            "future.write" => Self::FutureWrite,
            "future.cancel-read" => Self::FutureCancelRead,
            "future.cancel-write" => Self::FutureCancelWrite,
            "future.drop-readable" => Self::FutureDropReadable,
            "future.drop-writable" => Self::FutureDropWritable,
            "waitable-set.drop" => Self::WaitableSetDrop,
            "thread.yield" => Self::ThreadYield,
            "thread.index" => Self::ThreadIndex,
            "thread.resume-later" => Self::ThreadResumeLater,
            "thread.suspend-then-resume" => Self::ThreadSuspendThenResume,
            "thread.yield-then-resume" => Self::ThreadYieldThenResume,
            "thread.suspend-then-promote" => Self::ThreadSuspendThenPromote,
            "thread.yield-then-promote" => Self::ThreadYieldThenPromote,
            "thread.new-indirect" => Self::ThreadNewIndirect,
            "thread.suspend" => Self::ThreadSuspend,
            "thread.spawn-ref" => Self::ThreadSpawnRef,
            "thread.spawn-indirect" => Self::ThreadSpawnIndirect,
            "thread.available-parallelism" => Self::ThreadAvailableParallelism,
            "error-context.new" => Self::ErrorContextNew,
            "error-context.debug-message" => Self::ErrorContextDebugMessage,
            "error-context.drop" => Self::ErrorContextDrop,
            "resource.new" => Self::ResourceNew,
            "resource.rep" => Self::ResourceRep,
            "resource.drop" => Self::ResourceDrop,
            "backpressure.inc" => Self::BackpressureInc,
            "backpressure.dec" => Self::BackpressureDec,
            "context.get" => Self::ContextGet,
            "context.set" => Self::ContextSet,
            _ => return None,
        })
    }
}

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
    WasiThreadSpawn,
    /// Component Model canonical built-in (module "canon"), VM-implemented —
    /// see [`CanonBuiltin`]. Args/results ride the operand stack; the
    /// builtin body pops what it needs.
    /// A canonical built-in, with the type immediate its `canon`
    /// definition carried (`None` when the import declared none).
    Canon(CanonBuiltin, Option<u32>),
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

/// Order-insensitive equality of two import tables (compile emits imports in
/// HashMap order, which varies run to run). A genuine edit changes the SET.
fn imports_equal_as_set(a: &[crate::chunk::Import], b: &[crate::chunk::Import]) -> bool {
    if a.len() != b.len() {
        return false;
    }
    let mut av: Vec<(&str, &str)> = a
        .iter()
        .map(|i| (i.module.as_str(), i.name.as_str()))
        .collect();
    let mut bv: Vec<(&str, &str)> = b
        .iter()
        .map(|i| (i.module.as_str(), i.name.as_str()))
        .collect();
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
    Iterator,
}

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
    pub(crate) arity: u8,
}

/// A `func_switch` — the Overview's alternative way of defining a function:
///
/// > `func_switch ($call_tag $func)* $func_switch?`, specifying essentially a
/// > switch statement that calls a `$func` if the given call tag matches the
/// > corresponding `$call_tag`. If there is no corresponding call tag, then if
/// > `$func_switch` is specified the call tag and arguments are forwarded to
/// > it, otherwise the fall-back handler of the call tag is (tail) called.
///
/// This is what makes interface-method dispatch cheap in the proposal's
/// motivating example: one `funcref` per descriptor slot, switching on the tag,
/// with unmatched tags forwarded to the superclass's `funcref` so a subclass
/// need not re-list everything it inherits.
#[derive(Debug, Clone, Default)]
pub(crate) struct FuncSwitch {
    /// `($call_tag $func)*` — matched in declaration order.
    pub(crate) arms: Vec<(u32, usize)>,
    /// The trailing `$func_switch?` — another func_switch to forward to.
    pub(crate) forward: Option<usize>,
}

/// A CALL tag (`proposals/call-tags`) — not an exception tag.
///
/// The identity that matters is the tag's INDEX, not its signature.
/// `call_tag.canon $functype` interns one tag per signature; `call_tag.new
/// $functype $func?` mints a fresh one over the *same* signature, so two funcs
/// that are structurally identical — which, after GC type canonicalisation,
/// means genuinely the same type — remain distinguishable at the call. That is
/// the property a structural type system cannot give you and the reason the
/// proposal exists.
#[derive(Debug, Clone)]
pub(crate) struct CallTagDef {
    /// Spelling, for traps and disassembly.
    pub(crate) debug_name: String,
    /// The tag's SHAPE — parameter count and result count. Kept because every
    /// arity check at a call site is expressed in it.
    pub(crate) params: u8,
    pub(crate) results: u8,
    /// The DECLARED functype spelling (`"i32->i32"`), empty when the producer
    /// supplied none.
    ///
    /// ⛔ A SHAPE IS NOT A FUNCTYPE, and canonical tags used to be interned on
    /// the shape alone: `canonical_call_tags: HashMap<(u8,u8), u32>`. So
    /// `call_tag.canon [i32]->[i32]` and `call_tag.canon [f64]->[f64]` were ONE
    /// tag, and an `i32`-shaped funcref answered the `f64` canonical tag —
    /// measured, accepted, wrong. The Overview derives the canonical tag *of a
    /// functype*; two functypes are two tags. Since `call_indirect $table
    /// $functype` is shorthand for `call_with_tag (call_tag.canon $functype)`,
    /// interning on the shape reduced the Security property to arity-safety.
    pub(crate) signature: String,
    /// Fall-back handler (`call_tag.new $functype $func`). When a `funcref`
    /// does not handle this tag, the Overview tail-calls this with the same
    /// arguments "but replacing the call-tag value with the value of the
    /// current `funcref`". `None` — which is always the case for a canonical
    /// tag — means an unhandled tag TRAPS.
    pub(crate) fallback: Option<Value>,
    /// True for a tag produced by `call_tag.canon`. Canonical tags never carry
    /// a fall-back, and the Overview says so: "For canonical call tags, the
    /// answer is simply that the program traps."
    pub(crate) canonical: bool,
}

/// Exception handler entry — pushed per catch clause by `try_table`,
/// popped (as a group) by TRY_END or on catch.
#[derive(Debug, Clone)]
pub(crate) struct ExceptionHandler {
    /// Spec `labelidx` — the RELATIVE BLOCK DEPTH this clause branches to,
    /// resolved against the label stack as it stood when the handler was
    /// installed (`label_depth` below), exactly as `br` resolves its operand.
    ///
    /// This was a byte offset (`catch_ip`), which is a different quantity: it
    /// cannot survive import from a conforming `.wasm`, and it truncated past a
    /// 64KB try body, resuming the handler mid-instruction. See
    /// `Chunk::emit_try_table_clauses`.
    pub(crate) catch_label: u16,
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
    pub(crate) group: u64,
}

/// How a host function relates to a Component Model resource.
///
/// A DOM node, a file handle and a database connection are all resources: the
/// host owns the thing, the guest holds a handle. Saying which resource a
/// function belongs to — and whether it takes `self` — is what makes
/// `append-child` a method on `node` rather than a free function that happens
/// to take a node-shaped value first.
#[derive(Debug, Clone, PartialEq)]
pub struct ResourceBinding {
    /// The resource type this function is a method on (`"node"`).
    pub resource: String,
    /// A constructor yields a fresh handle; a destructor consumes one.
    pub kind: ResourceMemberKind,
    /// `borrow<self>` (read-only, not dropped) vs `own<self>` (consuming).
    /// Every DOM operation borrows.
    pub borrows_self: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ResourceMemberKind {
    Constructor,
    Method,
    Static,
    Destructor,
}

/// Everything a host function DECLARES about itself.
///
/// One record, not a parameter list. Each dimension is a field with a default,
/// so a new one — effects, capabilities, whatever comes next — is additive and
/// no existing registration moves. That is the property this type exists for:
/// the registration API is changed once and then never again.
///
/// `sig: None` is the honest default. The registry has never carried
/// signatures, so `mount_host_exports` mounts every leaf with `arity: None`;
/// a declaration is what upgrades a `(module, func)` string pair into a typed
/// interface member, which is what `namespaceplan.md` requires of a leaf.
pub struct HostFnDecl {
    pub module: String,
    pub name: String,
    pub call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
    /// The Component Model signature. `None` = undeclared, not "no arguments".
    pub sig: Option<crate::component::FuncSig>,
    /// The resource this function is a member of, when it is one.
    pub resource: Option<ResourceBinding>,
    /// Does this function's type have a RECEIVER as parameter 0?
    ///
    /// ⛔ A PROPERTY OF THE CALLEE'S TYPE, NEVER OF THE CALL. `[[Call]]`
    /// (§10.2.1) passes a thisArgument to every function, but a host builtin
    /// either declares a slot for it (`ecma:array.join` is
    /// `(receiver, separator)` — `array_of(args, 0)` opens the receiver) or has
    /// no receiver at all (`encodeURIComponent(s)`). Asking the CALL instead
    /// gave one function two shapes — one argument when called directly, a
    /// receiver and then one when handed to `map` — which is what
    /// `receiver_argc`, `user_args` and `capture` existed to paper over.
    ///
    /// Default `true`: the overwhelming majority of registrations are methods
    /// and have always read their receiver at argument 0, so a registration
    /// that says nothing keeps exactly the shape it had.
    pub takes_receiver: bool,
}

impl HostFnDecl {
    /// An undeclared registration — exactly what `register_host_fn` produces.
    pub fn new(
        module: impl Into<String>,
        name: impl Into<String>,
        call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
    ) -> Self {
        Self {
            module: module.into(),
            name: name.into(),
            call,
            sig: None,
            resource: None,
            takes_receiver: true,
        }
    }

    /// Declare that this function's type has NO receiver parameter — a free
    /// function, not a method. See [`HostFnDecl::takes_receiver`].
    pub fn without_receiver(mut self) -> Self {
        self.takes_receiver = false;
        self
    }

    pub fn with_sig(mut self, sig: crate::component::FuncSig) -> Self {
        self.sig = Some(sig);
        self
    }

    /// Declare this function a method on a resource, borrowing `self`.
    pub fn method_on(mut self, resource: impl Into<String>) -> Self {
        self.resource = Some(ResourceBinding {
            resource: resource.into(),
            kind: ResourceMemberKind::Method,
            borrows_self: true,
        });
        self
    }

    pub fn resource_member(mut self, binding: ResourceBinding) -> Self {
        self.resource = Some(binding);
        self
    }
}

/// A language-agnostic bytecode virtual machine.
///
/// The VM has no built-in functions or language-specific semantics.
/// The host (compiler runtime) registers native functions via `register_host_fn`
/// and sets up globals before calling `run`.
pub struct VM {
    /// The thisArgument a host function has set for the callback it is about to
    /// make — see [`HostContext::host_receiver_slot`] for why the module global
    /// cannot carry it once the receiver becomes a parameter.
    pub(crate) host_receiver: Value,
    /// Set by `invoke_callback` immediately before it dispatches, and CONSUMED
    /// by the next `call_value_inner`. One-shot on purpose: a call the callee
    /// itself makes is ordinary bytecode and must NOT inherit "this came from
    /// host plumbing".
    pub(crate) host_originated_call: bool,
    /// Set only by `invoke_with_receiver`: the caller has already placed the
    /// receiver at argument 0, so `invoke_callback` must not add another.
    pub(crate) suppress_receiver_prepend: bool,
    pub chunks: Vec<Chunk>,
    pub(crate) frames: Vec<CallFrame>,
    pub(crate) stack: Vec<Value>,
    /// Module globals, indexed by `globalidx` — WASM's model. `GLOBAL_GET`/
    /// `GLOBAL_SET` operands index THIS directly; no string is consulted on
    /// the execution path.
    pub globals: Vec<Value>,
    /// Whether each slot has ever been ASSIGNED. A `GlobalInit` applies only to
    /// a slot nobody has written, which is not the same question as "is the
    /// value null/undefined": a program may legitimately store undefined in a
    /// global, and a host may pre-install a native over a helper. Overloading
    /// the value to answer this either skips every helper install (Undefined
    /// placeholder) or clobbers real values (widened guard). Both were tried.
    pub globals_assigned: Vec<bool>,
    /// name → globalidx, consulted only at INSTANTIATION (binding host and
    /// imported globals, which WASM also resolves by name at instantiate time)
    /// and by the debugger. Not one fact in two homes: the vector is the
    /// storage, this is a resolver used before execution starts.
    pub global_index: HashMap<String, u32>,
    pub(crate) open_upvalues: Vec<Arc<Mutex<Upvalue>>>,
    pub(crate) host_fns: Vec<HostFn>,
    /// Parallel to `host_fns`: does that function's TYPE have a receiver as
    /// parameter 0? See [`HostFnDecl::takes_receiver`]. Defaults to `true`, so
    /// a plain `register_host_fn` keeps the shape it has always had.
    pub(crate) host_fn_takes_receiver: Vec<bool>,
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
    /// CALL tags (`proposals/call-tags`) — a different entity from the EH tags
    /// above, which name exception types. Identity IS the index, which is the
    /// whole point: `call_tag.new` mints a FRESH identity over a signature that
    /// already has one, so two funcs with the same structural type stay
    /// distinguishable. Canonical tags (`call_tag.canon`) are interned by
    /// signature so the same functype always yields the same tag.
    pub(crate) call_tags: Vec<CallTagDef>,
    /// Canonical-tag interning: `(params, results)` → index in `call_tags`.
    pub(crate) canonical_call_tags: HashMap<String, u32>,
    /// `func_switch` definitions, keyed by the chunk index standing in for the
    /// definition. The Overview makes this "an alternative to `func`": it has
    /// no type, cannot be called directly, and exists to be reached by
    /// `call_with_tag` through a `funcref`.
    pub(crate) func_switches: HashMap<usize, FuncSwitch>,
    /// Per-chunk resolution: chunk-local call-tag index → VM tag id, the
    /// sibling of `chunk_tag_maps` for exception tags.
    pub(crate) chunk_call_tag_maps: Vec<Vec<u32>>,
    /// Call-tag name → id, so one name is one entity module-wide.
    pub(crate) call_tag_registry: HashMap<String, u32>,
    /// Call-tag validation failures found while resolving declarations.
    ///
    /// The spec validates at MODULE VALIDATION time; this VM resolves
    /// declarations lazily as chunks are installed, so the findings are
    /// collected here and surfaced by the first tagged call in the module.
    /// The program still cannot execute a call under an invalid declaration,
    /// which is the property that matters — it just learns at first use.
    pub(crate) call_tag_errors: Vec<String>,
    /// Which call tags a FUNC DEFINITION handles, keyed by chunk index.
    ///
    /// Keyed by the definition rather than stored on `Function` because the
    /// Overview hangs it off the `func` ("When defining a `func` … one can
    /// optionally specify `(call_tag $call_tag*)`"), so every closure over the
    /// same func shares it. Absent means the default the Overview states: the
    /// funcref handles exactly the canonical tag of its own signature.
    pub(crate) func_call_tags: HashMap<usize, Vec<u32>>,
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
    /// Which chunk owns the IMPORT TABLE each chunk's `CALL <idx>` operands
    /// index — parallel to `chunks`, same shape as `chunk_type_base`.
    ///
    /// ⛔ AN IMPORT INDEX IS MODULE-RELATIVE, AND ONLY ONE CHUNK KEPT THE
    /// TABLE. `link.rs` unifies a module's imports into `chunks[0].imports`
    /// and CLEARS every other chunk's, so `resolve_chunk_import` found nothing
    /// for a function body and fell through to the VM-wide positional
    /// `import_table` — which `run_linked_impl` CLEARS AND REBUILDS per
    /// program unit. A linked unit's baked `CALL 38` therefore indexed a LATER
    /// unit's table and silently called a DIFFERENT host function
    /// (`js-prototypes:configureAll` arriving at `js-string:compare`, returning
    /// a plausible number rather than an error).
    ///
    /// This restores the invariant `merge_global_table` already states for its
    /// own table: "`chunk.globals` describes `chunk.code`'s operands. Anything
    /// that rewrites those operands owns re-stating it here." Nothing rewrites
    /// import operands across units, so each chunk records the module whose
    /// table its operands were remapped into, and resolution goes back to
    /// module+NAME rather than a shared position.
    pub(crate) chunk_import_owner: Vec<usize>,
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
    /// The last `resume_fiber` re-parked instead of running.
    ///
    /// A fiber parked in a synchronous `canon stream.read` can be woken and
    /// find the data already taken, so it parks again WITHOUT executing. It
    /// therefore has no completion, and recording the placeholder as one would
    /// make a top-level-await program answer `null` (see
    /// `last_fiber_completion`, which `run()` surfaces as the final value).
    pub(crate) resume_reparked: bool,
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
    /// Where this VM currently lives, in a cell the VM does NOT own inline.
    ///
    /// `callback_invoker` is built once and reused, so whatever address it
    /// captures it keeps. A VM is MOVED after its first host call — an embedder
    /// hands it on by value and parks it in an `Rc` (`cli::run` → `launch_gui` →
    /// `Rc::new(RefCell::new(vm))`) — and the captured address then named the
    /// moved-out-of copy: a bitwise duplicate sharing every heap pointer with
    /// the live VM. A callback running through it grew `label_stack`, which
    /// reallocated and freed the buffer the LIVE VM still pointed at, and the
    /// program aborted at quit in `free` (`pointer being freed was not
    /// allocated`) dropping that same `Vec`.
    ///
    /// The cell is behind a `Box` precisely because moving the VM copies the
    /// Box's pointer and leaves the pointee where it is, so its address is the
    /// one thing here that survives a move. The closure captures THAT, and
    /// `get_invoker` republishes `self` into it on every host call, so the
    /// callback always reaches the VM where it now lives.
    pub(crate) invoker_vm_slot: Box<std::cell::Cell<*mut VM>>,
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
    /// Component-level types the canonical built-ins are parameterised by.
    ///
    /// A `canon` definition is NOT one function — `(canon future.read $t $opts
    /// (core func $f))` produces a DISTINCT core func per instantiation, and
    /// `$t` is how `future.read` knows how many bytes one element occupies
    /// (its signature is `(handle, ptr)` with no count). Registering imports by
    /// bare name gave every use of a built-in the same identity and no type at
    /// all, which is why `future.{read,write}` could not be implemented while
    /// `stream<u8>` could: a byte needs no type to measure.
    ///
    /// An import may carry the index as a `name@idx` suffix, so
    /// `canon`/`future.read@2` and `canon`/`future.read@5` are different core
    /// funcs over the same built-in — the distinctness the spec requires,
    /// expressed in the one channel a core import has.
    pub canon_types: Vec<Option<crate::component::ValType>>,
    /// The component's FUNCTION type index space — `canon lift`'s
    /// `ft:<typeidx>`.
    ///
    /// Separate from `canon_types` because they are different index spaces.
    /// `stream.read`'s `$t` names a VALUE type; `canon lift`'s `$ft` names a
    /// FUNCTION type. One `Vec` serving both is the `GLOBAL_GET` defect —
    /// a single integer meaning two things depending on who reads it.
    pub canon_functypes: Vec<Option<crate::canon_def::CanonFuncType>>,
    /// The component FUNCTION index space — comp funcidx -> defining canonidx.
    ///
    /// `canon lower`'s `$callee` indexes this. A component function defined
    /// HERE (by a `canon lift`) needs no linker to call: the linker is for
    /// component functions that arrive as IMPORTS.
    pub component_funcs: Vec<Option<u32>>,
    /// The module's canon section — `Binary.md` §"Canonical Definitions".
    /// A canon import resolves to a row here at link time, which is the spec's
    /// instantiation-time capture of `$callee` / `$opts` / `$ft`.
    pub canon_defs: Vec<crate::canon_def::CanonDef>,
    /// The component instance — `CanonicalABI.md class ComponentInstance`.
    /// Owns `may_enter`/`may_leave`, the backpressure counter and the thread
    /// table. One instance, so `current_instance()` is always this.
    pub cm_instance: crate::cm_instance::ComponentInstance,
    /// `current_thread()` — index into `cm_instance.threads` of the thread
    /// executing right now, or `None` outside any lifted call.
    ///
    /// The spec's implicit thread is spawned by `canon_lift`; there is no other
    /// way for one to exist, which is why every 🧵 built-in was unreachable
    /// while `canon lift` was a stub.
    pub current_thread: Option<u32>,
    /// The `$t` immediate of the canonical built-in currently executing, set
    /// by the dispatch arm just before the call. Carried on the VM rather than
    /// threaded through `exec_canon_builtin` because only the typed copies read
    /// it, and every other built-in would have to ignore an extra parameter.
    pub(crate) canon_type_immediate: Option<u32>,
    /// Active CM3 tasks (keyed by task ID). Each async export invocation creates one.
    pub cm_tasks: Vec<crate::cm_task::CMTask>,
    /// Next CM3 task ID.
    #[allow(dead_code)]
    pub(crate) next_cm_task_id: u32,
    /// Waitable set registry.
    pub waitable_sets: crate::waitable::WaitableRegistry,
}

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
    /// Both halves: the values AND the name→index resolver, so a script-added
    /// global vanishes on reset instead of leaving a live index behind.
    globals: Vec<Value>,
    global_index: HashMap<String, u32>,
    memory: Vec<u8>,
    extra_memories: Vec<Vec<u8>>,
    wasm_tables: Vec<Vec<Value>>,
    dropped_data: HashSet<u32>,
    dropped_elems: HashSet<u32>,
    active_memory: usize,
    handle_table: crate::handle_table::HandleTable,
    waitable_sets: crate::waitable::WaitableRegistry,
    cm_tasks: Vec<crate::cm_task::CMTask>,
    /// The component model's CANON SECTION and its type spaces.
    ///
    /// ⛔ These were the one piece of component state the snapshot did not
    /// capture, and `merge_canon_section` REJECTS a second, different section
    /// rather than overwriting one — "this program declares two different canon
    /// sections; each numbers its canonidx space from zero, so they cannot be
    /// merged". So the FIRST job with a canon section poisoned every later one
    /// in the same warm worker, and 16 component tests failed under
    /// `--worker` while passing 5/5 alone and 2025/2025 under `--cold`.
    ///
    /// It read as flakiness because which tests share a worker varies with
    /// scheduling — the failing SET moved run to run while its size did not.
    /// `worker.rs`'s own doc states the rule this restores: "any difference
    /// between warm and cold execution is a bug in it, not a property of warm
    /// mode."
    canon_defs: Vec<crate::canon_def::CanonDef>,
    canon_types: Vec<Option<crate::component::ValType>>,
    canon_functypes: Vec<Option<crate::canon_def::CanonFuncType>>,
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
    /// Call-tag tables (Call Tags proposal) — restored together for the same
    /// reason `tag_entities` and `chunk_tag_maps` are: truncating the entity
    /// list without the registry that indexes it leaves a name pointing at an
    /// id that no longer exists.
    call_tags_len: usize,
    call_tag_registry: HashMap<String, u32>,
    canonical_call_tags: HashMap<String, u32>,
    chunk_call_tag_maps_len: usize,
    // Registries a SCRIPT extends by running. Every one of these used to
    // survive a reset — the doc comment on `reset_to` even advertised leaving
    // the type registry and modules alone — so a program's class definitions
    // and registered modules were still there for the NEXT tenant of a reused
    // VM. That is a correctness bug between test runs and an isolation hole
    // anywhere a VM serves more than one program.
    //
    // `type_registry` restores by value; `modules` holds non-Clone records, so
    // the baseline KEY SET is captured instead and script-added entries are
    // dropped by name; `deferred_sources` only ever grows, so a length is
    // enough.
    type_registry: crate::typedef::TypeRegistry,
    module_keys: HashSet<String>,
    deferred_sources_len: usize,
    /// Whether type RECORDING was on at boot. The bank is indexed by chunk
    /// index and chunk indices are REUSED after truncation, so a surviving
    /// bank would apply one tenant's observations to another's code.
    type_recording: bool,
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
    chunk_import_owner: Vec<usize>,
    // Coupled with `tag_entities` (maps imported-tag name → index into it). Must
    // restore together: truncating tag_entities without this would leave a
    // dangling index a later lookup could read out of bounds.
    imported_tag_registry: HashMap<String, usize>,
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
    /// True if this label closes a try_table. Normal END must also pop
    /// the active exception handler for the protected region.
    pub is_try: bool,
    /// For a try label, the handler `group` its clauses were installed under.
    /// A `try_table` pushes ONE label but one handler per CLAUSE, so the label
    /// has to name the whole group: disposing of "one handler per exited try
    /// label" left every clause but the first armed. Zero when `is_try` is
    /// false. Identifying the region by group rather than by a handler-stack
    /// index keeps it valid across a JSPI fiber capture, which re-indexes the
    /// handler stack but never renumbers a group.
    pub try_group: u64,
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
    /// The receiver this VM hands a callee that declares one — §10.2.1
    /// `[[Call]](thisArgument, argumentsList)` argument 0.
    ///
    /// ⛔ THE ONE CHANNEL. A caller driving this VM from outside (a dynamically
    /// compiled `Function(...)` body invoked as a constructor) has to place the
    /// receiver where the prepend reads it, and this is that place.
    pub fn host_receiver_value(&self) -> Value {
        self.host_receiver.clone()
    }

    /// Set the host receiver; returns the previous value so a caller can
    /// restore it around one invocation.
    pub fn set_host_receiver(&mut self, value: Value) -> Value {
        std::mem::replace(&mut self.host_receiver, value)
    }

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
            host_receiver: Value::Undefined,
            host_originated_call: false,
            suppress_receiver_prepend: false,
            chunks: Vec::new(),
            frames: Vec::new(),
            stack: Vec::with_capacity(256),
            globals: vec![Value::Undefined],
            globals_assigned: vec![true],
            global_index: {
                let mut ix = HashMap::new();
                ix.insert("undefined".to_string(), 0u32);
                ix
            },
            open_upvalues: Vec::new(),
            host_fns: Vec::new(),
            host_fn_takes_receiver: Vec::new(),
            host_registry: HashMap::new(),
            modules: HashMap::new(),
            import_table: Vec::<ImportTarget>::new(),
            exception_handlers: Vec::new(),
            tag_entities: vec![TagEntity {
                debug_name: "vybe:exception".into(),
                arity: 1,
            }],
            chunk_tag_maps: Vec::new(),
            try_group_counter: 0,
            imported_tag_registry: HashMap::from([("vybe:exception".to_string(), 0usize)]),
            call_tags: Vec::new(),
            canonical_call_tags: HashMap::new(),
            func_switches: HashMap::new(),
            chunk_call_tag_maps: Vec::new(),
            call_tag_registry: HashMap::new(),
            call_tag_errors: Vec::new(),
            func_call_tags: HashMap::new(),
            event_loop: Rc::new(RefCell::new(EventLoop::new())),
            scheduler: None,
            deferred_sources: Vec::new(),
            type_registry: crate::typedef::TypeRegistry::new(),
            module_type_names: Vec::new(),
            module_type_ids: Vec::new(),
            chunk_type_base: Vec::new(),
            chunk_import_owner: Vec::new(),
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
            resume_reparked: false,
            dbg_last_import: None,
            exec_floors: Vec::new(),
            next_fiber_id: 1,
            label_stack: Vec::new(),
            block_tables: HashMap::new(),
            callback_invoker: None,
            invoker_vm_slot: Box::new(std::cell::Cell::new(std::ptr::null_mut())),
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
            // Seeded with the primitive component types so a CORE module can
            // name one by index. A real component supplies its own type
            // section and the indices come from there; a bare core module —
            // which is what a `.wat` test or a Vybe-compiled program is — has
            // no such section, and the `$t` immediate would otherwise have
            // nothing to point at. These four are a documented BOOTSTRAP
            // convention, not a spec requirement, and a component that
            // registers its own types simply appends past them.
            canon_types: vec![
                Some(crate::component::ValType::Bool), // 0
                Some(crate::component::ValType::I32),  // 1
                Some(crate::component::ValType::I64),  // 2
                Some(crate::component::ValType::F64),  // 3
            ],
            // Both start EMPTY. A component registers its own types and
            // canon definitions; an absent row must be an error naming the
            // missing declaration, not index 0 standing in for it.
            canon_functypes: Vec::new(),
            component_funcs: Vec::new(),
            canon_defs: Vec::new(),
            cm_instance: crate::cm_instance::ComponentInstance::new(),
            current_thread: None,
            canon_type_immediate: None,
            cm_tasks: Vec::new(),
            next_cm_task_id: 1,
            waitable_sets: crate::waitable::WaitableRegistry::new(),
        }
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
            global_index: self.global_index.clone(),
            memory: self.memory.with_buffer(|b| b.to_vec()),
            extra_memories: self.extra_memories.clone(),
            wasm_tables: self.wasm_tables.clone(),
            dropped_data: self.dropped_data.clone(),
            dropped_elems: self.dropped_elems.clone(),
            active_memory: self.active_memory,
            canon_defs: self.canon_defs.clone(),
            canon_types: self.canon_types.clone(),
            canon_functypes: self.canon_functypes.clone(),
            handle_table: self.handle_table.clone(),
            waitable_sets: self.waitable_sets.clone(),
            cm_tasks: self.cm_tasks.clone(),
            try_group_counter: self.try_group_counter,
            cur_fiber_id: self.cur_fiber_id,
            next_fiber_id: self.next_fiber_id,
            next_thread_id: self.next_thread_id,
            next_cm_task_id: self.next_cm_task_id,
            chunks_len: self.chunks.len(),
            chunk_tag_maps_len: self.chunk_tag_maps.len(),
            tag_entities_len: self.tag_entities.len(),
            call_tags_len: self.call_tags.len(),
            call_tag_registry: self.call_tag_registry.clone(),
            canonical_call_tags: self.canonical_call_tags.clone(),
            chunk_call_tag_maps_len: self.chunk_call_tag_maps.len(),
            type_registry: self.type_registry.clone(),
            module_keys: self.modules.keys().cloned().collect(),
            deferred_sources_len: self.deferred_sources.len(),
            type_recording: self.type_recorder.is_some(),
            func_table: self.func_table.clone(),
            case_aliases: self.case_aliases.clone(),
            import_table: self.import_table.clone(),
            data_segments: self.data_segments.clone(),
            elem_segments: self.elem_segments.clone(),
            module_type_names: self.module_type_names.clone(),
            module_type_ids: self.module_type_ids.clone(),
            chunk_type_base: self.chunk_type_base.clone(),
            chunk_import_owner: self.chunk_import_owner.clone(),
            imported_tag_registry: self.imported_tag_registry.clone(),
        }
    }

    /// Restore the VM to a [`snapshot`](VM::snapshot) baseline: free the whole
    /// post-snapshot script generation (objects + cycles, via `heap::restore`),
    /// drop script-added globals / restore reassigned ones, reset wasm memory &
    /// tables to boot bytes, and clear all transient execution state. Leaves the
    /// VM byte-indistinguishable from a freshly-booted one.
    ///
    /// Anything a SCRIPT creates by running is flushed — including the type
    /// registry, registered modules and deferred sources, which this used to
    /// leave in place. Leaving them meant a program's classes and modules were
    /// still resolvable by the next program in a reused VM: wrong between test
    /// runs, and an isolation hole wherever one VM serves more than one tenant.
    ///
    /// What survives is BOOT infrastructure only, none of which a script
    /// creates: host fns and their registry, the boot chunks (prelude), the
    /// scheduler the embedder installed, and the debugger/eval/reload hooks.
    pub fn reset_to(&mut self, snap: &VmSnapshot) {
        // 1. Heap: force-clear the script generation (breaks cycles so refcounts
        //    collapse to 0) and rewire baseline objects to their boot contents.
        //    Runs FIRST: collect_since clears contents regardless of live roots,
        //    so cycles break here; steps 2/5 then drop the roots.
        crate::heap::restore(&snap.heap);
        // 2. Globals: script-added keys vanish; reassigned baseline keys restored.
        self.globals = snap.globals.clone();
        self.global_index = snap.global_index.clone();
        self.globals_assigned = vec![true; self.globals.len()];
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
        // The canon section and its type spaces — see `VmSnapshot::canon_defs`.
        // A second job's section is REJECTED, not overwritten, so leaving these
        // behind made the first component program in a worker poison the rest.
        self.canon_defs = snap.canon_defs.clone();
        self.canon_types = snap.canon_types.clone();
        self.canon_functypes = snap.canon_functypes.clone();
        // 4b. Drop the prior run's appended CODE (and its embedded string/data
        //     constants — security: no earlier tenant's bytes survive) + the
        //     chunk-parallel structures that grow with it. Everything below the
        //     boot length is baseline (prelude) and stays. Other per-chunk caches
        //     keyed by index (block_tables, funcref_cache) are cleared in step 5.
        self.chunks.truncate(snap.chunks_len);
        self.chunk_tag_maps.truncate(snap.chunk_tag_maps_len);
        self.tag_entities.truncate(snap.tag_entities_len);
        self.call_tags.truncate(snap.call_tags_len);
        self.chunk_call_tag_maps.truncate(snap.chunk_call_tag_maps_len);
        self.call_tag_registry = snap.call_tag_registry.clone();
        self.canonical_call_tags = snap.canonical_call_tags.clone();
        self.call_tag_errors.clear();
        self.func_switches.clear();
        self.func_call_tags.clear();
        // 4c. Registries a SCRIPT extends by running. These are the ones that
        //     used to survive: a program's class definitions stayed in the type
        //     registry, its modules stayed registered, and the next program in
        //     a reused VM resolved against them. Chunk indices are REUSED after
        //     the truncation above, so a surviving type-observation bank would
        //     also attribute one tenant's observations to another's code.
        self.type_registry = snap.type_registry.clone();
        self.modules
            .retain(|name, _| snap.module_keys.contains(name));
        // A deferred source registered at boot is infrastructure and stays —
        // but its QUEUE belongs to the program that just ran. Those entries are
        // `Value`s closing over chunks truncated one line above, whose indices
        // the next program reuses.
        for src in &self.deferred_sources {
            src.clear_pending();
        }
        self.deferred_sources.truncate(snap.deferred_sources_len);
        // 4d. Host-side per-program state, which no amount of VM restoring can
        //     reach: plugins own it directly (timer/animation callbacks, DOM
        //     listeners and documents, cached constructors, connections). It is
        //     reset here rather than in an embedder so that EVERY caller of
        //     `reset_to` gets a total reset — a warm worker, a `--serve` request
        //     loop, or anything else that lets one VM serve two tenants.
        crate::framework::reset_all_registered();
        // 4e. Then drop the host resources the VM itself owns on the plugins'
        //     behalf ([`crate::resources`]) — open descriptors, HTTP bodies,
        //     key material, cached constructors, DOM listeners. LAST, and after
        //     the hook above on purpose: a plugin whose `reset` performs an
        //     ACTION (closing a socket, zeroing a key) has to iterate its table
        //     to find the handles, so clearing the storage first would leave
        //     those actions with nothing to close and leak in silence.
        //
        //     A plugin that only needs its state gone writes no `reset` at all
        //     — that is the point of the store. See `resources.rs`.
        crate::resources::clear_all();
        self.type_recorder = snap
            .type_recording
            .then(crate::type_recorder::TypeRecorder::new);
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
        self.chunk_import_owner = snap.chunk_import_owner.clone();
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
            None => Err("expression eval unavailable (no compiler hook attached)".to_string()),
        };
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
            None => Err("event simulation unavailable (no gui hook attached)".to_string()),
        };
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
            None => Err("hot reload unavailable (no compiler hook attached)".to_string()),
        };
        self.reload_hook = hook;
        let new_chunks = compiled?;
        self.apply_reload(new_chunks)
    }

    fn apply_reload(&mut self, mut new_chunks: Vec<Chunk>) -> Result<String, String> {
        // Same load-time gate as `run_linked_impl` — reloaded bodies execute
        // without the dispatch loop re-validating each instruction.
        Self::validate_chunk_code(&new_chunks).map_err(|e| e.to_string())?;
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
            let idxs: Vec<usize> = ac
                .caller_fiber
                .frames
                .iter()
                .map(|f| f.chunk_index)
                .collect();
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
        // The executing frame names the type base, exactly as `resolve_gc_rtt`
        // does for the equivalent instruction.
        let base = self
            .frames
            .last()
            .and_then(|frame| self.chunk_type_base.get(frame.chunk_index))
            .copied()
            .unwrap_or(0);
        self.eval_const_expr_with_type_base(expr, base)
    }

    /// [`Self::eval_const_expr`] with the GC type base stated outright, for the
    /// instantiation-time callers that have no frame to derive it from.
    pub(crate) fn eval_const_expr_with_type_base(
        &self,
        expr: &crate::chunk::ConstExpr,
        type_base: usize,
    ) -> Value {
        use crate::chunk::ConstExpr;
        match expr {
            ConstExpr::Value(v) => v.clone(),
            ConstExpr::GlobalGet(name) => self.global(name).cloned().unwrap_or(Value::Null),
            ConstExpr::Add(left, right) => {
                let l = self.eval_const_expr_with_type_base(left, type_base);
                let r = self.eval_const_expr_with_type_base(right, type_base);
                match (&l, &r) {
                    (Value::I32(a), Value::I32(b)) => Value::I32(a.wrapping_add(*b)),
                    (Value::I64(a), Value::I64(b)) => Value::I64(a.wrapping_add(*b)),
                    (Value::F64(a), Value::F64(b)) => Value::F64(a + b),
                    _ => Value::F64(l.as_f64() + r.as_f64()),
                }
            }
            ConstExpr::Mul(left, right) => {
                let l = self.eval_const_expr_with_type_base(left, type_base);
                let r = self.eval_const_expr_with_type_base(right, type_base);
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
                    Value::Object(crate::heap::alloc(obj))
                } else {
                    Value::Null
                }
            }
            // An `i31ref` is UNBOXED: `Op::I31_NEW` masks to 31 bits and pushes
            // a plain `I32`. The constant form must produce the SAME value, or
            // `ref.eq` — which for i31 compares the integer, there being no
            // pointer — would separate a segment's `ref.i31 7` from an
            // instruction's.
            ConstExpr::RefI31(inner) => {
                let v = self.eval_const_expr_with_type_base(inner, type_base).as_i32();
                Value::I32(v & 0x7FFF_FFFF)
            }
            // Mirrors `Op::ARRAY_NEW_DEFAULT`: a defaulted array carrying the
            // resolved rtt, so a GC array built by an element expression traps
            // on out-of-bounds exactly like one built by the instruction.
            ConstExpr::ArrayNewDefault { typeidx, len } => {
                let n = self
                    .eval_const_expr_with_type_base(len, type_base)
                    .as_i32()
                    .max(0) as usize;
                self.const_array(*typeidx, type_base, vec![Value::Null; n])
            }
            ConstExpr::ArrayNew { typeidx, value, len } => {
                let v = self.eval_const_expr_with_type_base(value, type_base);
                let n = self
                    .eval_const_expr_with_type_base(len, type_base)
                    .as_i32()
                    .max(0) as usize;
                self.const_array(*typeidx, type_base, vec![v; n])
            }
        }
    }

    /// A GC array built by a constant expression, stamped with the rtt its
    /// type immediate names. Mirrors `Op::ARRAY_NEW`/`ARRAY_NEW_DEFAULT` so an
    /// array from an element segment traps on out-of-bounds exactly like one
    /// built by the instruction. `typeidx` is 1-based (0 = dynamic).
    fn const_array(&self, typeidx: u16, type_base: usize, elems: Vec<Value>) -> Value {
        let mut obj = Object::new_array(elems);
        obj.type_id = if typeidx == 0 {
            0
        } else {
            self.module_type_ids
                .get(type_base + typeidx as usize - 1)
                .copied()
                .unwrap_or(0)
        };
        Value::Object(crate::heap::alloc(obj))
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
            // A declared minimum we cannot allocate must FAIL instantiation,
            // never be attempted: `resize` on a 4-billion-element declaration
            // faults in every page and the host dies before the module runs.
            // The spec allows instantiation to fail on resource exhaustion.
            if size > TABLE_ALLOC_LIMIT {
                return Err(crate::VMError::new(
                    "table declaration exceeds the host allocation limit",
                ));
            }
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
                        "trap: out of bounds memory access: addr={} size={} limit={}",
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
                    "trap: out of bounds memory access: addr={} size={} limit={}",
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
                        "trap: out of bounds memory access: addr={} size={} limit={}",
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
                    "trap: out of bounds memory access: addr={} size={} limit={}",
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
        // A `try_table` pushes ONE structural label but one handler per CLAUSE,
        // all sharing a `group`. Branching out of it leaves the region exactly
        // as falling off its `end` does, so it must dispose of the whole group.
        // Disposing of one handler per exited try label left every clause but
        // the first armed, and a stale handler is still matched — a later throw
        // that must escape gets caught by a region the program already left.
        //
        // Every try_table still open inside the exited range has its own label
        // in this slice, so naming each exited label's group removes exactly
        // the regions being left and nothing enclosing them.
        let exited_groups: Vec<u64> = self.label_stack[new_len..]
            .iter()
            .filter(|label| label.is_try)
            .map(|label| label.try_group)
            .collect();
        if !exited_groups.is_empty() {
            self.exception_handlers
                .retain(|h| !exited_groups.contains(&h.group));
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
    /// Register a host function from a DECLARATION RECORD.
    ///
    /// One parameter, not a growing list of them. Every dimension a host
    /// function might declare — its signature today, its resource binding
    /// tomorrow, effects or capabilities after that — is a FIELD with a
    /// default, so adding one touches no existing caller. That property is the
    /// whole point: the 1,898 `register_host_fn` sites in this tree must never
    /// have to move again.
    ///
    /// [`Self::register_host_fn`] is this function with everything defaulted —
    /// not a legacy path, not a second mechanism, just the record with `sig:
    /// None`. An undeclared function behaves exactly as it always has:
    /// `mount_host_exports` mounts it with `arity: None`, which reports
    /// honestly that the registry never knew.
    pub fn register_host(&mut self, decl: HostFnDecl) {
        let HostFnDecl {
            module,
            name,
            call,
            sig,
            resource,
            takes_receiver,
        } = decl;
        if let Some(sig) = sig {
            // PROCESS-GLOBAL, not a VM field: the consumer is the COMPILER,
            // which never holds a VM. The namespace tree is global for the same
            // reason, and `mount_host_exports` already reads the capability
            // context that way.
            crate::declare_host_signature(&module, &name, sig, resource);
        }
        self.register_host_fn(&module, &name, call);
        if let Some(idx) = self.host_registry.get(&(module, name)).copied() {
            self.host_fn_takes_receiver[idx] = takes_receiver;
        }
    }

    /// Register a host function whose TYPE HAS NO RECEIVER — a free function
    /// (`encodeURIComponent(s)`, `Number(v)`), not a method spelled as a call.
    ///
    /// The receiver slot is then never filled for it, on either call path, so
    /// it has ONE shape: called directly and handed to `map` look identical to
    /// the callee. See [`HostFnDecl::takes_receiver`].
    pub fn register_free_fn(
        &mut self,
        module: &str,
        name: &str,
        f: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
    ) {
        self.register_host_fn(module, name, f);
        if let Some(idx) = self
            .host_registry
            .get(&(module.to_string(), name.to_string()))
            .copied()
        {
            self.host_fn_takes_receiver[idx] = false;
        }
    }

    /// Does the host function at `idx` declare a receiver parameter?
    ///
    /// Unknown indices answer `true` — the historical shape — so a lookup that
    /// races registration cannot silently drop a method's receiver.
    pub(crate) fn host_fn_declares_receiver(&self, idx: usize) -> bool {
        self.host_fn_takes_receiver.get(idx).copied().unwrap_or(true)
    }

    // ── Call tags (proposals/call-tags) ──────────────────────────────────
    //
    // Tags are module entities like EH tags, memories or globals: declared in
    // the tag section, referenced by index, importable and exportable. These
    // are the declaration half; `call_with_tag` in `dispatch.rs` is the use.

    /// `call_tag.canon $functype` — the canonical tag for a signature.
    ///
    /// INTERNED: the Overview defines `call_indirect $table $functype` as
    /// `call_with_tag (call_tag.canon $functype)`, so every module deriving the
    /// canonical tag for the same signature must get the SAME tag, or an
    /// indirect call would stop matching a func that handles it.
    pub fn call_tag_canon(&mut self, params: u8, results: u8, signature: &str) -> u32 {
        // ⛔ THE KEY IS THE FUNCTYPE, NOT THE SHAPE. A producer that supplies no
        // spelling (every non-wast frontend) falls back to the shape, so its
        // behaviour is unchanged; wast supplies the declared types and gets one
        // canonical tag per functype, which is what the Overview defines.
        let key = if signature.is_empty() {
            format!("#shape[{params}->{results}]")
        } else {
            signature.to_string()
        };
        // Same staleness rule as the name registry: an interned canonical id
        // is only valid while its entity is still present.
        if let Some(existing) = self.canonical_call_tags.get(&key)
            && (*existing as usize) < self.call_tags.len()
        {
            return *existing;
        }
        let idx = self.call_tags.len() as u32;
        self.call_tags.push(CallTagDef {
            debug_name: format!("canon[{key}]"),
            params,
            results,
            signature: signature.to_string(),
            // "For canonical call tags, the answer is simply that the program
            // traps" — a canonical tag has no fall-back, by definition.
            fallback: None,
            canonical: true,
        });
        self.canonical_call_tags.insert(key, idx);
        idx
    }

    /// `call_tag.new $functype $func?` — a FRESH tag over a signature that may
    /// already have one. Never interned: minting a distinct identity is the
    /// entire purpose.
    ///
    /// `fallback` is the Overview's optional handler, whose signature must be
    /// the tag's `[ti*]` plus a trailing `funcref` — it receives the funcref
    /// that failed to handle the tag, so it can adapt or reject.
    pub fn call_tag_new(
        &mut self,
        debug_name: &str,
        params: u8,
        results: u8,
        signature: &str,
        fallback: Option<Value>,
    ) -> u32 {
        let idx = self.call_tags.len() as u32;
        self.call_tags.push(CallTagDef {
            debug_name: debug_name.to_string(),
            signature: signature.to_string(),
            params,
            results,
            fallback,
            canonical: false,
        });
        idx
    }

    /// `(func … (call_tag $call_tag*))` — declare which tags a func's `funcref`
    /// handles. Declaring ANY replaces the default (its own canonical tag), which
    /// is what makes the Overview's security property work: "if one specifies no
    /// canonical call tags and only non-exported call tags, then one can be
    /// guaranteed that the function is only indirectly called by this module".
    pub fn declare_func_call_tags(&mut self, chunk_index: usize, tags: Vec<u32>) {
        self.func_call_tags.insert(chunk_index, tags);
    }

    /// Does the func defined by `chunk_index` handle `tag`?
    ///
    /// With no declaration the func handles exactly the canonical tag of its own
    /// signature — the Overview's default — so an undeclared func stays callable
    /// through `call_indirect` and nothing existing changes behaviour.
    pub(crate) fn func_handles_call_tag(&mut self, chunk_index: usize, tag: u32) -> bool {
        if let Some(declared) = self.func_call_tags.get(&chunk_index) {
            return declared.contains(&tag);
        }
        let Some(chunk) = self.chunks.get(chunk_index) else {
            return false;
        };
        // ⛔ `param_count`, NOT `arity`. `param_count` is documented as the
        // function's declared WASM parameter count (user params only, no
        // receiver) and is the half of the type SHAPE that `call_indirect`'s
        // runtime check compares against. `arity` can include a receiver, so
        // deriving the canonical tag from it would compute a different tag
        // than the call site does and reject a func that does handle it.
        let (params, results) = (chunk.param_count, chunk.result_arity);
        // ⛔ THE FUNCTYPE, NOT THE SHAPE. An undeclared func handles the
        // canonical tag OF ITS OWN TYPE, and `Chunk::func_sig` already carries
        // that — the declared VALUE TYPES, in order. Falls back to the shape
        // for any chunk that is not a wast function, which is all those can
        // answer and exactly what they did before.
        let sig = chunk
            .func_sig
            .as_ref()
            .map(|(p, r)| format!("{p}->{r}"))
            .unwrap_or_default();
        // ⚠ BOTH KEYS, and the reason is a real gap rather than caution.
        // `call_indirect` derives its canonical tag from the CALL SITE's
        // `$functype`, but at runtime that instruction carries only
        // `(argc, expected_results)` — counts, no types — so it can only ask
        // for the SHAPE tag. A wast func, which does know its own functype,
        // must therefore answer to both: its functype tag (so
        // `call_tag.canon [i32]->[i32]` and `[f64]->[f64]` stay distinct
        // entities) and its shape tag (so `call_indirect` still reaches it).
        //
        // ⛔ This is not the proposal's end state. Widening `call_indirect`'s
        // immediate to carry the functype is what would let the call site ask
        // for the typed tag — and that is an operand-width change, the exact
        // class of change that desynchronised every bytecode walker past a call
        // tag earlier today, so it wants its own step and its own gate.
        if !sig.is_empty() && self.call_tag_canon(params, results, &sig) == tag {
            return true;
        }
        self.call_tag_canon(params, results, "") == tag
    }

    pub fn register_host_fn(
        &mut self,
        module: &str,
        name: &str,
        f: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
    ) {
        let idx = self.host_fns.len();
        self.host_fns.push(Arc::from(f));
        self.host_fn_takes_receiver.push(true);
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
        // Clone the Rc so host functions can enqueue ready work without
        // holding a mutable borrow of the VM.
        let el = self.event_loop.clone();
        // Raw pointer to last_exception — safe: valid for host call duration.
        let exc_ptr = &mut self.last_exception as *mut Option<Value>;
        let exit_ptr = &mut self.pending_exit as *mut bool;
        let exit_code_ptr = &mut self.pending_exit_code as *mut i32;
        let host_receiver_ptr = &mut self.host_receiver as *mut Value;
        // The module chunk's ABI, read BEFORE the raw borrows below — it
        // decides which receiver channel the two accessors use.
        // ⛔ THE CALLING FRAME'S CHUNK, NOT `chunks[0]`. `chunks.first()` is ONE
        // answer for the whole BUNDLE: the per-unit stamp writes every chunk of
        // its own unit, so in a multi-language bundle chunk 0 belongs to
        // whichever unit happens to be first. A dart, csharp or powershell
        // frame then reported `ReceiverAbi::Parameter` because a js unit was
        // bundled alongside, and every host-callee path that asks
        // `receiver_is_parameter()` took the js branch — skipping the ambient
        // prepend and dropping the receiver.
        //
        // Measured: dart `for (var c in "abc".split(""))` regressed 99 tests,
        // csharp `a.SetEquals(b)` answered False, powershell `$s.ToUpper()`
        // returned empty — three languages, one wrong question. Forcing the
        // INVARIANT off fixed none of them, which is what finally pointed here:
        // the invariant reads the calling frame (correct) while this read the
        // bundle (wrong), so they disagreed about the same call.
        //
        // Same rule as `VM::module_receiver_abi`. With no frame the caller is
        // host code, which keeps the previous default.
        let module_abi = self
            .frames
            .last()
            .and_then(|f| self.chunks.get(f.chunk_index))
            .map_or(crate::chunk::ReceiverAbi::Ambient, |c| {
                c.module_receiver_abi
            });
        let globals_ptr = &mut self.globals as *mut Vec<Value>;
        let global_index_ptr = &mut self.global_index as *mut HashMap<String, u32>;
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
            host_receiver_slot: host_receiver_ptr,
            module_receiver_abi: module_abi,
            globals_slot: globals_ptr,
            global_index_slot: global_index_ptr,
            stack_slot: &self.stack as *const Vec<Value>,
            handle_table_slot: &mut self.handle_table as *mut crate::handle_table::HandleTable,
            shared_memory_slot: &self.memory as *const crate::shared_memory::SharedMemory,
            type_registry_slot: &self.type_registry as *const crate::typedef::TypeRegistry,
            call_tag_registry_slot: &self.call_tag_registry as *const HashMap<String, u32>,
            func_call_tags_slot: &self.func_call_tags as *const HashMap<usize, Vec<u32>>,
            chunks_slot: &self.chunks as *const Vec<crate::chunk::Chunk>,
        }
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
        // Republish where the VM is NOW, before handing the invoker out. The
        // closure below is built once and reused, so it cannot carry the
        // address — see `invoker_vm_slot` for what capturing one cost.
        let here = self as *mut VM;
        self.invoker_vm_slot.set(here);
        // This is stored as a field to avoid repeated allocation
        if self.callback_invoker.is_none() {
            let slot: *const std::cell::Cell<*mut VM> = &*self.invoker_vm_slot;
            self.callback_invoker = Some(Box::new(move |func_ref: &Value, args: &[Value]| {
                // SAFETY: `slot` is the VM's own boxed cell, alive as long as
                // the VM that owns this closure; the line above wrote the
                // current `&mut self` into it for this very host call, and the
                // VM cannot move while that borrow is outstanding.
                let vm = unsafe { &mut *(*slot).get() };
                vm.invoke_callback(func_ref, args)
            }));
        }
        self.callback_invoker.as_mut().unwrap().as_mut()
    }

    /// Does this callable bind argument 0 to its receiver?
    ///
    /// Resolves the value to its chunk and reads [`Chunk::takes_receiver`]. A
    /// host function is never a bytecode callee and answers `false`; so does a
    /// value that is not callable at all, which the call itself then rejects.
    /// Does the code making this call keep its class members on the prototype
    /// (so a member read must not fall through to the `TypeRegistry` vtable)?
    ///
    /// Asks the CALLING frame's chunk, not `chunks[0]`, for the same reason
    /// [`Self::module_receiver_abi`] does: in a multi-language bundle
    /// `chunks[0]` belongs to whichever unit happens to be first. With no frame
    /// the caller is host code, which keeps the existing behaviour.
    /// Does the CALLING frame's language make declared instance fields own
    /// properties? See [`crate::chunk::Chunk::module_instance_fields_are_own_properties`].
    pub(crate) fn calling_module_fields_are_own_properties(&self) -> bool {
        self.frames
            .last()
            .and_then(|f| self.chunks.get(f.chunk_index))
            .is_some_and(|c| c.module_instance_fields_are_own_properties)
    }

    pub(crate) fn calling_module_members_on_prototype(&self) -> bool {
        self.frames
            .last()
            .and_then(|f| self.chunks.get(f.chunk_index))
            .is_some_and(|c| c.module_members_on_prototype)
    }

    /// This module's receiver ABI, read off the module chunk.
    /// The receiver ABI of the code making the CALL — the currently executing
    /// frame's chunk, not `chunks[0]`.
    ///
    /// ⛔ ASK THE CALL SITE, NOT THE BUNDLE. Whether a receiver was pushed is
    /// decided by the unit that emitted the call, and in a multi-language
    /// bundle `chunks[0]` belongs to whichever unit happens to be first. With
    /// no frame the caller is host code, which pushes no receiver of its own —
    /// and that path is flagged `from_host` anyway.
    pub(crate) fn module_receiver_abi(&self) -> crate::chunk::ReceiverAbi {
        self.frames
            .last()
            .and_then(|f| self.chunks.get(f.chunk_index))
            .map_or(crate::chunk::ReceiverAbi::Ambient, |c| {
                c.module_receiver_abi
            })
    }

    /// Is this value a host function? Asked separately from
    /// [`Self::callee_takes_receiver`], which can only answer for callees that
    /// resolve to a chunk.
    fn is_host_function(func_ref: &Value) -> bool {
        let Value::Object(obj) = func_ref else {
            return false;
        };
        let Ok(o) = obj.lock() else { return false };
        matches!(o.kind, crate::value::ObjectKind::HostFunction(_))
    }

    fn callee_takes_receiver(&self, func_ref: &Value) -> bool {
        let Value::Object(obj) = func_ref else {
            return false;
        };
        let Ok(o) = obj.lock() else { return false };
        let crate::value::ObjectKind::Function(f) = &o.kind else {
            return false;
        };
        self.chunks
            .get(f.chunk_index)
            .is_some_and(|c| c.takes_receiver)
    }

    /// Invoke a VM function reference from host code.
    /// This is the WASM-compliant callback mechanism: host functions
    /// can call exported/internal VM functions during execution.
    ///
    /// Usage from a host function:
    ///   let result = vm.invoke_callback(&predicate, &[element]);
    /// Invoke with an EXPLICIT receiver and suppress the automatic prepend.
    ///
    /// ⛔ THE MISSING PRIMITIVE. `invoke_callback` prepends a receiver it
    /// infers, which is right when host code has none to give — a `map`
    /// callback, a microtask. It is wrong wherever the host DOES know the
    /// receiver: an accessor knows the object it is reading, and passing it
    /// through the argument list on top of the inferred one handed the callee
    /// TWO, so `this` became the inferred value and the first declared
    /// parameter got the real receiver.
    ///
    /// Host code that knows the receiver calls this; host code that does not
    /// calls `invoke_callback`. Under the ambient binding the receiver is bound
    /// to the module global instead, exactly as before.
    pub fn invoke_with_receiver(
        &mut self,
        func_ref: &Value,
        receiver: Value,
        args: &[Value],
    ) -> Value {
        // Under `Parameter` the receiver IS argument 0. Hand it over directly
        // and let `invoke_callback` add nothing: it only prepends for a callee
        // that declares one, and this callee has just been given it.
        let saved = std::mem::replace(&mut self.suppress_receiver_prepend, true);
        let mut all = Vec::with_capacity(args.len() + 1);
        all.push(receiver);
        all.extend_from_slice(args);
        let out = self.invoke_callback(func_ref, &all);
        self.suppress_receiver_prepend = saved;
        out
    }

    pub fn invoke_callback(&mut self, func_ref: &Value, args: &[Value]) -> Value {
        let saved_frame_depth = self.frames.len();
        // Save the stack height so we can restore it after the callback returns,
        // giving the callback an isolated value stack (WASM call-frame semantics).
        let saved_stack_len = self.stack.len();

        // Push function ref + args onto stack
        self.stack.push(func_ref.clone());
        // ECMA-262 §10.2.1 `[[Call]](thisArgument, argumentsList)`: a callee
        // that binds argument 0 to its RECEIVER must be handed one, and host
        // code has none to give — `Array.prototype.map` calls its callback with
        // the element list and nothing else. §10.2.1.1 OrdinaryCallBindThis
        // with a non-object thisArgument is `undefined`, so that is what goes
        // in, explicitly, rather than the argument list shifting into the
        // receiver slot and every parameter arriving one place late.
        //
        // ⛔ Asked of the CALLEE, never assumed from the call site: a region
        // that does not pass receivers leaves `takes_receiver` false and this
        // is inert.
        // ⛔ A HOST callee gets NO receiver here, and that is not an
        // inconsistency with the method-call path — it is the same rule read
        // correctly. THE RECEIVER IS A PROPERTY OF THE CALL, NOT OF THE
        // MODULE. `d.resolve(42)` from bytecode is a JS call and carries one;
        // the microtask drain invoking that same settler with the exact
        // arguments it constructed is host PLUMBING and carries none.
        //
        // Measured, when this prepended one for host callees on the module's
        // ABI alone: `promise_then_catch_finally_chaining` 5 → 21 regressed,
        // `promise_rejection_propagation` 12 → 130, `promise_finally_errors`
        // 6 → 141. Every internal settler had its arguments shifted.
        // `call_value_inner` tells the host function which of the two it was,
        // via `HostContext::call_receiver_argc`.
        let receiver_argc = usize::from(
            !self.suppress_receiver_prepend && self.callee_takes_receiver(func_ref),
        );
        if receiver_argc == 1 {
            // WHICH receiver — not a hard-coded `undefined`.
            //
            // `[1,2].map(fn, thisArg)`, `f.call(o)` and `f.apply(o, …)` reach a
            // bytecode callee through `ecma::function::invoke_with_explicit_this`,
            // which brackets the invocation with `set_js_this(this_arg)`
            // because under the ambient protocol that global WAS the channel.
            // Passing `undefined` here regardless drops the caller's explicit
            // receiver: measured, `map(fn)` worked and `map(fn, thisArg)`,
            // `.call` and `.apply` all threw.
            //
            // ⛔ NOT the ambient protocol returning. Under
            // `UniversalParameter` no EMITTED instruction reads or writes this
            // slot — the compiler has none left. It is a host-side, always
            // -restored parameter to this one call, read at the single boundary
            // where the host holds a receiver and the callee declares a
            // parameter for it. An unset slot reads as null, which becomes the
            // `undefined` thisArgument §10.2.1.1 gives an ordinary call.
            // ⛔ DO NOT COERCE NULL TO UNDEFINED. §10.2.1.1 passes the
            // thisArgument through UNCHANGED in strict mode, so `f.call(null)`
            // must see `this === null`. `host_receiver` already defaults to
            // `Undefined`, so a `Null` here can only be one someone set
            // deliberately — collapsing it loses the distinction the spec draws.
            let this_arg = self.host_receiver.clone();
            self.stack.push(this_arg);
        }
        for arg in args {
            self.stack.push(arg.clone());
        }

        // Tell the dispatch this call is host plumbing, so a host callee is not
        // charged a receiver slot it was never given.
        self.host_originated_call = true;
        // Call the function (pushes a new frame for compiled fns; inline for host fns)
        if self.call_value(args.len() + receiver_argc).is_err() {
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

    /// Validate every instruction of every incoming chunk BEFORE any of it
    /// can execute — the WASM spec's own architecture (validation is a phase
    /// preceding instantiation), applied to every bytecode source alike: our
    /// compilers, wast, reload, nested eval, and foreign `.wasm` binaries.
    ///
    /// This is what lets the dispatch loop construct `Op` WITHOUT the
    /// per-instruction `wasm_name_opt` probe (measured: a top-3 sample on a
    /// pure-arithmetic loop). The security posture is STRICTLY stronger than
    /// the probe it replaces: a malformed module used to execute up to its
    /// first bad opcode — side effects already done — where it is now
    /// rejected here having run nothing. And an op that somehow escaped this
    /// pass still lands in the dispatch match's final `Unhandled opcode`
    /// arm: an error, never undefined behaviour.
    fn validate_chunk_code(chunks: &[Chunk]) -> Result<(), VMError> {
        for chunk in chunks {
            let code = &chunk.code;
            let mut ip = 0usize;
            while ip + 3 < code.len() {
                let group = ((code[ip] as u16) << 8) | code[ip + 1] as u16;
                let sub = ((code[ip + 2] as u16) << 8) | code[ip + 3] as u16;
                let Some(op) = Op::decode(group, sub) else {
                    return Err(VMError::new(format!(
                        "invalid opcode 0x{:04X} 0x{:04X} at offset {} in chunk '{}' — module rejected at load",
                        group, sub, ip, chunk.name
                    )));
                };
                if op == Op::REF_FUNC {
                    // Variable shape `operand_format` cannot size:
                    // 4 opcode + 2 func_idx + 1 uv_count + uv_count × 3.
                    ip += 4 + 2 + 1;
                    if ip - 1 < code.len() {
                        let uv_count = (code[ip - 1] & 0x7f) as usize;
                        ip += uv_count * 3;
                    }
                    continue;
                }
                ip += 4;
                let fmt = op.operand_format();
                ip += fmt.size_in(code, ip);
            }
        }
        Ok(())
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
        Self::validate_chunk_code(&chunks)?;
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
        // Rewrite the incoming set's global operands into THIS VM's index
        // space before the chunks join it — each set was compiled against
        // its own table.
        self.merge_global_table(&mut adjusted);
        self.merge_canon_section(&adjusted)?;
        self.merge_canon_types(&adjusted)?;
        self.chunks.extend(adjusted);
        // ⛔ EVERY CHUNK OF THIS UNIT POINTS AT THIS UNIT'S TABLE. `link.rs`
        // remapped their `CALL` operands into `chunks[script_idx].imports`, so
        // that is the ONLY table those indices mean anything against — and the
        // VM-wide one is about to be replaced by the next unit's.
        self.chunk_import_owner.resize(self.chunks.len(), script_idx);
        for owner in self.chunk_import_owner[script_idx..].iter_mut() {
            *owner = script_idx;
        }
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
                ImportTarget::Canon(b, t) => {
                    self.import_table.push(ImportTarget::Canon(b, t));
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
            upvalues: Vec::new(),
        });
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
        // Rewrite the incoming set's global operands into THIS VM's index
        // space before the chunks join it — each set was compiled against
        // its own table.
        self.merge_global_table(&mut adjusted);
        self.merge_canon_section(&adjusted)?;
        self.merge_canon_types(&adjusted)?;
        self.chunks.extend(adjusted);
        // ⛔ EVERY CHUNK OF THIS UNIT POINTS AT THIS UNIT'S TABLE. `link.rs`
        // remapped their `CALL` operands into `chunks[script_idx].imports`, so
        // that is the ONLY table those indices mean anything against — and the
        // VM-wide one is about to be replaced by the next unit's.
        self.chunk_import_owner.resize(self.chunks.len(), script_idx);
        for owner in self.chunk_import_owner[script_idx..].iter_mut() {
            *owner = script_idx;
        }
        // Resolve call-tag declarations EAGERLY, as soon as the chunks that
        // carry them are installed.
        //
        // It used to happen lazily, on the first `call_with_tag` — but the
        // tables are also read by HOST functions (property dispatch asks which
        // receiver convention an accessor declares), and a program that never
        // executes a tagged call would leave them empty. The host would then
        // see every accessor as undeclared and treat a receiver-first one as
        // ambient.
        self.resolve_chunk_call_tags();
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
                                        upvalues: Vec::new(),
                                    };
                                    let mut obj = Object::new();
                                    obj.kind = ObjectKind::Function(func);
                                    Value::Object(crate::heap::alloc(obj))
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
                // Apply only to a slot nobody has written — a host-installed
                // native or a real program value both count as written.
                let unwritten = match self.global_index.get(&gi.name) {
                    Some(&i) => !self.globals_assigned.get(i as usize).copied().unwrap_or(true),
                    None => true,
                };
                if unwritten {
                    let val = self.eval_const_expr(&gi.init);
                    self.set_global(&gi.name, val);
                }
            }
        }

        // ⚠ ELEMENT SEGMENTS INSTANTIATE **AFTER** GLOBALS AND THE TYPE TABLE.
        //
        // WASM 3.0 §4.5.4 fixes the order of module instantiation: globals are
        // evaluated, THEN element segments, THEN data segments, and only then
        // are tables filled and `start` run. This block used to sit ~90 lines
        // earlier — before imports, before the GC type table, and before the
        // global initializers — which was survivable only for as long as a
        // segment could hold nothing but `ref.func`. It cannot: an element
        // expression is a CONSTANT EXPRESSION, and `(item (array.new_default
        // $t …))` needs the type table while `(item (global.get $g))` needs
        // the globals. Running first is what made both of those unreachable.
        //
        // The rtt base is taken from the SCRIPT CHUNK rather than the executing
        // frame, because at instantiation there is no frame yet — `resolve_gc_rtt`
        // would fall back to base 0 and name another module's type.
        let elem_items = self.chunks[script_idx].passive_elem_items.clone();
        let type_base = self
            .chunk_type_base
            .get(script_idx)
            .copied()
            .unwrap_or(0);
        for (seg_idx, items) in elem_items.iter().enumerate() {
            // ⛔ AN EMPTY SEGMENT IS STILL A SEGMENT. Skipping it left its index
            // unoccupied, so `table.init $e` on a dropped or empty segment
            // reported "missing element segment" instead of the out-of-bounds
            // SOURCE trap the spec asks for (`elem.wast`, implicitly-dropped
            // active segments).
            if items.is_empty() {
                self.set_elem_segment(seg_idx, Vec::new());
                continue;
            }
            let vals: Vec<crate::value::Value> = items
                .iter()
                .map(|e| self.eval_const_expr_with_type_base(e, type_base))
                .collect();
            self.set_elem_segment(seg_idx, vals);
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

        // Same two-sided rule as `invoke_callback`, which this had drifted from:
        // a bytecode callee that binds argument 0 to its receiver must be
        // handed one (§10.2.1), and a HOST callee must be told this dispatch
        // was host plumbing so it does not charge a receiver slot to the
        // arguments the caller built. The microtask drain reaches every promise
        // reaction through here, so a mismatch loses the reaction's value —
        // measured as `.then(v => …)` seeing `undefined`.
        let receiver_argc = usize::from(self.callee_takes_receiver(callee));
        self.push(callee.clone())?;
        if receiver_argc == 1 {
            // ⛔ DO NOT COERCE NULL TO UNDEFINED. §10.2.1.1 passes the
            // thisArgument through UNCHANGED in strict mode, so `f.call(null)`
            // must see `this === null`. `host_receiver` already defaults to
            // `Undefined`, so a `Null` here can only be one someone set
            // deliberately — collapsing it loses the distinction the spec draws.
            let this_arg = self.host_receiver.clone();
            self.push(this_arg)?;
        }
        for arg in args {
            self.push(arg.clone())?;
        }

        self.host_originated_call = true;
        self.call_value(args.len() + receiver_argc)?;

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
            // ⛔ `trap: ` IS LOAD-BEARING, NOT DECORATION. `VMError::is_trap`
            // classifies solely on that prefix, and only a message carrying it
            // is offered to a handler — so a bare "Stack overflow" escaped
            // UNCAUGHT and killed the run. The spec says exhausting the call
            // stack traps, and `(assert_exhaustion (invoke "runaway")
            // "call stack exhausted")` asserts exactly that it is catchable.
            // Even `assert_trap` quoting our own old wording could not catch
            // it, which is how the prefix rather than the text was identified
            // as the defect. Same lesson `Op::REF_CAST` records one file over.
            //
            // The wording is the reference interpreter's, so the fixture's
            // expected message matches as well.
            return Err(VMError::new("trap: call stack exhausted"));
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

    /// Read a fixed-width big-endian `u32` immediate, advancing `ip`.
    pub(crate) fn read_u32(&mut self) -> u32 {
        let a = self.read_byte() as u32;
        let b = self.read_byte() as u32;
        let c = self.read_byte() as u32;
        let d = self.read_byte() as u32;
        (a << 24) | (b << 16) | (c << 8) | d
    }

    /// Read an unsigned LEB128 `u32` immediate, advancing `ip`.
    ///
    /// Every index in WASM is a `u32` (`syntax idx = u32`), so this is the
    /// reader for an index operand — `read_u16` is for the shapes that carry a
    /// genuinely 16-bit field, never for an index.
    pub(crate) fn read_leb_u32(&mut self) -> u32 {
        let mut result: u32 = 0;
        let mut shift = 0u32;
        loop {
            let byte = self.read_byte();
            result |= ((byte & 0x7f) as u32) << shift;
            if byte & 0x80 == 0 {
                break;
            }
            shift += 7;
        }
        result
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

    /// Resolve a global BY NAME — instantiation, host binding, the debugger.
    /// WASM resolves imported globals by name at instantiate time too; what it
    /// never does is consult a name during execution.
    pub fn global(&self, name: &str) -> Option<&Value> {
        self.global_index
            .get(name)
            .and_then(|&i| self.globals.get(i as usize))
    }

    /// Install/overwrite a global by name, allocating an index if it is new.
    /// Setup-time only.
    pub fn set_global(&mut self, name: &str, value: Value) {
        match self.global_index.get(name) {
            Some(&i) => {
                let i = i as usize;
                if i >= self.globals.len() {
                    self.globals.resize(i + 1, Value::Null);
                    self.globals_assigned.resize(i + 1, false);
                }
                self.globals[i] = value;
                self.globals_assigned[i] = true;
            }
            None => {
                let i = self.globals.len() as u32;
                self.global_index.insert(name.to_string(), i);
                self.globals.push(value);
                self.globals_assigned.push(true);
            }
        }
    }

    /// `set_global` taking an owned name — host installation sites read
    /// `insert(name, value)` and this keeps them doing so.
    pub fn set_global_owned(&mut self, name: impl Into<String>, value: Value) {
        let name = name.into();
        self.set_global(&name, value);
    }

    /// Remove a global by name. The slot stays allocated (an index is stable
    /// for the module's lifetime, as in WASM); the binding is dropped.
    pub fn remove_global(&mut self, name: &str) -> Option<Value> {
        let idx = self.global_index.remove(name)? as usize;
        let old = self.globals.get(idx).cloned();
        if let Some(slot) = self.globals.get_mut(idx) {
            *slot = Value::Undefined;
        }
        old
    }

    pub fn has_global(&self, name: &str) -> bool {
        self.global_index.contains_key(name)
    }

    /// Every global as `(name, value)` — for the debugger and snapshots, which
    /// want the names. Never on the execution path.
    pub fn globals_by_name(&self) -> Vec<(String, Value)> {
        let mut out: Vec<(String, Value)> = self
            .global_index
            .iter()
            .filter_map(|(n, &i)| self.globals.get(i as usize).map(|v| (n.clone(), v.clone())))
            .collect();
        out.sort_by(|a, b| a.0.cmp(&b.0));
        out
    }

    pub(crate) fn get_constant(&self, index: u32) -> Value {
        let f = self.frame();
        self.chunks[f.chunk_index].constants[index as usize].clone()
    }

    /// Resolve a chunk-level tag index to its tag ENTITY (spec EH identity).
    /// Chunk-local CALL-tag index → VM tag id, the sibling of
    /// `resolve_chunk_tag` for the Call Tags proposal. Built lazily for the
    /// same reason: every chunk-installation path must be covered.
    /// Resolve a `call_with_tag` immediate — a constant-pool index naming the
    /// tag — to its VM entity id.
    ///
    /// By NAME, because that is what identity means for a call tag: the
    /// load-time pass interns every declaration by name, so an import and the
    /// export it resolves to, or a declaration and a use in another chunk, all
    /// meet at one id.
    pub(crate) fn resolve_chunk_call_tag(
        &mut self,
        chunk_index: usize,
        name_idx: u16,
    ) -> Result<u32, VMError> {
        if self.chunk_call_tag_maps.len() < self.chunks.len() {
            self.resolve_chunk_call_tags();
        }
        if let Some(err) = self.call_tag_errors.first() {
            return Err(VMError::new(format!("invalid call tag declaration: {err}")));
        }
        let name = self
            .chunks
            .get(chunk_index)
            .and_then(|c| c.constants.get(name_idx as usize))
            .map(|v| v.to_string())
            .ok_or_else(|| VMError::new(format!("call tag name constant {name_idx} missing")))?;
        self.call_tag_registry
            .get(name.as_str())
            .copied()
            .ok_or_else(|| VMError::new(format!("undefined call tag '{name}'")))
    }

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
                            arity: decl.arity,
                        });
                        let id = self.tag_entities.len() - 1;
                        self.imported_tag_registry.insert(decl.debug_name, id);
                        id
                    }
                } else {
                    self.tag_entities.push(TagEntity {
                        debug_name: decl.debug_name,
                        arity: decl.arity,
                    });
                    self.tag_entities.len() - 1
                };
                map.push(entity);
            }
            self.chunk_tag_maps.push(map);
        }
        self.resolve_chunk_call_tags();
    }

    /// Load-time resolution for CALL tags (`proposals/call-tags`), the sibling
    /// of `resolve_chunk_tags` above.
    ///
    /// Interned by NAME across the module, so a tag declared once and used from
    /// several chunks is one identity — which is the whole contract: two funcs
    /// are distinguishable because they answer different tag *identities*, not
    /// different signatures.
    pub(crate) fn resolve_chunk_call_tags(&mut self) {
        while self.chunk_call_tag_maps.len() < self.chunks.len() {
            let ci = self.chunk_call_tag_maps.len();
            let decls = self.chunks[ci].call_tag_decls.clone();
            let mut map = Vec::with_capacity(decls.len());
            for decl in decls {
                // An interned id is only reusable while the entity it names
                // still EXISTS. `call_tags` is truncated on VM restore, so a
                // surviving registry entry can point past the end — and
                // short-circuiting on it left the name resolving to an id with
                // nothing behind it ("undefined call tag 0" from a registry
                // that plainly had the name).
                if let Some(&existing) = self.call_tag_registry.get(&decl.debug_name)
                    && (existing as usize) < self.call_tags.len()
                {
                    map.push(existing);
                    continue;
                }
                // A fallback names a function; resolve it to a callable now so
                // the unhandled-tag path costs nothing at call time.
                // A wast func is a MEMBER of the module class, not a global —
                // `ref.func $f` lowers to `Member { module_class, f }` — so the
                // handler resolves through the chunk table and is materialised
                // as a callable here, once, rather than looked up per call.
                let fallback = decl
                    .fallback
                    .as_ref()
                    .and_then(|f| self.chunk_index_for_func(f))
                    .map(|ci| self.function_value_for_chunk(ci));
                // `(canon)` interns per signature; everything else is
                // `call_tag.new` — a FRESH identity. A missing fall-back does
                // NOT make a tag canonical: the Overview puts the `?` on
                // `$func`, not on the newness, and conflating the two collapsed
                // two differently-named tags over one signature into a single
                // id, so a func handling one answered calls under the other.
                let id = if decl.canonical {
                    self.call_tag_canon(decl.params, decl.results, &decl.signature)
                } else {
                    self.call_tag_new(
                        &decl.debug_name,
                        decl.params,
                        decl.results,
                        &decl.signature,
                        fallback,
                    )
                };
                self.call_tag_registry.insert(decl.debug_name, id);
                map.push(id);
            }
            self.chunk_call_tag_maps.push(map);

            let switches = self.chunks[ci].func_switch_decls.clone();
            for (name, arms, forward) in switches {
                let resolved_arms: Vec<(u32, usize)> = arms
                    .iter()
                    .filter_map(|(tag, func)| {
                        let t = self.call_tag_registry.get(tag).copied()?;
                        let f = self.chunk_index_for_func(func)?;
                        Some((t, f))
                    })
                    .collect();
                let forward_idx = forward.as_ref().and_then(|f| self.chunk_index_for_func(f));
                if let Some(own) = self.chunk_index_for_func(&name) {
                    self.func_switches.insert(
                        own,
                        FuncSwitch {
                            arms: resolved_arms,
                            forward: forward_idx,
                        },
                    );
                }
            }

            // Tags the chunk itself declares — the compiler-generated case.
            let own_tags = self.chunks[ci].handled_call_tags.clone();
            if !own_tags.is_empty() {
                let ids: Vec<u32> = own_tags
                    .iter()
                    .map(|t| {
                        // A compiler-declared tag may not have a `(call_tag …)`
                        // field anywhere; mint it on first sight so the name is
                        // one entity regardless of who declared it first.
                        match self.call_tag_registry.get(t).copied() {
                            Some(id) if (id as usize) < self.call_tags.len() => id,
                            _ => {
                                let id = self.call_tag_new(t, 2, 1, "", None);
                                self.call_tag_registry.insert(t.clone(), id);
                                id
                            }
                        }
                    })
                    .collect();
                self.declare_func_call_tags(ci, ids);
            }

            let func_tags = self.chunks[ci].func_call_tag_decls.clone();
            for (func, tags) in func_tags {
                let ids: Vec<u32> = tags
                    .iter()
                    .filter_map(|t| self.call_tag_registry.get(t).copied())
                    .collect();
                if let Some(idx) = self.chunk_index_for_func(&func) {
                    // "each `$call_tag`'s type must be a supertype of
                    // `[ti*] -> [to*]`" — a func may only claim to handle a tag
                    // whose type its own signature satisfies. Runtime types are
                    // erased here, so a function type IS its arity pair and the
                    // subtype relation reduces to equality of that shape.
                    let (fp, fr) = {
                        let c = &self.chunks[idx];
                        (c.param_count.max(c.arity), c.result_arity)
                    };
                    for (name, id) in tags.iter().zip(ids.iter()) {
                        if let Some(def) = self.call_tags.get(*id as usize)
                            && (def.params != fp || def.results != fr)
                        {
                            self.call_tag_errors.push(format!(
                                "call tag '{name}' has type [{}->{}], which is not a supertype of func '{func}' [{fp}->{fr}]",
                                def.params, def.results
                            ));
                        }
                    }
                    self.declare_func_call_tags(idx, ids);
                } else {
                    self.call_tag_errors
                        .push(format!("(call_tag …) names unknown func '{func}'"));
                }
            }

            // "This `$func` must have the same signature as `$functype` *except*
            // also accepting an additional `funcref`" — so a fall-back handler
            // takes exactly one more parameter than its tag.
            let decls = self.chunks[ci].call_tag_decls.clone();
            for decl in decls {
                let Some(fname) = decl.fallback.as_ref() else {
                    continue;
                };
                match self.chunk_index_for_func(fname) {
                    Some(fi) => {
                        let c = &self.chunks[fi];
                        let fp = c.param_count.max(c.arity);
                        if fp != decl.params.saturating_add(1) {
                            self.call_tag_errors.push(format!(
                                "fall-back handler '{fname}' for call tag '{}' takes {fp} parameter(s); the tag's type plus the trailing funcref is {}",
                                decl.debug_name,
                                decl.params as u16 + 1
                            ));
                        }
                    }
                    None => self.call_tag_errors.push(format!(
                        "call tag '{}' names unknown fall-back handler '{fname}'",
                        decl.debug_name
                    )),
                }
            }
        }
    }

    /// Chunk index of a wast function by its declared name.
    fn chunk_index_for_func(&self, name: &str) -> Option<usize> {
        let bare = name.trim_start_matches('$');
        self.chunks
            .iter()
            .position(|c| c.name == name || c.name.trim_start_matches('$') == bare)
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
        // A fiber that re-parked inside a synchronous copy ran no code, so it
        // has no completion to record (see `VM::resume_reparked`).
        if !self.resume_reparked {
            self.last_fiber_completion = Some(completion);
        }
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
    /// Adopt the module's GLOBAL INDEX SPACE.
    ///
    /// The bytecode's `GLOBAL_GET`/`GLOBAL_SET` operands were assigned against
    /// `chunks[0].globals` at COMPILE time. So name→index here must be that
    /// table — not whatever order names happen to arrive in at run time, which
    /// is a second, disagreeing index space and reads the wrong global.
    ///
    /// Values already installed by name (host builtins bound before the module
    /// loaded) are preserved and MOVED to their table slot; a name the table
    /// does not mention keeps working by being appended, which is what host
    /// interop and `eval` rely on.
    /// MERGE an incoming chunk set's global table into this VM's, remapping the
    /// set's `GLOBAL_GET`/`GLOBAL_SET` operands as it goes.
    ///
    /// ⚠ Chunk sets are APPENDED (`run_linked_impl`) — a prelude, a bundle, the
    /// user's module, an `eval` — and each was compiled against its OWN table.
    /// Adopting one set's table wholesale therefore leaves every other set's
    /// operands indexing a space that no longer exists: `__stdlib_sorted` in
    /// the bundle and the user's own globals end up numbered from two
    /// different origins. That is the same two-index-spaces failure this whole
    /// change exists to remove, one level up.
    ///
    /// The VM's table is authoritative and only ever grows; the incoming set is
    /// rewritten into it. Same shape as `normalize_import_table`, applied at
    /// load time instead of compile time.
    /// This VM's slot for `name`, allocating one if it is new.
    fn global_slot_for(&mut self, name: &str) -> u32 {
        match self.global_index.get(name) {
            Some(&i) => i,
            None => {
                let i = self.globals.len() as u32;
                self.global_index.insert(name.to_string(), i);
                self.globals.push(Value::Null);
                self.globals_assigned.push(false);
                i
            }
        }
    }

    /// `(constant index, global name)` for every global operand in a chunk —
    /// the pre-index encoding, still emitted by paths that skip the compiler's
    /// `normalize_global_table`.
    fn global_operand_names(chunk: &Chunk) -> Vec<(u32, String)> {
        let code = &chunk.code;
        let mut out = Vec::new();
        let mut ip = 0usize;
        while ip + 3 < code.len() {
            let group = ((code[ip] as u16) << 8) | code[ip + 1] as u16;
            let sub = ((code[ip + 2] as u16) << 8) | code[ip + 3] as u16;
            let Some(op) = crate::opcode::Op::decode(group, sub) else {
                ip += 4;
                continue;
            };
            let operand_start = ip + 4;
            if (op == crate::opcode::Op::GLOBAL_GET || op == crate::opcode::Op::GLOBAL_SET)
                && operand_start + 3 < code.len()
            {
                let idx = u32::from_be_bytes([
                    code[operand_start],
                    code[operand_start + 1],
                    code[operand_start + 2],
                    code[operand_start + 3],
                ]);
                if let Some(Value::String(name)) = chunk.constants.get(idx as usize) {
                    out.push((idx, name.to_string()));
                }
            }
            ip = operand_start + op.operand_format().size_in(code, operand_start);
        }
        out
    }

    /// Install the incoming set's CANON SECTION.
    ///
    /// Unlike the global table there is nothing to remap: a canonidx is
    /// module-level and the operands that carry one already agree with the
    /// section the same compile produced. What this does have to do is refuse a
    /// SECOND, different section — two components in one program would each
    /// number their canonidx space from zero, and quietly keeping the first
    /// would make every row of the second address the wrong definition.
    ///
    /// This is `VM::canon_defs`' first and only producer. Before it existed the
    /// field was declared, read in nine places, and written by nothing, so
    /// every canon row fell through to the identity fallback where canonidx is
    /// read as a typeidx — `cancellable?` false on every row, and no `$t` or
    /// `opts` immediate ever arriving.
    pub(crate) fn merge_canon_section(&mut self, incoming: &[Chunk]) -> Result<(), VMError> {
        let Some(section) = incoming.first().map(|c| c.canon_section.clone()) else {
            return Ok(());
        };
        if section.is_empty() {
            return Ok(());
        }
        if !self.canon_defs.is_empty() && self.canon_defs[..] != section[..] {
            return Err(VMError::new(
                "canon section: this program declares two different canon sections; \
                 each numbers its canonidx space from zero, so they cannot be merged",
            ));
        }
        self.canon_defs = section.to_vec();
        Ok(())
    }

    /// Merge a component's declared TYPE SPACE at load, like
    /// [`Self::merge_canon_section`].
    ///
    /// This is `VM::canon_functypes`' first producer. It started `Vec::new()`
    /// and nothing ever appended, so `canon lift` trapped with `$ft 0 is not
    /// registered ... (have 0)` even when the source declared the type — the
    /// declaration had nowhere to go.
    ///
    /// A declared space REPLACES the four bootstrap `canon_types` entries
    /// rather than appending past them. Those four are a documented
    /// convention for source that cannot spell a type space at all (a bare
    /// `(module …)` reaching a built-in through `@N`); once a component states
    /// its own space, that space is the authority and its typeidx numbering
    /// starts at zero like every other index space.
    pub(crate) fn merge_canon_types(&mut self, incoming: &[Chunk]) -> Result<(), VMError> {
        let Some(first) = incoming.first() else {
            return Ok(());
        };
        let (fts, vts) = (first.canon_functypes.clone(), first.canon_valtypes.clone());
        if fts.is_empty() && vts.is_empty() {
            return Ok(());
        }
        // Same rule as the canon section: two declared spaces each number from
        // zero, so a second DIFFERENT one cannot be merged into the first.
        let clash = (!self.canon_functypes.is_empty() && self.canon_functypes[..] != fts[..])
            || (self.canon_types.len() != 4 && self.canon_types[..] != vts[..]);
        if clash {
            return Err(VMError::new(
                "component type space: this program declares two different type spaces; \
                 each numbers its typeidx from zero, so they cannot be merged",
            ));
        }
        self.canon_functypes = fts.to_vec();
        self.canon_types = vts.to_vec();
        if let Some(fs) = incoming.first().map(|c| c.component_funcs.clone()) {
            if !fs.is_empty() {
                self.component_funcs = fs.to_vec();
            }
        }
        Ok(())
    }

    pub(crate) fn merge_global_table(&mut self, incoming: &mut [Chunk]) {
        let Some(table) = incoming.first().map(|c| c.globals.clone()) else {
            return;
        };

        // ⚠ A set with NO table was produced by a path that did not run
        // `normalize_global_table` — a bundle, a prelude, a dynamically
        // compiled fragment. Its operands are still CONSTANT indices naming
        // the global, the pre-index encoding. Returning early here let those
        // through untouched and the VM read them as global indices, which is a
        // silent wrong-global read: every bundle helper (`__stdlib_sorted`,
        // the SQL cursor family) came back `undefined is not callable`.
        //
        // So derive the table from the operands themselves and remap anyway.
        // There is no third case: either a set carries a table or it names its
        // globals in constants, and both are handled.
        let per_chunk_names: Vec<Vec<(u32, String)>> = incoming
            .iter()
            .map(|c| Self::global_operand_names(c))
            .collect();

        let mut remap: Vec<u32> = Vec::with_capacity(table.len());
        for name in table.iter() {
            remap.push(self.global_slot_for(name));
        }

        let legacy: Vec<std::collections::HashMap<u32, u32>> = if table.is_empty() {
            per_chunk_names
                .iter()
                .map(|names| {
                    names
                        .iter()
                        .map(|(c, n)| (*c, self.global_slot_for(n)))
                        .collect()
                })
                .collect()
        } else {
            Vec::new()
        };

        for (ci, chunk) in incoming.iter_mut().enumerate() {
            let code = &mut chunk.code;
            let mut ip = 0usize;
            while ip + 3 < code.len() {
                let group = ((code[ip] as u16) << 8) | code[ip + 1] as u16;
                let sub = ((code[ip + 2] as u16) << 8) | code[ip + 3] as u16;
                let Some(op) = crate::opcode::Op::decode(group, sub) else {
                    ip += 4;
                    continue;
                };
                let operand_start = ip + 4;
                let operand_len = op.operand_format().size_in(code, operand_start);
                if (op == crate::opcode::Op::GLOBAL_GET
                    || op == crate::opcode::Op::GLOBAL_SET)
                    && operand_start + 3 < code.len()
                {
                    let old = u32::from_be_bytes([
                        code[operand_start],
                        code[operand_start + 1],
                        code[operand_start + 2],
                        code[operand_start + 3],
                    ]);
                    let mapped = if legacy.is_empty() {
                        remap.get(old as usize).copied()
                    } else {
                        legacy[ci].get(&old).copied()
                    };
                    if let Some(new_idx) = mapped {
                        let b = new_idx.to_be_bytes();
                        code[operand_start..operand_start + 4].copy_from_slice(&b);
                    }
                }
                ip = operand_start + operand_len;
            }
        }

        // The operands now index the VM's authoritative space, so the table
        // that DESCRIBES them has to be that space too. Leaving the incoming
        // compile-time table in place made every chunk claim a numbering its
        // own bytecode no longer used: the disassembler resolved a remapped
        // operand against the pre-merge names and printed `(OUT-OF-RANGE)`.
        //
        // The invariant is `chunk.globals` describes `chunk.code`'s operands.
        // Anything that rewrites those operands owns re-stating it here.
        let authoritative = std::sync::Arc::new(self.global_operand_table());
        for chunk in incoming.iter_mut() {
            chunk.globals = authoritative.clone();
        }
    }

    /// The VM's global index space as a name table, index-aligned with the
    /// slots `GLOBAL_GET`/`GLOBAL_SET` operands carry after `merge_global_table`.
    fn global_operand_table(&self) -> Vec<String> {
        let mut table = vec![String::new(); self.globals.len()];
        for (name, &slot) in &self.global_index {
            if let Some(entry) = table.get_mut(slot as usize) {
                *entry = name.clone();
            }
        }
        table
    }

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
            self.set_global(&key, Value::String(value));
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
    pub fn resolve_import_target(&self, module: &str, name: &str) -> Result<ImportTarget, VMError> {
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
        if module == "canon" {
            // `name` or `name@<canonidx>` — an index into the module's CANON
            // SECTION (`VM::canon_defs`, `Binary.md` §"Canonical Definitions").
            //
            // It used to be a raw typeidx, which could carry exactly one
            // immediate. Binary.md gives most rows more than one — `stream.read
            // t opts`, `context.get v i`, `thread.new-indirect ft tbl` — so a
            // single integer could never name them. Now the integer names a
            // ROW, and the row holds every immediate that row declares.
            //
            // Resolving it HERE, at link time, is the spec's instantiation-time
            // capture (`Store.lift`/`Store.lower`): it is what makes two
            // instantiations of one built-in distinct core funcs.
            let (bare, canon_idx) = CanonBuiltin::split_type_immediate(name);
            if let Some(b) = CanonBuiltin::by_name(bare) {
                return Ok(ImportTarget::Canon(b, canon_idx));
            }
        }
        if let Some(idx) = self.resolve_host_function_index(module, name) {
            return Ok(ImportTarget::Host(idx));
        }
        if module == "*" {
            let candidates = [name.to_string(), name.to_lowercase()];
            if let Some(global_name) = candidates
                .iter()
                .find(|g| self.has_global(g.as_str()))
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
            .find(|g| self.has_global(g.as_str()))
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
        // The chunk's own table when it has one; otherwise its MODULE's —
        // `link.rs` keeps the unified table on the module's first chunk only.
        let owner = self
            .chunk_import_owner
            .get(chunk_index)
            .copied()
            .unwrap_or(chunk_index);
        let Some(import) = self
            .chunks
            .get(chunk_index)
            .and_then(|chunk| chunk.imports.get(import_idx))
            .or_else(|| {
                self.chunks
                    .get(owner)
                    .and_then(|chunk| chunk.imports.get(import_idx))
            })
        else {
            return Ok(None);
        };
        self.resolve_import_target(&import.module, &import.name)
            .map(Some)
    }

    pub(crate) fn constant_str(&self, index: u32) -> String {
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

#[cfg(test)]
mod invoker_tests {
    use super::*;

    /// A VM is MOVED after its first host call — an embedder takes it by value
    /// and parks it in an `Rc`. The cached `callback_invoker` must still reach
    /// the VM where it now lives.
    ///
    /// It did not: the closure captured `self as *mut VM` once, so callbacks ran
    /// against the moved-out-of copy — a bitwise duplicate sharing every heap
    /// pointer with the live VM. Growing `label_stack` through it reallocated
    /// and freed the buffer the live VM still held, and the process aborted in
    /// `free` ("pointer being freed was not allocated") when that VM was
    /// dropped. Intermittent by nature: it needs a push that reallocates.
    ///
    /// Both halves are asserted, because either one alone would leave the bug:
    /// the captured cell keeps its address across the move (so the closure's
    /// capture stays valid), and the address published INSIDE it tracks the VM.
    #[test]
    fn cached_invoker_follows_the_vm_across_a_move() {
        let mut vm = VM::new();

        // First host call: builds and caches the invoker closure.
        let _ = vm.get_invoker();
        let captured_cell: *const std::cell::Cell<*mut VM> = &*vm.invoker_vm_slot;
        let before = vm.invoker_vm_slot.get();
        assert_eq!(before, &mut vm as *mut VM, "publishes its own address");

        // The move an embedder makes (`launch_gui(vm, …)`, then `Rc::new`).
        let mut moved = Box::new(vm);
        assert_ne!(
            before, &mut *moved as *mut VM,
            "the move must actually relocate the VM, or this proves nothing"
        );

        // A host call on the moved VM.
        let _ = moved.get_invoker();

        assert_eq!(
            captured_cell, &*moved.invoker_vm_slot as *const _,
            "the cell the closure captured must survive the move"
        );
        assert_eq!(
            moved.invoker_vm_slot.get(),
            &mut *moved as *mut VM,
            "the callback must reach the VM where it now lives, not the stale copy"
        );
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
        vm.set_global_owned("baseline", Value::Object(base_obj.clone()));

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
        vm.set_global_owned("script", Value::Object(a.clone()));
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
        assert!(vm.has_global("baseline"));
        let base = base_obj.lock().unwrap();
        assert_eq!(base.properties.get("boot"), Some(&Value::I32(7)));
        assert!(!base.properties.contains_key("mutated"));
        drop(base);
        // Script-added global gone.
        assert!(!vm.has_global("script"));
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
        vm.set_global_owned("keep", Value::Object(_base.clone()));
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
            vm.set_global_owned(format!("cyc{i}"), Value::Object(a));
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
