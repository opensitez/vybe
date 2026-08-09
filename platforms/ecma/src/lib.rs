//! `vybe_platform_ecma` — the `ecma:*` host surface, extracted from `vybe_host`.
//!
//! Native Rust implementations of the ECMA-262 standard library surface
//! (Array, Map, Set, Date, Math, JSON, Object, typed arrays, …) plus a few
//! JS-runtime-adjacent helpers (value reflection, fixed arrays, structured
//! clone). These are **Vybe-invented** `ecma:*` namespaces — NOT blessed by any
//! WebAssembly CG proposal.
//!
//! The `ecma:*` namespace lets each language compile to a single portable
//! import surface that **any JS runtime** (V8, SpiderMonkey, Node, …) can
//! satisfy with a trivial `importObject`. This crate also owns the
//! process-global shared prototypes (`Object`/`Function`/`Array`/… `.prototype`
//! singletons) and the ecma-coupled host-fn-ref helpers that stamp the shared
//! function prototype (`receiver_host_fn_ref`, `bound_host_fn_ref`).
//!
//! Marshaling + error-handling contract:
//! `crates/vybe_runtime/src/wasm/JS_BUILTIN_CONVENTIONS.md`.

// ── ECMA-262 spec types (one file per spec chapter) ──────────────────
pub mod array; // §23.1  Array
pub mod arraybuffer; // §25.1 ArrayBuffer + §25.3 DataView + SharedArrayBuffer
pub mod atomics; // §25.4  Atomics
pub mod bigint; // §21.2  BigInt
pub mod boolean; // §20.3  Boolean
pub mod date; // §21.4  Date
pub mod error; // §20.5  Error + native subclasses
pub mod function; // §20.2  Function.prototype.{bind,call,apply}
pub mod global_this; // §19.3  globalThis
pub mod intl; // §ECMA-402  Internationalization API (sub-modules)
pub mod iterator; // Stage-3  Iterator helpers
pub mod json; // §25.5  JSON
pub mod map; // §24.1  Map
pub mod math; // §21.3  Math (+ Stage-3 minOf/maxOf/sumPrecise)
pub mod number; // §21.1  Number + global parseInt/parseFloat/etc.
pub mod object; // §19.1  Object (keys/values/entries, etc.)
pub mod promise; // §27.7  Promise
pub mod reflect; // §28.1  Reflect
pub mod regexp; // §22.2  RegExp + String.prototype regex methods
pub mod scheduler;
pub mod set; // §24.2  Set
pub mod string; // §22.1  String + String.prototype
pub mod symbol; // §20.4  Symbol + well-knowns
pub mod timezone; // §ECMA-402  IANA time zone data (normative reference)
pub mod typedarray; // §23.2  TypedArray family
pub mod weakmap; // §24.3/§24.4  WeakMap + WeakSet (co-located)
pub mod weakref; // §26.1/§26.2  WeakRef + FinalizationRegistry (co-located)

// ── Global wiring (constructor↔prototype, globalThis) ────────────────
pub mod builtin_types; // TypeRegistry vtables for the JS/Intl surface; run in Plugin::finalize
pub mod ecma_globals; // stamps shared prototypes + wires constructors; run AFTER register()
pub mod plugin; // the ecma platform as one vybe_runtime::Plugin
pub use plugin::Plugin;

// ── JS-runtime helpers (not strictly ECMA-262) ───────────────────────
pub mod fixedarray; // V8-internal fixed-length array shape
pub mod generator; // §27.3 GeneratorFunction / §27.5 Generator protocol
pub mod global; // §19 Global object (isNaN, isFinite, eval, etc.)
pub mod proxy; // §28.3 Proxy
pub mod structured_clone; // HTML spec §2.8 — not ECMA, but JS-runtime-adjacent
pub mod value; // generic JS-value reflection (Vybe extension)

use std::sync::Arc;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{VM, Value};

/// Wire the ECMA global objects: stamp the shared prototypes with their
/// methods, pin each constructor↔prototype link, and install `globalThis` +
/// Symbol/Reflect/Atomics/BigInt/Iterator globals. MUST run AFTER [`register`]
/// (it resolves host functions by registry index). Moved out of vybe_host so
/// the ecma crate owns both the prototype objects and their contents.
pub fn register_globals(vm: &mut VM) {
    ecma_globals::register(vm);
}

/// Register the always-safe `ecma:*` host fns — pure computation, no
/// capabilities needed. Date is NOT included because `Date.now` / `Date.parse`
/// read the system clock; the caller gates it behind `Capability::Clock` via
/// `date::register(vm)`.
pub fn register(vm: &mut VM) {
    array::register(vm);
    arraybuffer::register(vm);
    atomics::register(vm);
    bigint::register(vm);
    boolean::register(vm);
    error::register(vm);
    function::register(vm);
    global_this::register(vm);
    iterator::register(vm);
    timezone::register(vm);
    json::register(vm);
    map::register(vm);
    math::register(vm);
    intl::register(vm);
    number::register(vm);
    object::register(vm);
    promise::register(vm);
    reflect::register(vm);
    regexp::register(vm);
    set::register(vm);
    string::register(vm);
    symbol::register(vm);
    typedarray::register(vm);
    weakmap::register(vm);
    weakref::register(vm);
    fixedarray::register(vm);
    generator::register(vm);
    global::register(vm);
    proxy::register(vm);
    structured_clone::register(vm);
    value::register(vm);
}

/// Force-initialize every process-global shared prototype (and `globalThis`)
/// so each is allocated through the tracked heap BEFORE a `VM::snapshot`, and
/// is therefore captured as part of the baseline the reset restores.
///
/// These prototypes are `static OnceLock<Arc<Mutex<Object>>>` shared across
/// every VM in the process (see `vmhotresetplan.md` bucket C, and the
/// `delete Object.prototype` poison-pill). A prototype first *touched* by a
/// script AFTER the snapshot would not be in the snapshot's baseline, so a
/// later reset's `collect_since` would wipe it — breaking every following run.
/// Priming here (on the boot thread, after `heap::enable_tracking()`, before
/// `snapshot()`) makes every prototype baseline, so script-added own-properties
/// on them are rolled back on reset instead of leaking across runs.
///
/// Every process-global `Arc<Mutex<Object>>` static in the host. Re-verified
/// host-wide: this list and the statics are in step. A new shared prototype
/// static MUST be primed here or it will leak mutations across resets.
///
/// `Map` and `Set` were missing while this comment already claimed the set was
/// complete. They were safe only incidentally — `ecma_globals::register` runs
/// at boot and touches them on its way to wiring the `Map`/`Set` globals, so
/// they landed pre-snapshot anyway. Priming them here makes it hold by
/// construction: were that global ever made lazy, the first tenant would
/// allocate `%Map.prototype%` post-snapshot, the next reset would gut it, and
/// every later tenant would get Maps with no methods — the exact defect the
/// lazily-built Error constructor cache had.
pub fn prime_shared_prototypes() {
    let _ = object::shared_object_prototype();
    let _ = function::shared_function_prototype();
    let _ = array::shared_array_prototype();
    let _ = string::shared_string_prototype();
    let _ = number::shared_number_prototype();
    let _ = boolean::shared_boolean_prototype();
    let _ = date::shared_date_prototype();
    let _ = regexp::shared_regexp_prototype();
    let _ = map::shared_map_prototype();
    let _ = set::shared_set_prototype();
    let _ = intl::shared_collator_prototype();
    let _ = intl::shared_number_format_prototype();
    let _ = intl::shared_date_time_format_prototype();
    let _ = intl::shared_relative_time_format_prototype();
    let _ = intl::shared_segmenter_prototype();
    let _ = global_this::shared_singleton();
}

/// Create a HostFunction Value for a receiver-style instance method.
///
/// The `__vybe_method_receiver` marker keeps property-resolved host methods on
/// the same calling convention as TypeRegistry-resolved methods: the VM passes
/// the receiver as the first argument when the function is invoked. Stamps the
/// shared function prototype, so it lives with ecma.
pub fn receiver_host_fn_ref(module: &str, name: &str, idx: usize) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__host_module".into(), Value::String(Arc::from(module)));
    obj.properties
        .insert("__host_name".into(), Value::String(Arc::from(name)));
    obj.properties
        .insert("__host_idx".into(), Value::F64(idx as f64));
    obj.properties
        .insert("__vybe_method_receiver".into(), Value::Bool(true));
    obj.properties
        .insert("__proto__".into(), function::shared_function_prototype());
    obj.properties
        .insert("name".into(), Value::String(Arc::from(name)));
    obj.kind = ObjectKind::HostFunction(idx);
    Value::Object(vybe_runtime::heap::alloc(obj))
}

/// Create a HostFunction Value with bound args (Function.prototype.bind shape).
/// When the resulting ref is called, `bound_args` are prepended to the runtime
/// args before the host fn runs. Standard ECMA-262 §20.2.3.2 semantics. The
/// VM-side dispatch lives in `vybe_runtime/src/calls.rs` HostFunction arm.
#[allow(dead_code)]
pub fn bound_host_fn_ref(vm: &VM, module: &str, name: &str, bound_args: Vec<Value>) -> Value {
    if let Some(&idx) = vm
        .host_registry
        .get(&(module.to_string(), name.to_string()))
    {
        let mut obj = Object::new();
        obj.properties
            .insert("__host_module".into(), Value::String(Arc::from(module)));
        obj.properties
            .insert("__host_name".into(), Value::String(Arc::from(name)));
        obj.properties
            .insert("__host_idx".into(), Value::F64(idx as f64));
        obj.properties
            .insert("__proto__".into(), function::shared_function_prototype());
        obj.properties
            .insert("name".into(), Value::String(Arc::from(name)));
        obj.properties.insert(
            "__bound_args".into(),
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(bound_args))),
        );
        obj.kind = ObjectKind::HostFunction(idx);
        Value::Object(vybe_runtime::heap::alloc(obj))
    } else {
        Value::Null
    }
}
