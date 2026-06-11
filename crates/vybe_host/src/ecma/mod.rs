//! # `ecma:*` host imports — JS runtime, ECMA-262 shape.
//!
//! Native Rust implementations of the ECMA-262 standard library surface
//! (Array, Map, Set, Date, Math, JSON, Object, typed arrays, …) plus a
//! few JS-runtime-adjacent helpers (value reflection, fixed arrays,
//! structured clone). These are **Vybe-invented** namespaces — NOT
//! blessed by any WebAssembly CG proposal. The only real `wasm:js-*`
//! names are in `crates/vybe_host/src/wasm/` (js-string-builtins,
//! js-primitive-builtins).
//!
//! The `ecma:*` namespace lets each language compile to a single
//! portable import surface that **any JS runtime** (V8, SpiderMonkey,
//! Node, etc.) can satisfy with a trivial `importObject` wrapping
//! native JS methods. Running Vybe-produced `.wasm` under such a
//! runtime needs only the matching JS shim — no Rust code at all.
//!
//! Marshaling + error-handling contract:
//! `crates/vybe_bytecode/src/wasm/JS_BUILTIN_CONVENTIONS.md`.

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
pub mod set; // §24.2  Set
pub mod string; // §22.1  String + String.prototype
pub mod symbol; // §20.4  Symbol + well-knowns
pub mod typedarray; // §23.2  TypedArray family
pub mod weakmap; // §24.3/§24.4  WeakMap + WeakSet (co-located)
pub mod weakref; // §26.1/§26.2  WeakRef + FinalizationRegistry (co-located)

// ── JS-runtime helpers (not strictly ECMA-262) ───────────────────────
pub mod fixedarray; // V8-internal fixed-length array shape
pub mod generator; // §27.3 GeneratorFunction / §27.5 Generator protocol
pub mod global; // §19 Global object (isNaN, isFinite, eval, etc.)
pub mod proxy; // §28.3 Proxy
pub mod structured_clone; // HTML spec §2.8 — not ECMA, but JS-runtime-adjacent
pub mod value; // generic JS-value reflection (Vybe extension)

use vybe_bytecode::VM;

/// Register the always-safe `ecma:*` host fns — pure computation, no
/// capabilities needed. Date is NOT included because `Date.now` /
/// `Date.parse` read the system clock; the caller gates it behind
/// `Capability::Clock` via `date::register(vm)`.
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
