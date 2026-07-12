//! Host-side table of `host-import → __vybe_* global` aliases.
//!
//! When the emitter bundles a stdlib chunk into the user's WASM, it
//! stores the chunk's function reference in a global named `__vybe_<op>`
//! (e.g. `__vybe_pow`, `__vybe_sorted`). At runtime, if the embedder
//! provides a native host function for the SAME operation (e.g.
//! `ecma:math::pow`), the host wants to overwrite the bundled stdlib
//! chunk with a direct reference to that native fn — so the VM picks
//! the optimised implementation and the polyfill becomes dead code.
//!
//! This table maps `(host_module, host_name)` → `__vybe_<op>` so the
//! host can walk it and install the native override in one pass. See
//! `override_stdlib_globals_with_host_fns` in `modules::mod`.
//!
//! Keep in sync with `compiler_common::bundle::MAPPINGS`: every global
//! name listed here must also exist as a stdlib chunk, otherwise the
//! override would point at an uninstalled polyfill.

pub const IMPORT_ALIASES: &[(&str, &str, &str)] = &[
    ("ecma:string", "String", "__vybe_tostring"),
    // `__vybe_count` (count substring occurrences) has no ECMA equivalent;
    // the bundled stdlib polyfill keeps running, no host override.
    // Math transcendentals removed — the `__vybe_*` polyfill chunks were dead
    // (every language routes `Math.*` → `ecma:math:*` directly), so there is
    // no stdlib global left to override.
    ("ecma:array", "toReversed", "__vybe_reversed"),
    // `__vybe_isnumeric` has no ECMA single-call equivalent
    // (`Number.isFinite(Number(s))` composes it). Polyfill keeps running.
    ("ecma:array", "splice", "__vybe_splice"),
    ("ecma:array", "slice", "__vybe_slice"),
    // `__vybe_hasproperty` retired — compiler normalises `key in obj`
    // arg order upstream and emits `ecma:object.hasOwn(obj, key)` directly.
    // `__vybe_instanceof` retired — `a instanceof TypeName` compiles to
    // `Op::REF_TEST` (WASM GC ref.test) with the type name as a const.

    // arrayInsert / arrayRemoveAt / arrayRemoveValue: dead alias entries
    // (host fns never registered). The bundled stdlib polyfills under
    // those names just keep running; no override possible until proper
    // host fns ship. Removing the entries here — add back when the
    // ecma:array equivalents (splice variants) are wired through.
];
