//! Host-side table of `host-import → __vybe_* global` aliases.
//!
//! When the emitter bundles a stdlib chunk into the user's WASM, it
//! stores the chunk's function reference in a global named `__vybe_<op>`
//! (e.g. `__vybe_pow`, `__vybe_range`). At runtime, if the embedder
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
    ("vybe:convert", "toString",         "__vybe_tostring"),
    ("vybe:string",  "count",            "__vybe_count"),
    ("ecma:math",    "pow",              "__vybe_pow"),
    ("ecma:math",    "sin",              "__vybe_sin"),
    ("ecma:math",    "cos",              "__vybe_cos"),
    ("ecma:math",    "tan",              "__vybe_tan"),
    ("ecma:math",    "asin",             "__vybe_asin"),
    ("ecma:math",    "acos",             "__vybe_acos"),
    ("ecma:math",    "atan",             "__vybe_atan"),
    ("ecma:math",    "atan2",            "__vybe_atan2"),
    ("ecma:math",    "log",              "__vybe_log"),
    ("ecma:math",    "log10",            "__vybe_log10"),
    ("ecma:math",    "exp",              "__vybe_exp"),
    ("ecma:math",    "sinh",             "__vybe_sinh"),
    ("ecma:math",    "cosh",             "__vybe_cosh"),
    ("ecma:math",    "tanh",             "__vybe_tanh"),
    ("ecma:math",    "sign",             "__vybe_sign"),
    ("ecma:math",    "clamp",            "__vybe_clamp"),
    ("vybe:array",   "range",            "__vybe_range"),
    ("vybe:array",   "sorted",           "__vybe_sorted"),
    ("vybe:array",   "reversed",         "__vybe_reversed"),
    ("vybe:array",   "enumerate",        "__vybe_enumerate"),
    ("vybe:array",   "zip",              "__vybe_zip"),
    ("vybe:array",   "sum",              "__vybe_sum"),
    ("vybe:array",   "pymin",            "__vybe_min"),
    ("vybe:array",   "pymax",            "__vybe_max"),
    ("vybe:convert", "isNumeric",        "__vybe_isnumeric"),
    ("vybe:array",   "splice",           "__vybe_splice"),
    ("ecma:math",    "floor",            "__vybe_floor"),
    ("vybe:array",   "slice",            "__vybe_slice"),
    ("vybe:object",  "keys",             "__vybe_keys"),
    ("vybe:object",  "hasProperty",      "__vybe_hasproperty"),
    ("vybe:object",  "assign",           "__vybe_assign"),
    ("vybe:object",  "instanceOf",       "__vybe_instanceof"),
    ("vybe:object",  "deleteProperty",   "__vybe_deleteproperty"),
    ("vybe:array",   "from",             "__vybe_from"),
    ("vybe:array",   "redim",            "__vybe_redim"),
    ("vybe:array",   "sliceStep",        "__vybe_slicestep"),
    ("ecma:math",    "dynMul",           "__vybe_dynmul"),
    ("ecma:math",    "fmod",               "__vybe_fmod"),
    ("vybe:array",   "arrayInsert",        "__vybe_array_insert"),
    ("vybe:array",   "arrayRemoveAt",      "__vybe_array_remove_at"),
    ("vybe:array",   "arrayRemoveValue",   "__vybe_array_remove_value"),
    ("vybe:array",   "arrayLastIndexOf",   "__vybe_array_last_index_of"),
];
