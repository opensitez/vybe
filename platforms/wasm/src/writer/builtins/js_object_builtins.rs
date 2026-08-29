//! # ecma:object builtins
//!
//! Host imports for `Object.*` statics and `Object.prototype.*` per
//! ECMA-262 §20.1.
//!
//! Also serves PHP's `array` (ordered string-or-int-keyed dictionary)
//! — see `JS_BUILTIN_CONVENTIONS.md` and the `appendAutoKey`
//! extension method below.
//!
//! See `JS_BUILTIN_CONVENTIONS.md` for marshaling rules.

use crate::encoding::*;

pub const MODULE: &str = "ecma:object";

pub const IMPORTS: &[&str] = &[
    // ── Construction ────────────────────────────────────────────────
    "new",         // new Object() / {}
    "create",      // Object.create(proto, props?)
    "fromEntries", // Object.fromEntries(iterable)
    "assign",      // Object.assign(target, ...sources)
    // ── Property access (indexed by string key) ─────────────────────
    "get",    // obj[key]
    "set",    // obj[key] = v
    "has",    // key in obj (walks prototype chain)
    "hasOwn", // Object.hasOwn(obj, key)
    "delete", // delete obj[key]
    // ── Enumeration ─────────────────────────────────────────────────
    "keys",                // Object.keys(obj)
    "values",              // Object.values(obj)
    "entries",             // Object.entries(obj)
    "getOwnPropertyNames", // own + non-enumerable
    "getOwnPropertySymbols",
    "length", // own enumerable key count
    // ── Descriptors ─────────────────────────────────────────────────
    "defineProperty",
    "defineProperties",
    "getOwnPropertyDescriptor",
    "getOwnPropertyDescriptors",
    // ── Prototype ───────────────────────────────────────────────────
    "getPrototypeOf",
    "setPrototypeOf",
    // ── Locking ─────────────────────────────────────────────────────
    "freeze",
    "isFrozen",
    "seal",
    "isSealed",
    "preventExtensions",
    "isExtensible",
    // ── Comparison ──────────────────────────────────────────────────
    "is", // Object.is(a, b) — SameValue
    // ── Prototype methods ───────────────────────────────────────────
    "hasOwnProperty",
    "isPrototypeOf",
    "propertyIsEnumerable",
    "toString",
    "toLocaleString",
    "valueOf",
    // ── Vybe PHP-array extension ────────────────────────────────────
    "appendAutoKey", // PHP $a[] = x — adds with next int key
];

pub fn write_signature(out: &mut Vec<u8>, name: &str) -> bool {
    match name {
        "new" => {
            write_leb128_u32(out, 0);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "create" | "fromEntries" => {
            // (proto/iterable) -> new_obj  (create also takes props but
            // we emit a 2-arg variant and handler handles null=no-props)
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "assign" => {
            // (target, source) -> target — pairwise; compiler chains
            // for multi-source assign.
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "get" => {
            // (obj, key) -> value (walks prototype chain per spec)
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "set" => {
            // ⛔ `[[Set]]` RETURNS A BOOLEAN (ECMA-262 §10.1.9 OrdinarySet), and
            // this declared `-> ()`. The VM pushes the flag like any other host
            // result, so the two disagreed: compiler-emitted call sites drop it
            // (correctly), while THIS writer's own `struct.set` lowering did
            // not — one import, two contradictory contracts. Declaring the
            // truthful signature makes the VM, the spec and V8 agree; the
            // writer drops the flag at its own emission sites.
            // (obj, key, value) -> boolean
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "has" | "hasOwn" | "delete" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "keys" | "values" | "entries" | "getOwnPropertyNames" | "getOwnPropertySymbols" => {
            // (obj) -> Array of the names/values/entries
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "length" => {
            // (obj) -> i32 — count of own enumerable keys
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "defineProperty" => {
            // (obj, key, descriptor) -> obj
            write_leb128_u32(out, 3);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "defineProperties" => {
            // (obj, descriptors_obj) -> obj
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "getOwnPropertyDescriptor" => {
            // (obj, key) -> descriptor_obj | undefined
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "getOwnPropertyDescriptors" => {
            // (obj) -> descriptors_obj
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "getPrototypeOf" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "setPrototypeOf" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "freeze" | "seal" | "preventExtensions" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "isFrozen" | "isSealed" | "isExtensible" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "is" => {
            // Object.is(a, b) — SameValue (NaN-aware, -0/+0 distinct)
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "hasOwnProperty" | "propertyIsEnumerable" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "isPrototypeOf" => {
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "toString" | "toLocaleString" | "valueOf" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_EXTERNREF);
        }
        "appendAutoKey" => {
            // (obj, value) -> assigned_key_as_i32
            // PHP extension — assigns next int key and returns it
            write_leb128_u32(out, 2);
            out.push(TYPE_EXTERNREF);
            out.push(TYPE_EXTERNREF);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        _ => return false,
    }
    true
}
