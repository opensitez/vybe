//! # `canon` — Component Model canonical built-ins
//!
//! The canonical ABI's built-ins reach a core module as ordinary `(core func)`
//! IMPORTS under module `"canon"` (the old `0xF0` instruction prefix is
//! retired), so they need real signatures like any other import.
//!
//! ⛔ WITHOUT THIS TABLE THEY WERE TYPED BY AN ARITY SCAN, AND EVERY ONE WAS
//! WRONG. The fallback declares `(externref …) -> externref` from the call
//! site's argument count, which for `canon stream.new` said `(result externref)`
//! where `CanonicalABI.md` says `(func (result i64))`. Our VM hid it — it
//! coerces — but the packed handle then reached `i64.shr_u` as an externref and
//! V8 rejected the module. `stream.drop-readable`/`drop-writable` were given a
//! result they do not have, which would have left a value on the stack at every
//! call had anything checked.
//!
//! Signatures below are the canonical ABI's own, cross-checked against the
//! emitters in `crates/vybe_compiler/src/primitives/io.rs`:
//!
//! | built-in                        | signature                          |
//! |---------------------------------|------------------------------------|
//! | `stream.new`                    | `() -> i64` (`ri \| (wi << 32)`)   |
//! | `stream.read` / `stream.write`  | `(i32 i32 i32) -> i32`             |
//! | `stream.cancel-read`            | `(i32) -> i32`                     |
//! | `stream.drop-readable/-writable`| `(i32) -> ()`                      |
//! | `resource.rep`                  | `(i32) -> i32`                     |
//!
//! ⚠ A built-in may carry a `@N` canonical suffix (`stream.read@0`) naming the
//! options block it was lowered with. The suffix picks the lowering, not the
//! type, so it is stripped before lookup.

use crate::encoding::*;

pub const MODULE: &str = "canon";

/// All `canon` built-ins the emitters import. Kept in step with the
/// `add_import("canon", …)` sites in `primitives/io.rs` and `primitives/fs_path.rs`.
pub const IMPORTS: &[&str] = &[
    "stream.new",
    "stream.read",
    "stream.write",
    "stream.cancel-read",
    "stream.drop-readable",
    "stream.drop-writable",
    "resource.rep",
];

/// Emit the WASM function signature for the given built-in, appending to `out`.
/// Returns `true` when the name is recognised. The caller has already pushed
/// the `TYPE_FUNC` tag byte.
pub fn write_signature(out: &mut Vec<u8>, name: &str) -> bool {
    // `stream.read@0` and `stream.read` are the same function type.
    let base = name.split('@').next().unwrap_or(name);
    match base {
        // ONE i64: readable end in the low 32 bits, writable in the high 32.
        "stream.new" | "future.new" => {
            write_leb128_u32(out, 0);
            write_leb128_u32(out, 1);
            out.push(TYPE_I64);
        }
        // (handle, ptr, num-elems) -> packed CopyResult.
        "stream.read" | "stream.write" => {
            write_leb128_u32(out, 3);
            out.push(TYPE_I32);
            out.push(TYPE_I32);
            out.push(TYPE_I32);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        "stream.cancel-read" | "stream.cancel-write" | "resource.rep" | "resource.drop" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
        }
        // ⛔ NO RESULT. `io.rs` emits no `DROP` after these two, which is the
        // call sites agreeing with the spec and disagreeing with the arity
        // scan's invented `-> externref`.
        "stream.drop-readable" | "stream.drop-writable" => {
            write_leb128_u32(out, 1);
            out.push(TYPE_I32);
            write_leb128_u32(out, 0);
        }
        _ => return false,
    }
    true
}
