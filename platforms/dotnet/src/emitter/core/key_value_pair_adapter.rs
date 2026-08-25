//! `System.Collections.Generic.KeyValuePair(Of TKey, TValue)`.
//!
//! It lived nowhere. VB's walker carried the STRING `"KeyValuePair"` in four
//! places so that `For Each kvp In dict` could rewrite `.Key` / `.Value`, and
//! that was the whole implementation — `New KeyValuePair(Of K, V)(k, v)` had no
//! constructor at all and produced an object with no fields, so every `.Key`
//! read answered empty. It is a `System.*` type, so it belongs here and C# gets
//! it for free.
//!
//! A `KeyValuePair` is a .NET **struct**: equality is by value. That is not
//! restated here — the instance is stamped with the shared
//! `primitives::classes::emit_value_equality_stamp`, the same `__value_eq` mark
//! a Kotlin `data class` and a Pascal `record` carry, and `Equals` calls the
//! shared `primitives::records::emit_value_fields_equal` that the `==` operator
//! already uses.

use vybe_compiler::primitives::classes;
use vybe_compiler::primitives::collections;
use vybe_compiler::primitives::convert;
use vybe_compiler::primitives::records;
use vybe_compiler::primitives::strings;
use vybe_runtime::chunk::Chunk;
use vybe_runtime::opcode::Op;

/// Field names are written in BOTH spellings.
///
/// VB is case-insensitive and lowercases a member read; C# does not. The
/// registered leaves are matched case-insensitively, but a raw field read is
/// not, so a single `Key` would answer empty for `kv.key` and a single `key`
/// would answer empty for C#'s `kv.Key`. The sibling adapters
/// (`PropertyChangedEventArgs`, `NotifyCollectionChangedEventArgs`) already
/// write both for exactly this reason.
const KEY_FIELDS: [&str; 2] = ["Key", "key"];
const VALUE_FIELDS: [&str; 2] = ["Value", "value"];

fn set_field(chunks: &mut [Chunk], current: usize, object: u16, field: &str, value: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
    chunks[current].emit_string_const(field, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn get_field(chunks: &mut [Chunk], current: usize, object: u16, field: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
    chunks[current].emit_string_const(field, line);
    collections::emit_get(chunks, current, line);
}

/// Stack: `[key, value]` → `[pair]`.
///
/// Serves BOTH `New KeyValuePair(Of K, V)(k, v)` and the static
/// `KeyValuePair.Create(k, v)` — .NET's factory is documented as returning
/// exactly `new KeyValuePair<K,V>(key, value)`, so it is the same body rather
/// than a second one that could drift.
pub fn emit_key_value_pair_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let (key, value, pair) = (base, base + 1, base + 2);
    // Args arrive in call order, so pop the LAST one first.
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);

    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pair, line);
    for field in KEY_FIELDS {
        set_field(chunks, current, pair, field, key, line);
    }
    for field in VALUE_FIELDS {
        set_field(chunks, current, pair, field, value, line);
    }
    // A struct, so `Equals` is structural — the SHARED stamp, not a local one.
    classes::emit_value_equality_stamp(&mut chunks[current], pair, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
}

/// `.ToString()` → `[key, value]`, .NET's own format (`KeyValuePair<K,V>`
/// renders the pair in brackets with ", " between the halves — verified on
/// `tools/vbrun`, not assumed).
///
/// Stack: `[pair]` → `[string]`.
pub fn emit_key_value_pair_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let pair = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pair, line);

    // ⛔ Both halves must be COERCED before concatenation:
    // `strings::emit_concat` lowers to `wasm:js-string.concat`, which TRAPS on
    // a non-string, and a KeyValuePair's key or value is very often a number
    // (`KeyValuePair(Of Integer, String)`). `convert::emit_to_string` is the
    // shared §7.1.17 coercion the rest of the tree already uses.
    chunks[current].emit_string_const("[", line);
    get_field(chunks, current, pair, "Key", line);
    convert::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_string_const(", ", line);
    get_field(chunks, current, pair, "Value", line);
    convert::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_string_const("]", line);
    strings::emit_concat(&mut chunks[current], 5, line);
}

/// `.Equals(other)` — structural, because a `KeyValuePair` is a struct.
///
/// This calls the SAME reader the `==` operator uses
/// (`operators.rs` → `records::emit_value_fields_equal`), not a second
/// comparison written here. `object::emit_equals` — what `dotnet.object_equals`
/// falls back to — never looks at the `__value_eq` stamp, so a member `.Equals`
/// answered reference identity and called two pairs with DIFFERENT values
/// equal.
///
/// Stack: `[left, right]` → `[bool]`.
pub fn emit_key_value_pair_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let (left, right) = (base, base + 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left, line);
    records::emit_value_fields_equal(chunks, current, left, right, line);
}
