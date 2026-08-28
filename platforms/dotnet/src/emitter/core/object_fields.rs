//! Storing a .NET PROPERTY on a value object, for every front end at once.
//!
//! ⛔ A CASE-INSENSITIVE FRONT END FOLDS THE MEMBER NAME. VB reads
//! `v.IsPowerOfTwo` as `ispoweroftwo`, so a field stored only in .NET's
//! PascalCase spelling is invisible to it — and invisible SILENTLY, answering
//! `undefined` rather than failing. C# sees the same object and reads the
//! field fine, which is what makes the bug look language-specific when it is
//! purely a spelling one. (Measured on `BigInteger`: `__type` and `__bi` read
//! back in both languages while `IsZero`/`IsEven`/`Sign` answered `undefined`
//! in VB alone — the two that worked are already lowercase.)
//!
//! Storing both spellings is what several adapters here already do inline
//! (`datetime_adapter`'s `["Ticks", "ticks"]` loops, `timezone_adapter`'s own
//! private copy). This is that rule in one place, so a new adapter inherits it
//! instead of rediscovering it.
//!
//! This is NOT the case-folding POLICY — that is a directive, and
//! `namespaces.rs` owns which queries fold. This is about a raw struct field,
//! which no resolver ever sees.

use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, ObjSource, PlainNames, ResolvedSlot, ValueSource,
};
use vybe_runtime::Chunk;

/// An engine-internal slot on a .NET value object — `__ms_pos`, `__content`,
/// `__bi` and the rest of this platform's representation state.
///
/// ⚠ `Internal`, never `InstanceField`. Only the variants carrying a `class`
/// reach the canonicalising path; these keys are literal storage names and a
/// canonicalised one is a DIFFERENT key, which reads back `undefined` rather
/// than failing. The `__ms_*` trio is the reason this is scoped by
/// representation rather than by adapter: binary_io, memory_stream and
/// stream_io all address the same three slots.
pub fn field_slot(key: &str) -> ResolvedSlot {
    class_slots::resolve(&ClassSlot::internal(key), &PlainNames)
}

/// Store the value in `value_slot` on the object in `obj_slot` under `key`,
/// under both the declared spelling and its folded one.
pub fn set_both_spellings(
    chunk: &mut Chunk,
    obj_slot: u16,
    value_slot: u16,
    key: &str,
    line: u32,
) {
    let folded = key.to_ascii_lowercase();
    let spellings: &[&str] = if folded == key {
        &[key]
    } else {
        &[key, folded.as_str()]
    };
    for spelling in spellings {
        class_slots::emit_class_set(
            chunk,
            ObjSource::Local(obj_slot),
            &field_slot(spelling),
            ValueSource::Local(value_slot),
            line,
        );
    }
}
