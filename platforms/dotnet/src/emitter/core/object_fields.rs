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

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use std::sync::Arc;

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
        chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
        let idx = chunk.add_constant(Value::String(Arc::from(*spelling)));
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, idx, line);
    }
}
