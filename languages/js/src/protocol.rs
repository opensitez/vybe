//! JavaScript source spelling -> shared protocol slot.
//!
//! JS-local by design: the shared class machinery sees only a
//! `SpecialMethodKind` (a numeric slot) and a canonical name. Which member
//! spells which role is JS's business and is decided here.

use vybe_bytecode::class_normalize::types::SpecialMethodKind;

/// Resolve a JS method name to `(canonical, slot?)`.
///
/// JS has no destructor syntax at all, so no spelling maps to `Destructor`.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    match name {
        "toString" => ("tostring".into(), Some(ToString)),
        "valueOf" => ("valueof".into(), Some(ValueOf)),
        "toJSON" => ("serialize".into(), Some(Serialize)),
        // `[Symbol.iterator]` / `[Symbol.asyncIterator]` / etc. arrive from the
        // walker as pseudo-names `Symbol.iterator` after computed-key
        // resolution.
        "Symbol.iterator" => ("iterator".into(), Some(Iterator)),
        "Symbol.asyncIterator" => ("asynciterator".into(), Some(AsyncIterator)),
        "Symbol.toPrimitive" => ("toprimitive".into(), Some(ToPrimitive)),
        "Symbol.hasInstance" => ("hasinstance".into(), Some(HasInstance)),
        "Symbol.toStringTag" => ("tostringtag".into(), None),
        _ => (name.to_string(), None),
    }
}
