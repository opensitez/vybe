//! JavaScript source spelling -> shared protocol slot.
//!
//! JS-local by design: the shared class machinery sees only a
//! `SpecialMethodKind` (a numeric slot) and a canonical name. Which member
//! spells which role is JS's business and is decided here.

use vybe_ast::class_normalize::types::SpecialMethodKind;

/// Resolve a JS method name to `(canonical, slot?)`.
///
/// JS has no destructor syntax at all, so no spelling maps to `Destructor`.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    let (canonical, slot) = canonical_method_ref(name);
    (canonical.to_owned(), slot)
}

/// Borrow unchanged source spelling, or the immutable canonical protocol name.
/// Construction callers can transfer directly into their final owned AST field.
pub(crate) fn canonical_method_ref(name: &str) -> (&str, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    match name {
        "toString" => ("tostring", Some(ToString)),
        "valueOf" => ("valueof", Some(ValueOf)),
        "toJSON" => ("serialize", Some(Serialize)),
        // `[Symbol.iterator]` / `[Symbol.asyncIterator]` / etc. arrive from the
        // walker as pseudo-names `Symbol.iterator` after computed-key
        // resolution.
        "Symbol.iterator" => ("iterator", Some(Iterator)),
        "Symbol.asyncIterator" => ("asynciterator", Some(AsyncIterator)),
        "Symbol.toPrimitive" => ("toprimitive", Some(ToPrimitive)),
        "Symbol.hasInstance" => ("hasinstance", Some(HasInstance)),
        "Symbol.toStringTag" => ("tostringtag", None),
        _ => (name, None),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn borrowed_protocol_names_preserve_exact_spelling_and_slots() {
        use SpecialMethodKind::*;
        for (source, canonical, slot) in [
            ("toString", "tostring", Some(ToString)),
            ("valueOf", "valueof", Some(ValueOf)),
            ("toJSON", "serialize", Some(Serialize)),
            ("Symbol.iterator", "iterator", Some(Iterator)),
            ("Symbol.asyncIterator", "asynciterator", Some(AsyncIterator)),
            ("Symbol.toPrimitive", "toprimitive", Some(ToPrimitive)),
            ("Symbol.hasInstance", "hasinstance", Some(HasInstance)),
            ("Symbol.toStringTag", "tostringtag", None),
            ("ToString", "ToString", None),
            ("#private", "#private", None),
            ("méthode", "méthode", None),
        ] {
            assert_eq!(canonical_method_ref(source), (canonical, slot));
            assert_eq!(canonical_method(source), (canonical.to_owned(), slot));
        }
        let unchanged = String::from("ordinary");
        assert_eq!(canonical_method_ref(&unchanged).0.as_ptr(), unchanged.as_ptr());
    }
}
