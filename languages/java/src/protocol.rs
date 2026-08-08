//! Java source spelling -> shared protocol slot.
//!
//! Java-local by design: the shared class machinery sees only a
//! `SpecialMethodKind` (a numeric slot) and a canonical name. Which member
//! spells which role is Java's business and is decided here.

use vybe_ast::class_normalize::types::SpecialMethodKind;

/// Resolve a Java method name to `(canonical, slot?)`.
///
/// Java objects are GC-finalised with no member destructor form, so no
/// spelling maps to `Destructor`.
///
/// UNRESERVED NAMES: `next`, `close`, `clone`, `intValue` and `doubleValue`
/// are ordinary identifiers, not reserved words — a class can define `next()`
/// meaning something other than iteration. Claiming a slot cannot capture the
/// member (slot keys are numeric), but the first consumer of `Next` / `Exit` /
/// `Clone` MUST check the contract before calling through, or re-verify these
/// rows against the declared interface. The same reasoning is why JS `next` is
/// deliberately absent from `languages/js/src/protocol.rs`.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    match name {
        "toString" => ("tostring".into(), Some(ToString)),
        "hashCode" => ("hash".into(), Some(Hash)),
        "equals" => ("eq".into(), Some(Eq)),
        "compareTo" => ("compare".into(), Some(Compare)),
        "iterator" => ("iterator".into(), Some(Iterator)),
        "next" => ("next".into(), Some(Next)),
        // `Collection.isEmpty()` / `String.isEmpty()` — the shared emptiness
        // role. Keeps its SOURCE spelling as the canonical name: only the SLOT
        // is added, so `x.isEmpty()` still resolves by name and nothing moves.
        "isEmpty" => ("isEmpty".into(), Some(IsEmpty)),
        // `AutoCloseable.close` is the try-with-resources release hook — the
        // same role Python spells `__exit__` and C# spells `Dispose`.
        "close" => ("exit".into(), Some(Exit)),
        "clone" => ("clone".into(), Some(Clone)),
        "intValue" => ("int".into(), Some(Int)),
        "doubleValue" => ("float".into(), Some(Float)),
        _ => (name.to_string(), None),
    }
}
