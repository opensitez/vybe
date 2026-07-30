//! Kotlin source spelling -> shared protocol slot.

use vybe_ast::class_normalize::types::SpecialMethodKind;

/// Resolve a Kotlin method name to `(canonical, slot?)`.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    match name {
        "toString" => ("tostring".into(), Some(ToString)),
        "hashCode" => ("hash".into(), Some(Hash)),
        "equals" => ("eq".into(), Some(Eq)),
        "compareTo" => ("compare".into(), Some(Compare)),
        "iterator" => ("iterator".into(), Some(Iterator)),
        "next" => ("next".into(), Some(Next)),
        _ => (name.to_string(), None),
    }
}
