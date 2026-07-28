//! Fortran source spelling -> shared protocol slot.
//!
//! Fortran-local by design: the shared class machinery sees only a
//! `SpecialMethodKind` (a numeric slot) and a canonical name. Which binding
//! spells which role is Fortran's business and is decided here.

use vybe_bytecode::class_normalize::types::SpecialMethodKind;

/// Resolve a Fortran type-bound binding name to `(canonical, slot?)`.
///
/// The walker normalises an `interface_designator` before this point, so a
/// generic binding arrives already spelled `operator(+)` / `assignment(=)` /
/// `write(formatted)` rather than as a bare sigil. Fortran is
/// case-insensitive, so everything is matched lowercased.
///
/// NOT handled here: `final :: cleanup` — Fortran's destructor. A FINAL
/// binding carries the user's own subroutine name, so no table row can hold
/// it; the walker has to state the role (`Modifiers::protocol_slot`) the way
/// C# does for its `~Foo()` sigil.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    match name.to_lowercase().as_str() {
        "operator(+)" => ("add".into(), Some(Add)),
        "operator(-)" => ("sub".into(), Some(Sub)),
        "operator(*)" => ("mul".into(), Some(Mul)),
        "operator(/)" => ("div".into(), Some(Div)),
        "operator(**)" => ("pow".into(), Some(Pow)),
        "operator(==)" | "operator(.eq.)" => ("eq".into(), Some(Eq)),
        "operator(/=)" | "operator(.ne.)" => ("ne".into(), Some(Ne)),
        "operator(<)" | "operator(.lt.)" => ("lt".into(), Some(Lt)),
        "operator(<=)" | "operator(.le.)" => ("le".into(), Some(Le)),
        "operator(>)" | "operator(.gt.)" => ("gt".into(), Some(Gt)),
        "operator(>=)" | "operator(.ge.)" => ("ge".into(), Some(Ge)),
        "operator(.and.)" => ("and".into(), Some(And)),
        "operator(.or.)" => ("or".into(), Some(Or)),
        "operator(.not.)" => ("not".into(), Some(Not)),
        "operator(.xor.)" | "operator(.neqv.)" => ("xor".into(), Some(Xor)),
        // Defined I/O: `write(formatted)` IS how a Fortran derived type says
        // "here is my text representation" — the same role Python spells
        // `__str__` and Ruby spells `to_s`.
        "write(formatted)" | "write(unformatted)" => ("tostring".into(), Some(ToString)),
        "read(formatted)" | "read(unformatted)" => ("deserialize".into(), Some(Deserialize)),
        // `operator(//)` is string concatenation and `assignment(=)` is
        // defined assignment. Neither is a role in the shared vocabulary, so
        // neither claims a slot rather than being forced into a near-miss.
        other => (other.to_string(), None),
    }
}
