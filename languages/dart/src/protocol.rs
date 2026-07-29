//! Dart source spelling -> shared protocol slot.
//!
//! Dart-local by design: the shared class machinery sees only a
//! `SpecialMethodKind` (a numeric slot) and a canonical name. Which member
//! spells which role is Dart's business and is decided here.

use vybe_ast::class_normalize::types::SpecialMethodKind;

/// Resolve a Dart member name to `(canonical, slot?)`.
///
/// Dart finalisation is `Finalizer` from `dart:ffi`, not a member, so no
/// spelling maps to `Destructor`.
///
/// UNRESERVED NAMES: `contains`, `call` and `iterator` are ordinary
/// identifiers. Claiming a slot cannot capture the member (slot keys are
/// numeric — this is precisely what the deleted synonym table got wrong, which
/// is why the walker had to rename a user's `add` out of the way), but the
/// first consumer of `Contains` MUST check the contract rather than assume the
/// signature.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    match name {
        "toString" => ("tostring".into(), Some(ToString)),
        "hashCode" => ("hash".into(), Some(Hash)),
        "call" => ("call".into(), Some(Call)),
        "noSuchMethod" => ("callmissing".into(), Some(CallMissing)),
        "compareTo" => ("compare".into(), Some(Compare)),
        // `iterator` and `contains` keep their own canonical spelling, so the
        // member is still stored under the name Dart code calls; only the slot
        // is added. `length` is deliberately absent — it is a property, and a
        // property's canonical name IS its storage key, so mapping it to `len`
        // would move `obj.length` out from under the source name.
        "iterator" => ("iterator".into(), Some(Iterator)),
        "moveNext" => ("next".into(), Some(Next)),
        "contains" => ("contains".into(), Some(Contains)),
        // `operator +` arrives from the walker as `"operator+"`; unary minus
        // as `"operator-@unary"` so it does not collide with binary `-`.
        "operator+" => ("add".into(), Some(Add)),
        "operator-" => ("sub".into(), Some(Sub)),
        "operator*" => ("mul".into(), Some(Mul)),
        "operator/" => ("div".into(), Some(Div)),
        "operator~/" => ("floordiv".into(), Some(FloorDiv)),
        "operator%" => ("mod".into(), Some(Mod)),
        "operator-@unary" => ("neg".into(), Some(Neg)),
        "operator==" => ("eq".into(), Some(Eq)),
        "operator<" => ("lt".into(), Some(Lt)),
        "operator<=" => ("le".into(), Some(Le)),
        "operator>" => ("gt".into(), Some(Gt)),
        "operator>=" => ("ge".into(), Some(Ge)),
        "operator&" => ("and".into(), Some(And)),
        "operator|" => ("or".into(), Some(Or)),
        "operator^" => ("xor".into(), Some(Xor)),
        "operator~" => ("not".into(), Some(Not)),
        "operator<<" => ("lshift".into(), Some(LShift)),
        "operator>>" => ("rshift".into(), Some(RShift)),
        "operator[]" => ("getitem".into(), Some(GetItem)),
        "operator[]=" => ("setitem".into(), Some(SetItem)),
        _ => (name.to_string(), None),
    }
}
