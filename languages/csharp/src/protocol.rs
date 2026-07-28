//! C# source spelling -> shared protocol slot.
//!
//! C#-local by design: the shared class machinery sees only a
//! `SpecialMethodKind` (a numeric slot) and a canonical name. Which member
//! spells which role is C#'s business and is decided here.

use vybe_bytecode::class_normalize::types::SpecialMethodKind;

/// Resolve a C# member name to `(canonical, slot?)`.
///
/// C# spells its destructor as a SIGIL (`~Foo()`), so the name varies per
/// class and no table row can hold it — it is matched by prefix.
///
/// UNRESERVED NAMES: `Dispose`, `Clone` and `MoveNext` are ordinary
/// identifiers, not reserved words. Claiming a slot cannot capture the member
/// (slot keys are numeric), but the first consumer of `Exit` / `Clone` /
/// `Next` MUST check the contract before calling through.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    if name.starts_with('~') {
        return ("destructor".into(), Some(Destructor));
    }

    match name {
        "ToString" => ("tostring".into(), Some(ToString)),
        "GetHashCode" => ("hash".into(), Some(Hash)),
        "Equals" => ("eq".into(), Some(Eq)),
        "CompareTo" => ("compare".into(), Some(Compare)),
        "GetEnumerator" => ("iterator".into(), Some(Iterator)),
        "MoveNext" => ("next".into(), Some(Next)),
        // `IDisposable.Dispose` is the `using` block's release hook — the same
        // role Python spells `__exit__` and Java spells `close`.
        "Dispose" => ("exit".into(), Some(Exit)),
        "Clone" => ("clone".into(), Some(Clone)),
        // `public static T operator +(...)` arrives from the walker as source
        // name `"operator+"`; same for the others.
        "operator+" => ("add".into(), Some(Add)),
        "operator-" => ("sub".into(), Some(Sub)),
        "operator*" => ("mul".into(), Some(Mul)),
        "operator/" => ("div".into(), Some(Div)),
        "operator%" => ("mod".into(), Some(Mod)),
        "operator==" => ("eq".into(), Some(Eq)),
        "operator!=" => ("ne".into(), Some(Ne)),
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
        // C# resolves members case-sensitively at the source, but the vtable
        // key is lowercase.
        _ => (name.to_lowercase(), None),
    }
}
