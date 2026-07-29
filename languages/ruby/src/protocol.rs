//! Ruby source spelling -> shared protocol slot.
//!
//! Ruby-local by design: the shared class machinery sees only a
//! `SpecialMethodKind` (a numeric slot) and a canonical name. Which method
//! name spells which role is Ruby's business and is decided here.

use vybe_ast::class_normalize::types::SpecialMethodKind;

/// Resolve a Ruby method name to `(canonical, slot?)`.
///
/// Ruby has no member destructor form (objects are GC-finalised), so no
/// spelling maps to `Destructor` — a language with no such concept declares
/// none rather than being given one.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    match name {
        "to_s" => ("tostring".into(), Some(ToString)),
        "inspect" => ("repr".into(), Some(Repr)),
        "to_i" => ("int".into(), Some(Int)),
        "to_f" => ("float".into(), Some(Float)),
        "each" => ("iterator".into(), Some(Iterator)),
        "reverse_each" => ("reversed".into(), Some(Reversed)),
        "+" => ("add".into(), Some(Add)),
        "-" => ("sub".into(), Some(Sub)),
        "*" => ("mul".into(), Some(Mul)),
        "/" => ("div".into(), Some(Div)),
        "%" => ("mod".into(), Some(Mod)),
        "**" => ("pow".into(), Some(Pow)),
        "-@" => ("neg".into(), Some(Neg)),
        "+@" => ("pos".into(), Some(Pos)),
        "abs" => ("abs".into(), Some(Abs)),
        "==" => ("eq".into(), Some(Eq)),
        "!=" => ("ne".into(), Some(Ne)),
        "<=>" => ("compare".into(), Some(Compare)),
        "<" => ("lt".into(), Some(Lt)),
        "<=" => ("le".into(), Some(Le)),
        ">" => ("gt".into(), Some(Gt)),
        ">=" => ("ge".into(), Some(Ge)),
        "&" => ("and".into(), Some(And)),
        "|" => ("or".into(), Some(Or)),
        "^" => ("xor".into(), Some(Xor)),
        "~" => ("not".into(), Some(Not)),
        "<<" => ("lshift".into(), Some(LShift)),
        ">>" => ("rshift".into(), Some(RShift)),
        "[]" => ("getitem".into(), Some(GetItem)),
        "[]=" => ("setitem".into(), Some(SetItem)),
        "include?" => ("contains".into(), Some(Contains)),
        "size" | "length" => ("len".into(), Some(Len)),
        "call" => ("call".into(), Some(Call)),
        // Ruby's missing-method hook is PHP's `__call` and Dart's
        // `noSuchMethod` — one slot, three spellings.
        "method_missing" => ("callmissing".into(), Some(CallMissing)),
        "respond_to_missing?" => ("hasattr".into(), Some(HasAttr)),
        "initialize_copy" => ("clone".into(), Some(Clone)),
        "hash" => ("hash".into(), Some(Hash)),
        _ => (name.to_string(), None),
    }
}
