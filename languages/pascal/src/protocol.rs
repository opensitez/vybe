//! Pascal source spelling -> shared protocol slot.
//!
//! Pascal-local by design: the shared class machinery sees only a
//! `SpecialMethodKind` (a numeric slot) and a canonical name. Which member
//! spells which role is Pascal's business and is decided here.

use vybe_ast::class_normalize::types::SpecialMethodKind;

/// Resolve a Pascal member name to `(canonical, slot?)`.
///
/// Pascal resolves members case-insensitively, so everything is matched
/// lowercased. `class operator Add` arrives as `"Add"` from the
/// operator-overload grammar rule; `destructor Destroy` as `"Destroy"`.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    match name.to_lowercase().as_str() {
        "destroy" => ("destructor".into(), Some(Destructor)),
        "tostring" => ("tostring".into(), Some(ToString)),
        "add" => ("add".into(), Some(Add)),
        "subtract" => ("sub".into(), Some(Sub)),
        "multiply" => ("mul".into(), Some(Mul)),
        "divide" => ("div".into(), Some(Div)),
        "intdivide" => ("floordiv".into(), Some(FloorDiv)),
        "modulus" => ("mod".into(), Some(Mod)),
        "negative" => ("neg".into(), Some(Neg)),
        "positive" => ("pos".into(), Some(Pos)),
        "equal" => ("eq".into(), Some(Eq)),
        "notequal" => ("ne".into(), Some(Ne)),
        "lessthan" => ("lt".into(), Some(Lt)),
        "lessthanorequal" => ("le".into(), Some(Le)),
        "greaterthan" => ("gt".into(), Some(Gt)),
        "greaterthanorequal" => ("ge".into(), Some(Ge)),
        "bitwiseand" => ("and".into(), Some(And)),
        "bitwiseor" => ("or".into(), Some(Or)),
        "bitwisexor" => ("xor".into(), Some(Xor)),
        "logicalnot" | "bitwisenot" => ("not".into(), Some(Not)),
        "leftshift" => ("lshift".into(), Some(LShift)),
        "rightshift" => ("rshift".into(), Some(RShift)),
        other => (other.to_string(), None),
    }
}
