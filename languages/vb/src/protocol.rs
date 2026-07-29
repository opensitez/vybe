//! VB source spelling -> shared protocol slot.
//!
//! VB-local by design: the shared class machinery sees only a
//! `SpecialMethodKind` (a numeric slot) and a canonical name. Which member
//! spells which role is VB's business and is decided here.

use vybe_ast::class_normalize::types::SpecialMethodKind;

/// Resolve a VB member name to `(canonical, slot?)`.
///
/// VB resolves members case-insensitively, so the fallthrough canonical name
/// is lowercased. `Finalize` is VB's destructor.
///
/// UNRESERVED NAMES: `Dispose`, `Clone` and `MoveNext` are ordinary
/// identifiers. Claiming a slot cannot capture the member (slot keys are
/// numeric), but the first consumer of `Exit` / `Clone` / `Next` MUST check
/// the contract before calling through.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    // Operator overloads and the dunder spellings the VB walker emits are
    // matched case-SENSITIVELY, before the case-insensitive member table.
    let sigil = match name {
        "__add__" => Some(("add", Add)),
        "__sub__" => Some(("sub", Sub)),
        "__mul__" => Some(("mul", Mul)),
        "__truediv__" => Some(("div", Div)),
        "__floordiv__" => Some(("floordiv", FloorDiv)),
        "__mod__" => Some(("mod", Mod)),
        "__neg__" => Some(("neg", Neg)),
        "__eq__" => Some(("eq", Eq)),
        "__ne__" => Some(("ne", Ne)),
        "__lt__" => Some(("lt", Lt)),
        "__le__" => Some(("le", Le)),
        "__gt__" => Some(("gt", Gt)),
        "__ge__" => Some(("ge", Ge)),
        "__bitnot__" => Some(("not", Not)),
        "__istrue__" | "__isfalse__" => Some(("bool", Bool)),
        "__getitem__" => Some(("getitem", GetItem)),
        "__setitem__" => Some(("setitem", SetItem)),
        "__call__" => Some(("call", Call)),
        "operator+" => Some(("add", Add)),
        "operator-" => Some(("sub", Sub)),
        "operator*" => Some(("mul", Mul)),
        "operator/" => Some(("div", Div)),
        // `\` is VB's integer division — truncating, so `FloorDiv`, not the
        // same slot as `/`.
        "operator\\" => Some(("floordiv", FloorDiv)),
        "operatorMod" | "operatormod" => Some(("mod", Mod)),
        "operator=" => Some(("eq", Eq)),
        // `<>` is VB's NOT-equal. It resolved to `Eq` until 2026-07-28, which
        // put `=` and `<>` on one slot with opposite meanings.
        "operator<>" => Some(("ne", Ne)),
        "operator<" => Some(("lt", Lt)),
        "operator<=" => Some(("le", Le)),
        "operator>" => Some(("gt", Gt)),
        "operator>=" => Some(("ge", Ge)),
        "operatorAnd" | "operatorand" => Some(("and", And)),
        "operatorOr" | "operatoror" => Some(("or", Or)),
        "operatorXor" | "operatorxor" => Some(("xor", Xor)),
        "operatorNot" | "operatornot" => Some(("not", Not)),
        _ => None,
    };
    if let Some((canonical, kind)) = sigil {
        return (canonical.to_string(), Some(kind));
    }

    match name.to_lowercase().as_str() {
        "finalize" => ("destructor".into(), Some(Destructor)),
        "tostring" => ("tostring".into(), Some(ToString)),
        "gethashcode" => ("hash".into(), Some(Hash)),
        "equals" => ("eq".into(), Some(Eq)),
        "compareto" => ("compare".into(), Some(Compare)),
        "getenumerator" => ("iterator".into(), Some(Iterator)),
        "movenext" => ("next".into(), Some(Next)),
        "dispose" => ("exit".into(), Some(Exit)),
        "clone" => ("clone".into(), Some(Clone)),
        other => (other.to_string(), None),
    }
}
