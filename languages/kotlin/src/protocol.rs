//! Kotlin source spelling -> shared protocol slot.
//!
//! Kotlin-local by design: the shared class machinery sees only a
//! `SpecialMethodKind` (a numeric slot) and a canonical name. Which member
//! spells which role is Kotlin's business and is decided here.
//!
//! OPERATOR CONVENTIONS: Kotlin's operator names (`plus`, `get`, `invoke`,
//! `contains`, …) are ordinary identifiers — a class may declare a plain
//! `fun plus(other: Counter)` that must NOT define `+`. The `operator`
//! modifier is what makes the difference, so the walker hands those methods
//! over prefixed with `"operator "` and this is the only place that prefix is
//! understood. The canonical name drops it again: the member stays reachable
//! as `a.plus(b)`, exactly as Kotlin allows, and only the slot is added.

use vybe_ast::class_normalize::types::SpecialMethodKind;

/// Resolve a Kotlin method name to `(canonical, slot?)`.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    if let Some(op) = name.strip_prefix("operator ") {
        let slot = match op {
            // Arithmetic
            "plus" => Some(Add),
            "minus" => Some(Sub),
            "times" => Some(Mul),
            "div" => Some(Div),
            "rem" | "mod" => Some(Mod),
            "unaryMinus" => Some(Neg),
            "unaryPlus" => Some(Pos),
            // Augmented assignment — Kotlin's `plusAssign` family mutates in
            // place, which is exactly what the `I*` slots are for; without
            // them `a += b` silently degrades into `a = a + b`.
            "plusAssign" => Some(IAdd),
            "minusAssign" => Some(ISub),
            "timesAssign" => Some(IMul),
            "divAssign" => Some(IDiv),
            "remAssign" | "modAssign" => Some(IMod),
            // Comparison / equality
            "compareTo" => Some(Compare),
            "equals" => Some(Eq),
            "not" => Some(Not),
            // Container protocol
            "get" => Some(GetItem),
            "set" => Some(SetItem),
            "contains" => Some(Contains),
            // Callable
            "invoke" => Some(Call),
            // Iteration
            "iterator" => Some(Iterator),
            "next" => Some(Next),
            // `inc`/`dec` and `rangeTo` have no shared slot: `++` is a rebind
            // in Kotlin rather than an operator on a value, and ranges are
            // built by `common:collections.range_*`. Both keep their source
            // name and stay ordinary methods.
            _ => None,
        };
        return (op.to_string(), slot);
    }

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
