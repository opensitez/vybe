//! Go source spelling -> shared protocol slot.
//!
//! Go-local by design: the shared class machinery sees only a
//! `SpecialMethodKind` (a numeric slot) and a canonical name.
//!
//! Go has no `__dunder__` surface at all — a type fills a role by having a
//! method with the right NAME and signature, which is how it satisfies an
//! interface (`fmt.Stringer` is `String() string`, nothing more). So every
//! row here is an ordinary exported identifier, and the caveat that applies
//! to one or two rows elsewhere applies to ALL of them: a numeric slot key
//! cannot capture the member, but a consumer of one of these roles must check
//! the contract rather than assume the signature.

use vybe_ast::class_normalize::types::SpecialMethodKind;

/// Resolve a Go method name to `(canonical, slot?)`.
///
/// Go has no destructor — cleanup is `defer` at the call site and finalizers
/// are `runtime.SetFinalizer`, neither of which is a member.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    match name {
        // `fmt.Stringer`. NOT `Error() string` — the `error` interface is a
        // separate method a type can declare ALONGSIDE `String()`, and one
        // slot cannot hold two methods.
        "String" => ("tostring".into(), Some(ToString)),
        // `sort.Interface` is `Len`/`Less`/`Swap`. `Swap` has no role.
        "Len" => ("len".into(), Some(Len)),
        "Less" => ("lt".into(), Some(Lt)),
        "Compare" => ("compare".into(), Some(Compare)),
        "Equal" => ("eq".into(), Some(Eq)),
        // `io.Closer` — the release half of a `defer c.Close()` pair, the same
        // role Python spells `__exit__` and Java spells `close`.
        "Close" => ("exit".into(), Some(Exit)),
        // `json.Marshaler` / `json.Unmarshaler`. `MarshalText` is deliberately
        // absent: a type may declare both, and they would collide on one slot.
        "MarshalJSON" => ("serialize".into(), Some(Serialize)),
        "UnmarshalJSON" => ("deserialize".into(), Some(Deserialize)),
        "Clone" => ("clone".into(), Some(Clone)),
        _ => (name.to_string(), None),
    }
}
