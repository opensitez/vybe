//! PHP source spelling -> shared protocol slot.
//!
//! PHP-local by design: the shared class machinery sees only a
//! `SpecialMethodKind` (a numeric slot) and a canonical name. Which magic
//! method spells which role is PHP's business and is decided here.

use vybe_ast::class_normalize::types::SpecialMethodKind;

/// Resolve a PHP method name to `(canonical, slot?)`.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    let (canonical, slot) = canonical_method_ref(name);
    (canonical.to_owned(), slot)
}

/// Borrow the source spelling or immutable canonical name until AST construction.
pub(crate) fn canonical_method_ref(name: &str) -> (&str, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    match name {
        "__destruct" => ("destructor", Some(Destructor)),
        "__toString" => ("tostring", Some(ToString)),
        "__debugInfo" => ("repr", Some(Repr)),
        "__invoke" => ("call", Some(Call)),
        // `__call` shared `Call` with `__invoke` until 2026-07-28; a class
        // defining both published one method under the other's slot.
        "__call" => ("callmissing", Some(CallMissing)),
        "__callStatic" => ("callstatic", Some(CallStatic)),
        "__get" => ("getattr", Some(GetAttr)),
        "__set" => ("setattr", Some(SetAttr)),
        "__isset" => ("hasattr", Some(HasAttr)),
        "__unset" => ("delattr", Some(DelAttr)),
        "__clone" => ("clone", Some(Clone)),
        // A protocol role is not a PHP method name. Keep these callable
        // members distinct from serialize()/deserialize() and each other.
        "__serialize" | "__sleep" => (name, Some(Serialize)),
        "__unserialize" | "__wakeup" => (name, Some(Deserialize)),
        // SPL interfaces are PHP's protocol surface: `Countable::count`,
        // `ArrayAccess::offset*`, `IteratorAggregate::getIterator`. They fill
        // the same roles a Python dunder does, so they resolve to the same
        // slots rather than staying PHP-shaped.
        "count" => ("len", Some(Len)),
        // `Ds\Collection::isEmpty()` asks emptiness directly rather than
        // through `count() === 0`. Keeps its source spelling; only the slot
        // is added. PHP's `empty()` is a language construct on a VALUE, not a
        // member, so it is not a member spelling this can map.
        "isEmpty" => ("isEmpty", Some(IsEmpty)),
        "offsetGet" => ("getitem", Some(GetItem)),
        "offsetSet" => ("setitem", Some(SetItem)),
        "offsetExists" => ("hasitem", Some(HasItem)),
        "offsetUnset" => ("delitem", Some(DelItem)),
        "getIterator" => ("iterator", Some(Iterator)),
        _ => (name, None),
    }
}
