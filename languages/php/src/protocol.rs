//! PHP source spelling -> shared protocol slot.
//!
//! PHP-local by design: the shared class machinery sees only a
//! `SpecialMethodKind` (a numeric slot) and a canonical name. Which magic
//! method spells which role is PHP's business and is decided here.

use vybe_ast::class_normalize::types::SpecialMethodKind;

/// Resolve a PHP method name to `(canonical, slot?)`.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    match name {
        "__destruct" => ("destructor".into(), Some(Destructor)),
        "__toString" => ("tostring".into(), Some(ToString)),
        "__debugInfo" => ("repr".into(), Some(Repr)),
        "__invoke" => ("call".into(), Some(Call)),
        // `__call` shared `Call` with `__invoke` until 2026-07-28; a class
        // defining both published one method under the other's slot.
        "__call" => ("callmissing".into(), Some(CallMissing)),
        "__callStatic" => ("callstatic".into(), Some(CallStatic)),
        "__get" => ("getattr".into(), Some(GetAttr)),
        "__set" => ("setattr".into(), Some(SetAttr)),
        "__isset" => ("hasattr".into(), Some(HasAttr)),
        "__unset" => ("delattr".into(), Some(DelAttr)),
        "__clone" => ("clone".into(), Some(Clone)),
        "__serialize" => ("serialize".into(), Some(Serialize)),
        "__unserialize" => ("deserialize".into(), Some(Deserialize)),
        // SPL interfaces are PHP's protocol surface: `Countable::count`,
        // `ArrayAccess::offset*`, `IteratorAggregate::getIterator`. They fill
        // the same roles a Python dunder does, so they resolve to the same
        // slots rather than staying PHP-shaped.
        "count" => ("len".into(), Some(Len)),
        "offsetGet" => ("getitem".into(), Some(GetItem)),
        "offsetSet" => ("setitem".into(), Some(SetItem)),
        "offsetExists" => ("hasitem".into(), Some(HasItem)),
        "offsetUnset" => ("delitem".into(), Some(DelItem)),
        "getIterator" => ("iterator".into(), Some(Iterator)),
        _ => (name.to_string(), None),
    }
}
