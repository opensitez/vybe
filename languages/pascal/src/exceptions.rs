//! Pascal's exception spellings, declared as tree types.
//!
//! Pascal used to SYNTHESIZE `Exception` plus ten `E*` subclasses as Pascal
//! source in the walker — the same prelude pattern the `Generics.Collections`
//! classes used, and with the same consequence: the resulting object was an
//! ordinary Pascal class carrying none of the shared exception stamps, so a
//! Pascal `EDivByZero` could not be caught as `Exception` from PHP or Java and
//! never canonicalised to `ZeroDivisionError`.
//!
//! `primitives/errors.rs` already models exceptions for every language. What
//! was missing was only Pascal DECLARING which of its spellings means which
//! shared exception — the same shape as `protocol.rs` declaring which member
//! spelling fills which protocol slot.
//!
//! The mapping lives here, in the language, and reaches the shared model as a
//! BOUND ARGUMENT through `common:pascal.exc_*`. Nothing shared learns a
//! Pascal name: `canonical_exception_name` in `errors.rs` already answers for
//! the spellings it knows, and for `E*` names it does not know it returns the
//! name unchanged — which is the correct answer for a genuinely Pascal-only
//! exception.

use vybe_runtime::namespaces::{self, NamespaceNode, Subtree};

/// The `SysUtils` exception family, and the shared exception each one names.
///
/// A row whose second field equals the first is Pascal-only: it still goes
/// through the shared constructor (so it gets `message`, the `__types` MRO and
/// the reflection stamps), it simply has no cross-language twin.
pub const EXCEPTION_TYPES: &[(&str, &str)] = &[
    // `Exception` is Delphi's root and the shared root spelling too.
    ("Exception", "Exception"),
    ("EDivByZero", "ZeroDivisionError"),
    ("EZeroDivide", "ZeroDivisionError"),
    ("EConvertError", "ValueError"),
    ("ERangeError", "IndexError"),
    ("EArgumentException", "ValueError"),
    ("EInvalidArgument", "ValueError"),
    ("EOverflow", "OverflowError"),
    ("EIntOverflow", "OverflowError"),
    ("EFOpenError", "FileNotFoundError"),
    ("EInOutError", "IOError"),
    // No shared twin — Delphi-specific conditions.
    ("EAccessViolation", "EAccessViolation"),
    ("EInvalidOp", "EInvalidOp"),
    ("EAssertionFailed", "EAssertionFailed"),
];

/// `common:pascal.exc_<Spelling>` — the dispatch key for one exception type.
pub fn emit_key(spelling: &str) -> String {
    format!("pascal.exc_{}", spelling.to_lowercase())
}

/// Register the exception family under the `pascal` root.
///
/// Each type is a `Type` whose `ctor_call` is a `CommonEmit` — exactly how
/// `plib` declares `TList`'s constructor. `Create(msg)` therefore resolves
/// through the ordinary tree path with no Pascal-specific construction rule.
pub fn register_namespace_tree() {
    let mut classes = Subtree::new();
    for (spelling, _canonical) in EXCEPTION_TYPES {
        classes.insert(
            spelling.to_lowercase(),
            NamespaceNode::Type {
                ctor: None,
                ctor_call: Some(Box::new(NamespaceNode::CommonEmit(emit_key(spelling)))),
                statics: Subtree::new(),
                methods: std::collections::BTreeMap::new(),
                member_returns: std::collections::BTreeMap::new(),
            },
        );
    }
    namespaces::register_namespace_tree("pascal", NamespaceNode::Namespace(classes));
}
