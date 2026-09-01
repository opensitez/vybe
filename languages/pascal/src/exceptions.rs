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

use vybe_compiler::primitives::namespaces::{self, CtorSpec, NamespaceNode, Subtree};

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
    ("EOutOfMemory", "EOutOfMemory"),
    ("EAccessViolation", "EAccessViolation"),
    ("EInvalidOp", "EInvalidOp"),
    ("EAssertionFailed", "EAssertionFailed"),
];

/// The identity chain a registered exception type publishes: its own spelling
/// first, then the shared canonical name it maps onto, then the root.
///
/// This is the same `__types` chain `emit_stamp_exception_ancestors` writes at
/// construction — declared here so the type is IDENTIFIABLE before anything is
/// constructed, which is what a user class naming it as a base needs.
fn ancestry(spelling: &str, canonical: &str) -> Vec<String> {
    let mut chain = vec![spelling.to_string()];
    for name in [canonical, "Exception"] {
        if !chain.iter().any(|held| held == name) {
            chain.push(name.to_string());
        }
    }
    chain
}

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
    for (spelling, canonical) in EXCEPTION_TYPES {
        classes.insert(
            spelling.to_lowercase(),
            NamespaceNode::Type {
                // IDENTITY-ONLY, exactly the shape `platforms/jvm` registers its
                // types in: empty `params`/`fields` so the shared resolver's
                // `describes_construction` test is false and construction still
                // runs through `ctor_call` — the shared exception model in
                // `primitives/errors.rs`. What the spec adds is the class
                // IDENTITY that was missing: the type is now REGISTERED, so a
                // user class naming it as a base resolves against the registry
                // like any other declared type instead of finding nothing.
                ctor: Some(CtorSpec {
                    ancestry: ancestry(spelling, canonical),
                    ..Default::default()
                }),
                ctor_call: Some(Box::new(NamespaceNode::CommonEmit(emit_key(spelling)))),
                statics: Subtree::new(),
                methods: std::collections::BTreeMap::new(),
                member_returns: std::collections::BTreeMap::new(),
            },
        );
    }
    namespaces::register_namespace_tree("pascal", NamespaceNode::Namespace(classes));
}
