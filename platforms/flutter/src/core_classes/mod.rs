//! Flutter classes the adapter provides, synthesized as ordinary AST.
//!
//! **This is what a Flutter adapter class is.** A `ChangeNotifier` is a CLASS:
//! declared here as a `StmtKind::ClassDecl`, spliced into the module, and
//! walked down the same path a user class takes — `normalize_class` →
//! `NormalClass` → `compile_class` — so it inherits a reserved type slot, a
//! real rtt, `TypeRegistry` registration, the prototype stamp that makes
//! member dispatch RECEIVER-based, MRO, and protocol-slot binding. It is the
//! same move `languages/dart/src/core_classes` makes for `dart:core` and
//! `platforms/dotnet` makes for `System.*`.
//!
//! **What it replaces.** Flutter's catalog ([`crate::emitter::catalog`]) can
//! only describe a class STRUCTURALLY — name, parent, interfaces, fields — so
//! a class there has no behaviour at all, and `tree_register` consequently
//! registers every one of them with an empty method map. Behaviour therefore
//! had nowhere to live but `runtime.dart`, the Dart prelude that
//! `documentation/guiplan.md` says must be *deleted, not ported*. Every class
//! that moves here is behaviour leaving that prelude for a real class.
//!
//! **Why not a component descriptor.** `platforms/dotnet` declares its classes
//! through `component_descriptor` + an `emit_dispatch`, which requires the
//! platform to emit bytecode itself. `platforms/plib` shows the lighter form —
//! tree leaves pointing at SHARED primitives (`CommonEmit("collections.push")`)
//! with no platform emitter at all. Neither is needed for a class whose body
//! is ordinary Dart-shaped logic: declaring the AST is strictly less machinery
//! than either, and it keeps the behaviour readable as the Flutter semantics
//! it implements.
//!
//! Instance methods are NOT registered in the namespace tree: member dispatch
//! is receiver-based off the vtable, and `vybe_runtime::namespaces` states that
//! only `ctor` and `statics` are reachable by a path walk. `plib`'s registrar
//! records the same rule.

mod builders;
mod change_notifier;
mod notifier;
mod value_notifier;

use vybe_ast::Statement;

/// Every Flutter class the adapter declares, paired with its builder.
///
/// This is the list `is_core_class` and [`declarations_for`] read; adding a
/// class means adding a row here and a builder beside it.
pub const CORE_CLASSES: &[(&str, fn() -> Statement)] = &[
    ("ChangeNotifier", change_notifier::change_notifier),
    ("ValueNotifier", value_notifier::value_notifier),
];

/// True when `name` is one of the classes above.
pub fn is_core_class(name: &str) -> bool {
    CORE_CLASSES.iter().any(|(n, _)| *n == name)
}

/// The classes `source` actually needs, built and in declaration order.
///
/// A program that never writes `ValueNotifier` should not compile one, so a
/// class is built only when its name appears in the source — the same gate
/// `languages/dart/src/core_classes::declarations_for` applies, and the same
/// gate the Flutter runtime prelude already uses. Substring containment is
/// deliberately conservative: a name inside a comment declares a class nobody
/// uses, which costs a little time and nothing else; it cannot go the other
/// way, since a name that is constructed or extended is by definition present
/// in the text.
///
/// `is_user_declared` lets a user's own `class ChangeNotifier` win, exactly as
/// a user `class StringBuffer` wins over `dart:core`'s.
pub fn declarations_for(source: &str, is_user_declared: impl Fn(&str) -> bool) -> Vec<Statement> {
    CORE_CLASSES
        .iter()
        .filter(|(name, _)| source.contains(name) && !is_user_declared(name))
        .map(|(_, build)| build())
        .collect()
}
