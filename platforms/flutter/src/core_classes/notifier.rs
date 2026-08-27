//! State shared by the `foundation.dart` notifier classes.
//!
//! `ChangeNotifier` and `ValueNotifier` publish the same observable surface —
//! in real Flutter the second EXTENDS the first. The members are built here
//! once so the two class files declare the same behaviour without an
//! inheritance link between two synthesized classes, which would make them
//! order-dependent in the spliced body for no gain.

use super::builders::*;
use vybe_ast::{ClassMember, Expression, Statement};

/// The listener list. Underscore-prefixed so it reads as private in Dart and
/// cannot collide with a member on a user subclass.
pub(super) const LISTENERS: &str = "_vybeListeners";
/// Set by `dispose()`. Dart throws on any use after disposal.
pub(super) const DISPOSED: &str = "_vybeDisposed";

pub(super) fn listeners() -> Expression {
    this_field(LISTENERS)
}

fn disposed() -> Expression {
    this_field(DISPOSED)
}

/// `if (this._vybeDisposed) throw StateError('… after being disposed.');`
///
/// Dart 3.10.4 throws a `FlutterError` whose `toString` begins "A <Type> was
/// used after being disposed". `StateError` is the closest exception the dart
/// frontend already declares, and the corpus only asserts that SOMETHING is
/// thrown — so this stays honest about being an approximation rather than
/// minting an unregistered type.
pub(super) fn throw_if_disposed() -> Statement {
    if_stmt(
        disposed(),
        vec![throw(call_value(
            ident("StateError"),
            vec![str_lit("A ChangeNotifier was used after being disposed.")],
        ))],
    )
}

/// Every backing field, assigned explicitly in the constructor.
///
/// A field INITIALISER is not enough: with an explicit constructor declared it
/// did not run, so `this._vybeListeners` was undefined and `hasListeners` threw
/// before any test reached its assertion. `languages/dart`'s core classes set
/// their fields in the constructor for the same reason.
pub(super) fn init_state() -> Vec<Statement> {
    vec![
        assign(listeners(), empty_list()),
        assign(disposed(), bool_lit(false)),
    ]
}

/// The observable surface both notifier classes publish.
pub(super) fn members() -> Vec<ClassMember> {

    vec![
        field(LISTENERS, "List", empty_list()),
        field(DISPOSED, "bool", bool_lit(false)),
        // `bool get hasListeners` — a GETTER, matching Dart's spelling. It is
        // `@protected` in Flutter, which the corpus reads directly; there is no
        // protection model here to enforce it.
        // ⛔The FINAL spelling, not the source one. These bodies are spliced
        // in after the walk, so nothing normalises them: a plain
        // `this._vybeListeners.length` stays a raw member read of an ARRAY and
        // never becomes the length builtin, which is what made this getter
        // throw. `__dart_is_not_empty` is the emitted form and says exactly
        // what `hasListeners` means.
        getter(
            "hasListeners",
            "bool",
            // An ACCESSOR body — the receiver is the identifier form.
            vec![ret(call_value(
                ident("__dart_is_not_empty"),
                vec![accessor_field(LISTENERS)],
            ))],
        ),
        // Duplicate registrations are KEPT: Flutter documents that adding the
        // same closure twice makes it fire twice, and `removeListener` removes
        // only ONE registration. A plain list gives both for free — a set would
        // silently break the duplicate tests.
        method(
            "addListener",
            vec![param("listener", None)],
            Some("void"),
            vec![
                throw_if_disposed(),
                expr_stmt(call_member(listeners(), "add", vec![ident("listener")])),
            ],
        ),
        // `remove` on the shared list drops the FIRST equal element, which is
        // exactly "one registration".
        method(
            "removeListener",
            vec![param("listener", None)],
            Some("void"),
            vec![expr_stmt(call_member(
                listeners(),
                "remove",
                vec![ident("listener")],
            ))],
        ),
        // Iterate a COPY. A listener is allowed to add or remove listeners —
        // including disposing the notifier — while being notified, and walking
        // the live list while it mutates skips or repeats callbacks. Flutter
        // handles this with a growable copy + a null-out scheme; a copy is the
        // same guarantee in one line.
        method(
            "notifyListeners",
            vec![],
            Some("void"),
            vec![
                throw_if_disposed(),
                local("__vybeSnapshot", call_member(listeners(), "toList", vec![])),
                for_in(
                    "__vybeListener",
                    ident("__vybeSnapshot"),
                    vec![expr_stmt(call_value(ident("__vybeListener"), vec![]))],
                ),
            ],
        ),
        // Dispose drops the listeners so a disposed notifier holds no
        // references — the "memory leak prevention" the corpus checks — and
        // arms the guard.
        method(
            "dispose",
            vec![],
            Some("void"),
            vec![
                assign(disposed(), bool_lit(true)),
                assign(listeners(), empty_list()),
            ],
        ),
    ]
}
