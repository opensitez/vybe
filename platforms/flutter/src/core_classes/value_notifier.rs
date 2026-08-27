//! `package:flutter/foundation.dart`'s `ValueNotifier`, as a class.

use super::builders::*;
use super::notifier;
use vybe_ast::Statement;

/// The current value.
const VALUE: &str = "_vybeValue";

/// `class ValueNotifier<T> extends ChangeNotifier { … }`
///
/// Carries the notifier members directly rather than through `extends`: both
/// classes are synthesized into the same module, and an inheritance link
/// between two spliced declarations would make the pair order-dependent while
/// publishing exactly the same surface.
pub(super) fn value_notifier() -> Statement {
    let mut members = vec![
        field(VALUE, "dynamic", null_lit()),
        // ⛔The parameter must NOT be called `value`. Dart declares
        // `ValueNotifier(this.value)`, but these bodies are synthesized and the
        // class publishes a PROPERTY named `value`, so a bare `value` in the
        // constructor resolved to the property GETTER — reading the field it
        // was about to initialise — and every notifier was born holding null.
        constructor(vec![param("initialValue", None)], {
            let mut body = notifier::init_state();
            body.push(assign(this_field(VALUE), ident("initialValue")));
            body
        }),
        // Assigning an EQUAL value must NOT notify — Flutter compares with `==`
        // and returns early. `value_notifier_update` asserts exactly that by
        // writing 1 twice and expecting one notification.
        property(
            "value",
            "dynamic",
            // ACCESSOR bodies — receiver in the identifier form throughout.
            vec![ret(accessor_field(VALUE))],
            "newValue",
            vec![
                // ⛔`__dart_eq`, not a bare `Binary{Eq}`. Dart's `==` dispatches
                // to a user `operator ==`, and the slot ladder that does so is
                // applied by the walker — which never sees a spliced body. With
                // the raw node the comparison fell back to identity, so
                // `vn.value = Eq(1)` notified even though the value compared
                // equal.
                if_stmt(
                    call_value(
                        ident("__dart_eq"),
                        vec![accessor_field(VALUE), ident("newValue")],
                    ),
                    vec![ret_void()],
                ),
                assign(accessor_field(VALUE), ident("newValue")),
                expr_stmt(call_member(accessor_this(), "notifyListeners", vec![])),
            ],
        ),
    ];
    members.extend(notifier::members());
    class_decl("ValueNotifier", members)
}
