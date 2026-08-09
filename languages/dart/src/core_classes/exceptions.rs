//! `dart:core`'s exception and error types, as classes.
//!
//! Pascal's `synthesize_exception_classes` is the same move
//! ([[project_pascal_exception_classes_are_real_classes]]): an exception is a
//! CLASS, and declaring it as one is what gives `catch` a real identity to test.
//!
//! # What the marker got wrong, measured 2026-08-09
//!
//! The six `__dart_*` builtins this replaces built an anonymous `struct.new 0`
//! and stamped strings onto it. Because `emit_exception_new_finalize`
//! CANONICALIZES the name for cross-language catch (`FormatException` →
//! `ValueError`, `primitives/errors.rs:360`), while `emit_instanceof_chain`
//! stamped the Dart names into `__types`, one object carried two different
//! answers to "what am I". Four observable consequences, none of which any test
//! covers:
//!
//! | expression | marker | Dart |
//! |---|---|---|
//! | `FormatException("bad").runtimeType` | `[object FormatException]` | `FormatException` |
//! | `FormatException("bad").toString()` | `bad` | `FormatException: bad` |
//! | `Exception("boom").toString()` | `boom` | `Exception: boom` |
//! | `"$e"` on a caught `FormatException` | `FormatException.FormatException` | `FormatException: x` |
//!
//! A real class answers all four from one place: the rtt IS the runtime type,
//! and `toString` is an ordinary member.
//!
//! # Why catch still works
//!
//! `statements.rs:1563` tests a catch arm with `reflection::emit_is_instance_of`,
//! which unions the rtt (`ref.test`) with the `__types` ancestry chain
//! ([[project_identity_rtt_union_types]]). `compile_class` stamps that chain
//! from the MRO (`classes.rs:3931`), so a declared [`parents`] entry is what
//! makes `on Error` catch a `RangeError`. Dart is `throwable_is_root = true`
//! (`profile:59`), so catch types are NOT canonicalized — they match the real
//! class names.
//!
//! # Why every class restates `message`, its constructor and `toString`
//!
//! [[project_derived_exception_loses_base_ctor]] records that a derived
//! exception class loses its base constructor. Each class here is therefore
//! SELF-CONTAINED: it declares its own field, its own constructor and its own
//! `toString`. `parents` carries the type chain and nothing else, so the
//! base-ctor path is never on the critical path for construction.
//!
//! `message` is a FIELD, so it never enters the flat, class-less
//! `defined_class_methods` set (`calls.rs:5995`) that diverts untyped
//! receivers — the hazard that keeps `Duration.compareTo` undeclared. The only
//! method these classes declare is `toString`, which `user_typed_receiver_shadow`
//! exempts by name.

use super::builders::*;
use vybe_ast::{ExprKind, Expression, InterpolPart, Statement};

/// The field every one of these carries. Dart spells it `message` and it is
/// read directly (`e.message`), so the storage name is the source name.
const MESSAGE: &str = "message";

/// Each type, its parents, and the prefix its `toString` writes.
///
/// The prefixes are Dart's, not a scheme: `StateError` renders `Bad state:` and
/// `ArgumentError` renders `Invalid argument(s):`, neither of which is the
/// class name. Anything derived would be wrong for those two.
///
/// `Error` and `Exception` are declared even though no builtin constructed
/// them, because they are the ANCESTORS the catch arms name — `on Error`
/// appears 12 times and `on Exception` 4 times across the dart suite, and an
/// ancestor that is not a class is not in anybody's `__types` chain.
const EXCEPTION_CLASSES: &[(&str, &[&str], &str)] = &[
    ("Exception", &[], "Exception"),
    ("FormatException", &["Exception"], "FormatException"),
    ("Error", &[], "Error"),
    ("StateError", &["Error"], "Bad state"),
    ("ArgumentError", &["Error"], "Invalid argument(s)"),
    // Dart: `class RangeError extends ArgumentError`. Declaring the real
    // parent is what lets `on ArgumentError` catch a `RangeError`.
    ("RangeError", &["ArgumentError"], "RangeError"),
    ("UnimplementedError", &["Error"], "UnimplementedError"),
];

/// The builder `CORE_CLASSES` names, one per row. `fn() -> Statement` cannot
/// capture, so each is a one-liner that names its own row; `build` panics if a
/// name is not in the table, which surfaces on the first dart program compiled.
pub(super) fn exception() -> Statement {
    build("Exception")
}
pub(super) fn format_exception() -> Statement {
    build("FormatException")
}
pub(super) fn error() -> Statement {
    build("Error")
}
pub(super) fn state_error() -> Statement {
    build("StateError")
}
pub(super) fn argument_error() -> Statement {
    build("ArgumentError")
}
pub(super) fn range_error() -> Statement {
    build("RangeError")
}
pub(super) fn unimplemented_error() -> Statement {
    build("UnimplementedError")
}

/// The parents of `name`, or empty for a type this module does not declare.
///
/// The ancestor closure in `mod.rs` reads this: a class is only worth compiling
/// when the program names it, and a parent it never names is still needed for
/// the `__types` chain that answers `on <Parent>`.
pub(super) fn parents_of(name: &str) -> &'static [&'static str] {
    EXCEPTION_CLASSES
        .iter()
        .find(|(n, _, _)| *n == name)
        .map(|(_, parents, _)| *parents)
        .unwrap_or(&[])
}

fn build(name: &str) -> Statement {
    let (name, parents, prefix) = *EXCEPTION_CLASSES
        .iter()
        .find(|(n, _, _)| *n == name)
        .unwrap_or_else(|| panic!("no EXCEPTION_CLASSES row for {name}"));
    class_extending(
        name,
        parents,
        vec![
            field(MESSAGE, "String", str_lit("")),
            // `X([Object message = ''])` — Dart's optional single argument.
            // The interpolation is the shared `to_string` slot, so a non-String
            // argument (`RangeError(42)`) renders the way every other
            // stringification in the language does.
            constructor(
                vec![param("message", Some("Object"), Some(str_lit("")))],
                vec![set_this(MESSAGE, stringify(ident("message")))],
            ),
            method(
                "toString",
                vec![],
                Some("String"),
                vec![to_string_body(prefix)],
            ),
        ],
    )
}

/// `return message.isEmpty ? '<prefix>' : '<prefix>: $message';`
///
/// Dart drops the separator when there is no message — `Exception().toString()`
/// is `Exception`, not `Exception: `. The empty test stands in for Dart's
/// `message == null`, since the constructor defaults to `''` rather than null.
fn to_string_body(prefix: &str) -> Statement {
    let bare = str_lit(prefix);
    let with_message = Expression::with_span(
        ExprKind::Interpolation(vec![
            InterpolPart::Text(format!("{prefix}: ")),
            InterpolPart::Expr(this_field(MESSAGE)),
        ]),
        span(),
    );
    ret(ternary(
        call_member(this_field(MESSAGE), "isEmpty", vec![]),
        bare,
        with_message,
    ))
}
