//! `dart:core` classes the runtime provides, synthesized as ordinary AST.
//!
//! A builtin class is a CLASS. It is declared here as a `StmtKind::ClassDecl`
//! and appended to the module body, so it flows through the same path a user
//! class does — `normalize_class` → `NormalClass` → `compile_class` — and
//! inherits every piece of machinery that path already provides: a reserved
//! type slot and a real rtt at `struct.new_default $T`, the runtime
//! `TypeRegistry` registration, the prototype stamp that makes member dispatch
//! RECEIVER-based, MRO, and protocol-slot binding (flexclassplan §2b's
//! "explicit registration from the frontend").
//!
//! That is the difference from the `[builtins]` entry this replaces. A builtin
//! emitted an anonymous `struct.new 0` carrying a private marker field: no
//! type, no vtable, no prototype, so `sb.toString()` had nothing to dispatch
//! to and fell through to `Object.prototype.toString` → `[object Object]`.
//! Registering the name in the namespace tree fixes RESOLUTION, not identity —
//! namespaceplan §246 records that the write-side rtt migration never happened
//! and that 184 sites stamp a `__type` STRING instead. A normalized class is
//! the one shape that does not add a 185th.
//!
//! Declaring it AS AST rather than as source text is what keeps it free: no
//! scan of the program text, no second parse. Pascal's
//! `synthesize_exception_classes` is the same move.
//!
//! Bodies are deliberately plain: `+`, interpolation, `.join`, `.length`. Each
//! lowers through the shared string/collection machinery, so the class carries
//! no Dart-private buffer emitter and the semantics are the ones every other
//! language gets.
//!
//! # Layout
//!
//! One file per class, plus [`builders`] for the AST-construction vocabulary
//! they share. A class's own rationale — which members are getters, which are
//! methods, and what each choice was measured to cost — lives beside its
//! builder, not here. This module owns only the list and what is true of every
//! entry on it.

mod builders;
mod datetime;
mod duration;
mod exceptions;
mod string_buffer;
mod uri;

use vybe_ast::Statement;

/// Every `dart:core` class the walker declares, paired with its builder. The
/// walker skips any name the program declares itself, so a user
/// `class StringBuffer` still wins.
///
/// This is also what `tree_register.rs` reads to declare the same names in the
/// namespace tree, so the tree and the AST can never disagree about which
/// classes exist.
/// Order matters for the exception rows only: a class must be declared after
/// the one it extends, so the ancestor's MRO is resolved when the child's
/// `__types` chain is stamped.
pub const CORE_CLASSES: &[(&str, fn() -> Statement)] = &[
    ("StringBuffer", string_buffer::string_buffer),
    ("Duration", duration::duration),
    // After `Duration`: `DateTime.difference` constructs one.
    ("DateTime", datetime::datetime),
    ("Exception", exceptions::exception),
    ("FormatException", exceptions::format_exception),
    ("Error", exceptions::error),
    ("StateError", exceptions::state_error),
    ("ArgumentError", exceptions::argument_error),
    ("RangeError", exceptions::range_error),
    ("UnimplementedError", exceptions::unimplemented_error),
    ("Uri", uri::uri),
];

/// True when `name` is one of the classes above.
///
/// The walker needs this during the WALK — `dart_call_or_new` must know that
/// `StringBuffer(...)` is a construction, and the class it will append does
/// not exist yet at that point. It cannot ask the namespace tree instead:
/// tree registration happens at compile time, after the walk.
pub fn is_core_class(name: &str) -> bool {
    CORE_CLASSES.iter().any(|(n, _)| *n == name)
}

/// The classes `source` actually needs, built and in declaration order.
///
/// A program that never writes `RangeError` should not compile one, so a class
/// is built only when its name appears in the source. `parse` already gates the
/// Flutter runtime the same way (`walker.rs:1061`).
///
/// **This is hygiene, not a measured speedup.** Compiling all nine
/// unconditionally versus none was timed at 0.15–0.27s either way on a warm
/// binary (2026-08-09) — indistinguishable. A slice appearing 6× slower right
/// after a `cargo build` is the 260MB binary's page cache, not the class count;
/// re-run once warm before believing any timing here.
///
/// Substring containment is deliberately CONSERVATIVE: a name inside a comment
/// or a longer identifier declares a class nobody uses, which costs time and
/// nothing else. It cannot go the other way — a name that is constructed,
/// caught or annotated is by definition present in the text.
///
/// Ancestors are closed over explicitly. `RangeError` extends `ArgumentError`,
/// and a program naming only `RangeError` has no substring that would reach its
/// parent — leaving it out would break `on ArgumentError`, since the missing
/// class is a missing link in the `__types` chain.
pub fn declarations_for(source: &str, is_user_declared: impl Fn(&str) -> bool) -> Vec<Statement> {
    let mut needed: Vec<&str> = CORE_CLASSES
        .iter()
        .map(|(name, _)| *name)
        .filter(|name| source.contains(name))
        .collect();
    // `DateTime.difference` answers a `Duration`, so a program that names only
    // the first still needs the second compiled. A class dependency is closed
    // over exactly like an ancestor is.
    if needed.contains(&"DateTime") && !needed.contains(&"Duration") {
        needed.push("Duration");
    }
    let mut i = 0;
    while i < needed.len() {
        for parent in parents_of(needed[i]) {
            if !needed.contains(parent) {
                needed.push(parent);
            }
        }
        i += 1;
    }
    let mut out: Vec<Statement> = CORE_CLASSES
        .iter()
        .filter(|(name, _)| needed.contains(name) && !is_user_declared(name))
        .map(|(_, build)| build())
        .collect();
    out
}

/// The classes `name` extends, for the ancestor closure above. Only the
/// exception tree has any; `StringBuffer` and `Duration` stand alone.
fn parents_of(name: &str) -> &'static [&'static str] {
    exceptions::parents_of(name)
}

/// Member names the walker rewrites from a property read into a zero-arg CALL
/// (`is_dart_zero_arg_getter`, `walker.rs:6455`) before dispatch ever sees them.
///
/// **A core class must declare any of these as a zero-arg METHOD, never as a
/// getter or a field.** Dart spells them `bool get isEmpty`, but by the time the
/// class is dispatched against, `sb.isEmpty` is already `sb.isEmpty()` — reading
/// a field or property there yields the VALUE and then invokes it:
/// "bool is not callable (type: true)". Measured on both `StringBuffer.isEmpty`
/// and `Duration.isNegative`.
///
/// This list is the checklist for a new core class; the walker owns the real one.
/// A name leaves it only when its ROLE is consumed on the shared member-read
/// path — not because the name reads like a property.
#[allow(dead_code)]
const ZERO_ARG_GETTER_NAMES: &[&str] = &[
    "length",
    "isEmpty",
    "isNotEmpty",
    "isEven",
    "isOdd",
    "isNegative",
    "isNaN",
    "isFinite",
    "isInfinite",
    "sign",
    "first",
    "last",
    "single",
    "singleOrNull",
    "runes",
    "codeUnits",
    "keys",
    "values",
    "entries",
    "reversed",
    "isRunning",
    "elapsed",
    "elapsedMilliseconds",
    "elapsedMicroseconds",
];
