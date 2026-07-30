//! `Class.forName` as a real dynamic symbol lookup.
//!
//! `Class.forName(name)` is Java's runtime type resolution: look the name up,
//! let the ClassLoader try to supply it if it is not there yet, and throw
//! `ClassNotFoundException` when it still is not. That is the same shape as
//! PHP's autoload — a symbol miss, a registered resolver, a retry — so it
//! belongs on the shared `dynamic_symbols` machinery rather than on a
//! per-language stub.
//!
//! The existing `class_for_name_simple_name` / `_package_name` tests pass
//! against a stub that returns the ARGUMENT STRING unchanged
//! (`walker.rs`: `if type_name == "Class" && method == "forName" { return
//! args[0].value.clone() }`), so they only work because a dotted JDK name
//! happens to survive string-method dispatch. These cover what the stub cannot.

use crate::helpers::run_in_main;

/// A user-declared class must be findable by name, and the result must be the
/// class itself — not the string that was passed in.
#[test]
fn for_name_resolves_a_user_declared_class() {
    let out = run_in_main(
        r#"System.out.println(Class.forName("Widget").getSimpleName());"#,
        r#"static class Widget {}"#,
    );
    assert_eq!(out, vec!["Widget"]);
}

/// `Class.forName(X)` and `X.class` must denote the same thing. In this
/// compiler a class is represented by its name, so resolving must agree with
/// the class literal rather than inventing a second representation.
#[test]
fn for_name_agrees_with_the_class_literal() {
    let out = run_in_main(
        r#"System.out.println(Class.forName("Gadget") == Gadget.class);"#,
        r#"static class Gadget {}"#,
    );
    assert_eq!(out, vec!["true"]);
}

/// An unresolvable name throws `ClassNotFoundException`. The stub cannot: it
/// hands back the string and reports success for any input at all.
#[test]
fn for_name_throws_for_an_unknown_class() {
    let out = run_in_main(
        r#"try {
               Class.forName("NoSuchClassAnywhere");
               System.out.println("no-throw");
           } catch (ClassNotFoundException e) {
               System.out.println("threw");
           }"#,
        "",
    );
    assert_eq!(out, vec!["threw"]);
}
