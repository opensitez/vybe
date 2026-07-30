//! `Type.GetType(name)` — runtime type lookup by name.
//!
//! The .NET counterpart of Java's `Class.forName`, but with the opposite
//! miss behaviour: `Type.GetType` returns **null** when the name does not
//! resolve, and only throws when asked to (`throwOnError: true`). So this
//! exercises the resolve-or-null shape rather than resolve-or-throw, on the
//! same shared dynamic-symbol path.
//!
//! Note `x.GetType()` (the instance method — "what type is this value") is a
//! different operation and is already handled; these cover the static
//! name → type lookup.

use crate::helpers::assert_csharp;

#[test]
fn get_type_resolves_a_user_declared_class() {
    assert_csharp(
        r#"
class Widget {}
class Program {
    static void Main() {
        System.Console.WriteLine(System.Type.GetType("Widget") != null);
    }
}
"#,
        &["True"],
    );
}

#[test]
fn get_type_returns_null_for_an_unknown_name() {
    assert_csharp(
        r#"
class Program {
    static void Main() {
        System.Console.WriteLine(System.Type.GetType("NoSuchTypeAnywhere") == null);
    }
}
"#,
        &["True"],
    );
}

#[test]
fn get_type_name_round_trips() {
    assert_csharp(
        r#"
class Gadget {}
class Program {
    static void Main() {
        System.Console.WriteLine(System.Type.GetType("Gadget").Name);
    }
}
"#,
        &["Gadget"],
    );
}
