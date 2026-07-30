//! `Type.GetType(name)` — runtime type lookup by name, in VB.
//!
//! Same operation and same shared dynamic-symbol path as C#: both are .NET, so
//! one primitive answers both. `Type.GetType` yields the type's NAME (agreeing
//! with how VB already represents a type) and `Nothing` when the name does not
//! resolve — .NET returns null on a miss rather than throwing.
//!
//! Distinct from `value.GetType()`, the instance form asking what a value is.

use crate::helpers::run_vb;

#[test]
fn get_type_resolves_a_user_declared_class() {
    let out = run_vb(
        r#"
Class Widget
End Class

Module Program
    Sub Main()
        Console.WriteLine(Type.GetType("Widget") IsNot Nothing)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn get_type_returns_nothing_for_an_unknown_name() {
    let out = run_vb(
        r#"
Module Program
    Sub Main()
        Console.WriteLine(Type.GetType("NoSuchTypeAnywhere") Is Nothing)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}
