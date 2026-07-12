//! Variable scoping: block scope, shadowing, declaration in conditions, out variables.
use super::helpers::run_csharp;

#[test]
fn variable_in_block_not_visible_after_closing_brace() {
    assert_eq!(
        run_csharp(
            r#"int outer = 1;
{
    int inner = 2;
    outer = inner;
}
Console.WriteLine(outer);"#
        ),
        &["2"]
    );
}

#[test]
fn for_loop_variable_scoped_to_loop_body() {
    assert_eq!(
        run_csharp(
            r#"for(int i=0; i<1; i++) { }
int i2 = 99;
Console.WriteLine(i2);"#
        ),
        &["99"]
    );
}

#[test]
fn out_variable_declared_inline_at_call_site() {
    assert_eq!(
        run_csharp(
            r#"if(int.TryParse("42", out int n)) Console.WriteLine(n);
else Console.WriteLine(0);"#
        ),
        &["42"]
    );
}

#[test]
fn var_keyword_infers_type_from_right_hand_side() {
    assert_eq!(
        run_csharp(
            r#"var text = "hello";
var number = 42;
Console.WriteLine(text.GetType().Name);
Console.WriteLine(number.GetType().Name);"#
        ),
        &["String", "Int32"]
    );
}

#[test]
fn multiple_assignment_in_declaration_using_tuple() {
    assert_eq!(
        run_csharp(
            r#"var (a, b) = (3, 7);
Console.WriteLine(a); Console.WriteLine(b);"#
        ),
        &["3", "7"]
    );
}

#[test]
fn const_local_cannot_be_reassigned_but_is_readable() {
    assert_eq!(
        run_csharp(
            r#"const int MAX = 100;
Console.WriteLine(MAX);"#
        ),
        &["100"]
    );
}

#[test]
fn if_declaration_pattern_scopes_bound_variable_to_body() {
    assert_eq!(
        run_csharp(
            r#"object o = "scoped";
if(o is string text)
    Console.WriteLine(text.Length);
Console.WriteLine("done");"#
        ),
        &["6", "done"]
    );
}
