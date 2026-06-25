//! Conditional expressions: ternary, null-coalescing, null-coalescing assignment.
use super::helpers::run_csharp;

#[test]
fn ternary_evaluates_true_branch() {
    assert_eq!(
        run_csharp(r#"int x=10;
Console.WriteLine(x>5?"big":"small");"#),
        &["big"]
    );
}

#[test]
fn ternary_nested_three_way_comparison() {
    assert_eq!(
        run_csharp(r#"int n=0;
Console.WriteLine(n>0?"pos":n<0?"neg":"zero");"#),
        &["zero"]
    );
}

#[test]
fn null_coalescing_returns_left_when_non_null() {
    assert_eq!(
        run_csharp(r#"string s="hello";
Console.WriteLine(s??"default");"#),
        &["hello"]
    );
}

#[test]
fn null_coalescing_returns_right_when_null() {
    assert_eq!(
        run_csharp(r#"string s=null;
Console.WriteLine(s??"default");"#),
        &["default"]
    );
}

#[test]
fn null_coalescing_assignment_sets_only_when_null() {
    assert_eq!(
        run_csharp(r#"string a=null; a??="assigned";
string b="existing"; b??="assigned";
Console.WriteLine(a); Console.WriteLine(b);"#),
        &["assigned", "existing"]
    );
}

#[test]
fn null_conditional_short_circuits_entire_chain() {
    assert_eq!(
        run_csharp(r#"string s=null;
Console.WriteLine(s?.ToUpper()??"nil");"#),
        &["nil"]
    );
}

#[test]
fn conditional_expression_in_argument_position() {
    assert_eq!(
        run_csharp(r#"int n=7;
Console.WriteLine(string.Format("{0}",n%2==0?"even":"odd"));"#),
        &["odd"]
    );
}
