//! Type testing and casting: `is` with declaration, `as`, direct cast, pattern `is`.
use super::helpers::run_csharp;

#[test]
fn is_declaration_binds_matched_variable() {
    assert_eq!(
        run_csharp(
            r#"object o=42;
if(o is int n) Console.WriteLine(n*2);"#
        ),
        &["84"]
    );
}

#[test]
fn is_with_not_pattern_negates_test() {
    assert_eq!(
        run_csharp(
            r#"object o="hello";
Console.WriteLine(o is not int);"#
        ),
        &["True"]
    );
}

#[test]
fn as_returns_typed_reference_for_compatible_type() {
    assert_eq!(
        run_csharp(
            r#"object o="world";
string s=o as string;
Console.WriteLine(s!=null); Console.WriteLine(s);"#
        ),
        &["True", "world"]
    );
}

#[test]
fn as_returns_null_for_incompatible_reference_type() {
    assert_eq!(
        run_csharp(
            r#"object o=42;
string s=o as string;
Console.WriteLine(s==null);"#
        ),
        &["True"]
    );
}

#[test]
fn direct_cast_succeeds_for_compatible_type() {
    assert_eq!(
        run_csharp(
            r#"object o=3.14;
double d=(double)o;
Console.WriteLine(d);"#
        ),
        &["3.14"]
    );
}

#[test]
fn pattern_match_in_switch_dispatches_based_on_runtime_type() {
    assert_eq!(
        run_csharp(
            r#"object o=42;
string r=o switch{int n=>$"int:{n}",string s=>$"str:{s}",_=>"other"};
Console.WriteLine(r);"#
        ),
        &["int:42"]
    );
}

#[test]
fn is_null_constant_pattern_detects_null_reference() {
    assert_eq!(
        run_csharp(
            r#"object o=null;
Console.WriteLine(o is null);"#
        ),
        &["True"]
    );
}
