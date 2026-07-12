//! Null checks, null-conditional, null-coalescing, and null reference semantics.
use super::helpers::run_csharp;

#[test]
fn null_conditional_member_access_returns_null_when_source_null() {
    assert_eq!(
        run_csharp(
            r#"string s = null;
Console.WriteLine(s?.Length == null);"#
        ),
        &["True"]
    );
}

#[test]
fn null_conditional_chains_through_nested_properties() {
    assert_eq!(
        run_csharp(
            r#"class Node { public Node Next; public int Value; }
Node head = null;
Console.WriteLine(head?.Next?.Value ?? -1);"#
        ),
        &["-1"]
    );
}

#[test]
fn null_conditional_indexer_returns_null_on_null_collection() {
    assert_eq!(
        run_csharp(
            r#"int[] arr = null;
Console.WriteLine(arr?[0] ?? -1);"#
        ),
        &["-1"]
    );
}

#[test]
fn null_conditional_method_call_does_not_execute_when_source_null() {
    assert_eq!(
        run_csharp(
            r#"string s = null;
int count = 0;
s?.ToUpper();
Console.WriteLine(count);"#
        ),
        &["0"]
    );
}

#[test]
fn null_coalescing_chained_selects_first_non_null() {
    assert_eq!(
        run_csharp(
            r#"string a=null, b=null, c="found";
Console.WriteLine(a ?? b ?? c);"#
        ),
        &["found"]
    );
}

#[test]
fn is_null_pattern_matches_null_reference() {
    assert_eq!(
        run_csharp(
            r#"object o = null;
Console.WriteLine(o is null);"#
        ),
        &["True"]
    );
}

#[test]
fn is_not_null_pattern_matches_non_null_reference() {
    assert_eq!(
        run_csharp(
            r#"object o = "hello";
Console.WriteLine(o is not null);"#
        ),
        &["True"]
    );
}

#[test]
fn null_conditional_invoke_on_event_is_safe() {
    assert_eq!(
        run_csharp(
            r#"System.Action callback = null;
callback?.Invoke();
Console.WriteLine("safe");"#
        ),
        &["safe"]
    );
}
