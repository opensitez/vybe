//! Collection expression syntax (C# 12) and spread element.
use super::helpers::run_csharp;

#[test]
fn collection_expression_creates_list_directly() {
    assert_eq!(
        run_csharp(r#"System.Collections.Generic.List<int> list=[1,2,3];
Console.WriteLine(list.Count); Console.WriteLine(list[1]);"#),
        &["3", "2"]
    );
}

#[test]
fn collection_expression_creates_array_directly() {
    assert_eq!(
        run_csharp(r#"int[] arr=[10,20,30];
Console.WriteLine(arr.Length);"#),
        &["3"]
    );
}

#[test]
fn collection_expression_spread_merges_two_spans() {
    assert_eq!(
        run_csharp(r#"int[] a=[1,2,3];
int[] b=[4,5,6];
int[] c=[..a,..b];
Console.WriteLine(c.Length); Console.WriteLine(c[3]);"#),
        &["6", "4"]
    );
}

#[test]
fn collection_expression_empty_array_has_zero_length() {
    assert_eq!(
        run_csharp(r#"int[] empty=[];
Console.WriteLine(empty.Length);"#),
        &["0"]
    );
}

#[test]
fn span_collection_expression_works_with_stack_alloc_semantics() {
    assert_eq!(
        run_csharp(r#"System.Span<int> s=[1,2,3];
Console.WriteLine(s.Length);"#),
        &["3"]
    );
}
