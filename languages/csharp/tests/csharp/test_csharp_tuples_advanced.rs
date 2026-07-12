//! Value tuple operations, named elements, deconstruction, LINQ projections.
use super::helpers::run_csharp;

#[test]
fn named_tuple_elements_accessed_by_name() {
    assert_eq!(
        run_csharp(
            r#"var p = (X: 3, Y: 4);
Console.WriteLine(p.X); Console.WriteLine(p.Y);"#
        ),
        &["3", "4"]
    );
}

#[test]
fn unnamed_tuple_elements_accessed_by_item1_item2() {
    assert_eq!(
        run_csharp(
            r#"var t = (1, "hello");
Console.WriteLine(t.Item1); Console.WriteLine(t.Item2);"#
        ),
        &["1", "hello"]
    );
}

#[test]
fn tuple_deconstruction_assigns_to_separate_locals() {
    assert_eq!(
        run_csharp(
            r#"var (a, b, c) = (10, 20, 30);
Console.WriteLine(a+b+c);"#
        ),
        &["60"]
    );
}

#[test]
fn tuple_returned_from_method_and_destructured_at_call_site() {
    assert_eq!(
        run_csharp(
            r#"(int Min, int Max) Bounds(int[] arr) =>
    (arr.Min(), arr.Max());
var (lo, hi) = Bounds(new[]{5,1,9,3});
Console.WriteLine(lo); Console.WriteLine(hi);"#
        ),
        &["1", "9"]
    );
}

#[test]
fn tuple_equality_compares_element_wise() {
    assert_eq!(
        run_csharp(
            r#"var a = (1, "x"); var b = (1, "x");
Console.WriteLine(a == b);"#
        ),
        &["True"]
    );
}

#[test]
fn discard_in_deconstruction_ignores_unwanted_element() {
    assert_eq!(
        run_csharp(
            r#"var (first, _, third) = (1, 2, 3);
Console.WriteLine(first); Console.WriteLine(third);"#
        ),
        &["1", "3"]
    );
}

#[test]
fn tuple_in_linq_select_creates_anonymous_projection() {
    assert_eq!(
        run_csharp(
            r#"var items = new[]{"apple","kiwi","pear"};
var proj = items.Select(s => (Name: s, Len: s.Length));
foreach(var x in proj) Console.WriteLine(x.Len);"#
        ),
        &["5", "4", "4"]
    );
}

#[test]
fn tuple_with_eight_elements_uses_rest_field() {
    assert_eq!(
        run_csharp(
            r#"var t = (1,2,3,4,5,6,7,8);
Console.WriteLine(t.Item8);"#
        ),
        &["8"]
    );
}
