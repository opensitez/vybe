//! LINQ `GroupBy`, `Join`, `GroupJoin`, and `SelectMany` semantics.
use super::helpers::run_csharp;

#[test]
fn group_by_clusters_elements_by_key_and_counts_each_group() {
    assert_eq!(
        run_csharp(
            r#"
var words = new[] { "apple", "ant", "banana", "bear", "avocado" };
var groups = words
    .GroupBy(w => w[0])
    .OrderBy(g => g.Key)
    .Select(g => $"{g.Key}:{g.Count()}");
foreach (var s in groups) Console.WriteLine(s);
"#
        ),
        &["a:3", "b:2"]
    );
}

#[test]
fn join_links_two_sequences_on_matching_keys() {
    assert_eq!(
        run_csharp(
            r#"
var ids  = new[] { 1, 2, 3 };
var names = new[] { (Id:1, Name:"one"), (Id:2, Name:"two") };
var joined = ids.Join(names, id => id, n => n.Id, (id, n) => n.Name);
foreach (var s in joined) Console.WriteLine(s);
"#
        ),
        &["one", "two"]
    );
}

#[test]
fn select_many_flattens_nested_sequences() {
    assert_eq!(
        run_csharp(
            r#"
var nested = new[] { new[]{1,2}, new[]{3,4} };
var flat = nested.SelectMany(x => x);
int sum = 0;
foreach (var n in flat) sum += n;
Console.WriteLine(sum);
"#
        ),
        &["10"]
    );
}

#[test]
fn group_by_with_element_selector_transforms_group_members() {
    assert_eq!(
        run_csharp(
            r#"
var nums = new[] { 1, 2, 3, 4 };
var groups = nums.GroupBy(n => n % 2 == 0 ? "even" : "odd",
                          n => n * 10);
int evenSum = 0;
foreach (var g in groups)
    if (g.Key == "even") foreach (var v in g) evenSum += v;
Console.WriteLine(evenSum);
"#
        ),
        &["60"]
    );
}

#[test]
fn order_by_then_by_applies_secondary_sort_on_equal_primary_keys() {
    assert_eq!(
        run_csharp(
            r#"
var items = new[] { (Name:"b",Age:2),(Name:"a",Age:3),(Name:"a",Age:1) };
var sorted = items.OrderBy(x => x.Name).ThenBy(x => x.Age);
foreach (var x in sorted) Console.WriteLine($"{x.Name}{x.Age}");
"#
        ),
        &["a1", "a3", "b2"]
    );
}
