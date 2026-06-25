//! LINQ projection operators: Select, SelectMany, Where, Take, Skip, TakeWhile, SkipWhile.
use super::helpers::run_csharp;

#[test]
fn select_transforms_each_element() {
    assert_eq!(
        run_csharp(
            r#"var result = new[]{1,2,3}.Select(x => x*x);
foreach(var n in result) Console.WriteLine(n);"#
        ),
        &["1", "4", "9"]
    );
}

#[test]
fn where_filters_to_matching_elements() {
    assert_eq!(
        run_csharp(
            r#"var result = new[]{1,2,3,4,5}.Where(x => x%2==0);
foreach(var n in result) Console.WriteLine(n);"#
        ),
        &["2", "4"]
    );
}

#[test]
fn take_returns_first_n_elements() {
    assert_eq!(
        run_csharp(
            r#"var result = new[]{10,20,30,40}.Take(2);
foreach(var n in result) Console.WriteLine(n);"#
        ),
        &["10", "20"]
    );
}

#[test]
fn skip_omits_first_n_elements() {
    assert_eq!(
        run_csharp(
            r#"var result = new[]{10,20,30,40}.Skip(2);
foreach(var n in result) Console.WriteLine(n);"#
        ),
        &["30", "40"]
    );
}

#[test]
fn take_while_stops_at_first_failing_predicate() {
    assert_eq!(
        run_csharp(
            r#"var result = new[]{1,3,5,4,7}.TakeWhile(x => x%2!=0);
foreach(var n in result) Console.WriteLine(n);"#
        ),
        &["1", "3", "5"]
    );
}

#[test]
fn skip_while_skips_until_predicate_fails() {
    assert_eq!(
        run_csharp(
            r#"var result = new[]{1,2,3,4,5}.SkipWhile(x => x<3);
foreach(var n in result) Console.WriteLine(n);"#
        ),
        &["3", "4", "5"]
    );
}

#[test]
fn order_by_descending_reverses_default_sort_order() {
    assert_eq!(
        run_csharp(
            r#"var result = new[]{3,1,4,1,5}.OrderByDescending(x => x).Distinct();
foreach(var n in result) Console.WriteLine(n);"#
        ),
        &["5", "4", "3", "1"]
    );
}

#[test]
fn select_with_index_provides_position() {
    assert_eq!(
        run_csharp(
            r#"var result = new[]{"a","b","c"}.Select((x,i) => $"{i}:{x}");
foreach(var s in result) Console.WriteLine(s);"#
        ),
        &["0:a", "1:b", "2:c"]
    );
}

#[test]
fn to_list_materializes_query_to_mutable_list() {
    assert_eq!(
        run_csharp(
            r#"var list = new[]{1,2,3}.Select(x => x*2).ToList();
Console.WriteLine(list.GetType().Name);"#
        ),
        &["List`1"]
    );
}

#[test]
fn to_dictionary_builds_map_from_sequence() {
    assert_eq!(
        run_csharp(
            r#"var dict = new[]{"a","bb","ccc"}.ToDictionary(s => s, s => s.Length);
Console.WriteLine(dict["bb"]);"#
        ),
        &["2"]
    );
}
