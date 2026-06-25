//! `List<T>` mutation, search, and ordering operations.
use super::helpers::run_csharp;

#[test]
fn add_appends_element_and_increases_count() {
    assert_eq!(
        run_csharp(
            r#"var list = new System.Collections.Generic.List<int>();
list.Add(1); list.Add(2);
Console.WriteLine(list.Count);"#
        ),
        &["2"]
    );
}

#[test]
fn insert_places_element_at_specified_index() {
    assert_eq!(
        run_csharp(
            r#"var list = new System.Collections.Generic.List<int>{1,3};
list.Insert(1, 2);
Console.WriteLine(list[1]);"#
        ),
        &["2"]
    );
}

#[test]
fn remove_deletes_first_matching_element() {
    assert_eq!(
        run_csharp(
            r#"var list = new System.Collections.Generic.List<int>{1,2,2,3};
list.Remove(2);
Console.WriteLine(list.Count); Console.WriteLine(list[1]);"#
        ),
        &["3", "2"]
    );
}

#[test]
fn remove_at_deletes_element_by_index() {
    assert_eq!(
        run_csharp(
            r#"var list = new System.Collections.Generic.List<string>{"a","b","c"};
list.RemoveAt(0);
Console.WriteLine(list[0]);"#
        ),
        &["b"]
    );
}

#[test]
fn contains_returns_true_for_present_element() {
    assert_eq!(
        run_csharp(
            r#"var list = new System.Collections.Generic.List<int>{10,20,30};
Console.WriteLine(list.Contains(20));"#
        ),
        &["True"]
    );
}

#[test]
fn index_of_returns_first_position_of_element() {
    assert_eq!(
        run_csharp(
            r#"var list = new System.Collections.Generic.List<int>{5,10,5};
Console.WriteLine(list.IndexOf(5));"#
        ),
        &["0"]
    );
}

#[test]
fn add_range_appends_all_elements_of_collection() {
    assert_eq!(
        run_csharp(
            r#"var list = new System.Collections.Generic.List<int>{1};
list.AddRange(new[]{2,3,4});
Console.WriteLine(list.Count);"#
        ),
        &["4"]
    );
}

#[test]
fn sort_orders_elements_ascending() {
    assert_eq!(
        run_csharp(
            r#"var list = new System.Collections.Generic.List<int>{3,1,2};
list.Sort();
Console.WriteLine(list[0]); Console.WriteLine(list[2]);"#
        ),
        &["1", "3"]
    );
}

#[test]
fn reverse_inverts_element_order() {
    assert_eq!(
        run_csharp(
            r#"var list = new System.Collections.Generic.List<int>{1,2,3};
list.Reverse();
Console.WriteLine(list[0]);"#
        ),
        &["3"]
    );
}

#[test]
fn find_returns_first_element_satisfying_predicate() {
    assert_eq!(
        run_csharp(
            r#"var list = new System.Collections.Generic.List<int>{1,4,7,8};
Console.WriteLine(list.Find(x => x > 5));"#
        ),
        &["7"]
    );
}

#[test]
fn remove_all_deletes_every_matching_element() {
    assert_eq!(
        run_csharp(
            r#"var list = new System.Collections.Generic.List<int>{1,2,3,4,5};
list.RemoveAll(x => x % 2 == 0);
Console.WriteLine(list.Count);"#
        ),
        &["3"]
    );
}

#[test]
fn to_array_converts_list_to_fixed_array() {
    assert_eq!(
        run_csharp(
            r#"var list = new System.Collections.Generic.List<int>{7,8,9};
var arr = list.ToArray();
Console.WriteLine(arr.GetType().IsArray);"#
        ),
        &["True"]
    );
}

#[test]
fn exists_returns_true_when_predicate_satisfied() {
    assert_eq!(
        run_csharp(
            r#"var list = new System.Collections.Generic.List<int>{1,2,3};
Console.WriteLine(list.Exists(x => x > 2));"#
        ),
        &["True"]
    );
}
