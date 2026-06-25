//! `Array.BinarySearch` requires a sorted array and returns match index or bitwise complement.
use super::helpers::run_csharp;

#[test]
fn array_binary_search_returns_index_of_existing_sorted_element() {
    assert_eq!(
        run_csharp(
            r#"
int[] sorted = { 2, 4, 6, 8 };
Console.WriteLine(System.Array.BinarySearch(sorted, 6));
"#
        ),
        &["2"]
    );
}

#[test]
fn array_binary_search_returns_complement_for_missing_insertion_point() {
    assert_eq!(
        run_csharp(
            r#"
int[] sorted = { 2, 4, 8 };
int index = System.Array.BinarySearch(sorted, 5);
Console.WriteLine(index < 0);
Console.WriteLine(~index);
"#
        ),
        &["True", "2"]
    );
}
