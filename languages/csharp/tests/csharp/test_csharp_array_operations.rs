//! `Array` static methods and instance operations beyond basic indexing.
use super::helpers::run_csharp;

#[test]
fn array_sort_orders_elements_ascending() {
    assert_eq!(
        run_csharp(
            r#"int[] a = {3,1,4,1,5};
System.Array.Sort(a);
Console.WriteLine(a[0]); Console.WriteLine(a[4]);"#
        ),
        &["1", "5"]
    );
}

#[test]
fn array_reverse_inverts_element_order() {
    assert_eq!(
        run_csharp(
            r#"int[] a = {1,2,3};
System.Array.Reverse(a);
Console.WriteLine(a[0]);"#
        ),
        &["3"]
    );
}

#[test]
fn array_copy_transfers_elements_to_destination() {
    assert_eq!(
        run_csharp(
            r#"int[] src = {10,20,30}; int[] dst = new int[3];
System.Array.Copy(src, dst, 3);
Console.WriteLine(dst[1]);"#
        ),
        &["20"]
    );
}

#[test]
fn array_index_of_returns_first_matching_position() {
    assert_eq!(
        run_csharp(
            r#"string[] a = {"a","b","c","b"};
Console.WriteLine(System.Array.IndexOf(a,"b"));"#
        ),
        &["1"]
    );
}

#[test]
fn array_exists_detects_matching_element() {
    assert_eq!(
        run_csharp(
            r#"int[] a = {1,3,5,7};
Console.WriteLine(System.Array.Exists(a, x => x > 4));"#
        ),
        &["True"]
    );
}

#[test]
fn array_find_returns_first_match() {
    assert_eq!(
        run_csharp(
            r#"int[] a = {1,3,5,7};
Console.WriteLine(System.Array.Find(a, x => x > 3));"#
        ),
        &["5"]
    );
}

#[test]
fn array_find_all_returns_all_matches() {
    assert_eq!(
        run_csharp(
            r#"int[] a = {1,2,3,4,5};
int[] evens = System.Array.FindAll(a, x => x%2==0);
Console.WriteLine(evens.Length);"#
        ),
        &["2"]
    );
}

#[test]
fn array_resize_grows_array_and_preserves_existing_elements() {
    assert_eq!(
        run_csharp(
            r#"int[] a = {1,2,3};
System.Array.Resize(ref a, 5);
Console.WriteLine(a.Length); Console.WriteLine(a[2]);"#
        ),
        &["5", "3"]
    );
}

#[test]
fn array_clear_fills_range_with_default_values() {
    assert_eq!(
        run_csharp(
            r#"int[] a = {1,2,3,4,5};
System.Array.Clear(a, 1, 3);
Console.WriteLine(a[0]); Console.WriteLine(a[2]);"#
        ),
        &["1", "0"]
    );
}

#[test]
fn array_fill_sets_all_elements_to_given_value() {
    assert_eq!(
        run_csharp(
            r#"int[] a = new int[4];
System.Array.Fill(a, 7);
Console.WriteLine(a[3]);"#
        ),
        &["7"]
    );
}
