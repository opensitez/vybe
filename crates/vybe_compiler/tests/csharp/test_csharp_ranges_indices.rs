//! `Index` (`^`) and `Range` (`..`) operators on arrays and strings.
use super::helpers::run_csharp;

#[test]
fn index_from_end_one_is_last_element() {
    assert_eq!(
        run_csharp(r#"int[] a={1,2,3,4,5}; Console.WriteLine(a[^1]);"#),
        &["5"]
    );
}

#[test]
fn range_slice_returns_sub_array() {
    assert_eq!(
        run_csharp(r#"int[] a={1,2,3,4,5}; var s=a[1..4];
Console.WriteLine(s.Length); Console.WriteLine(s[0]); Console.WriteLine(s[2]);"#),
        &["3", "2", "4"]
    );
}

#[test]
fn range_from_start_to_end_returns_full_array() {
    assert_eq!(
        run_csharp(r#"int[] a={1,2,3}; var s=a[..];
Console.WriteLine(s.Length);"#),
        &["3"]
    );
}

#[test]
fn range_open_start_takes_prefix() {
    assert_eq!(
        run_csharp(r#"int[] a={1,2,3,4,5}; var s=a[..3];
Console.WriteLine(s.Length); Console.WriteLine(s[2]);"#),
        &["3", "3"]
    );
}

#[test]
fn range_open_end_takes_suffix() {
    assert_eq!(
        run_csharp(r#"int[] a={1,2,3,4,5}; var s=a[3..];
Console.WriteLine(s.Length); Console.WriteLine(s[0]);"#),
        &["2", "4"]
    );
}

#[test]
fn range_on_string_returns_substring() {
    assert_eq!(
        run_csharp(r#"string s="hello world"; Console.WriteLine(s[6..]);"#),
        &["world"]
    );
}

#[test]
fn index_variable_used_in_array_access() {
    assert_eq!(
        run_csharp(r#"int[] a={10,20,30,40,50};
System.Index i=^2;
Console.WriteLine(a[i]);"#),
        &["40"]
    );
}
