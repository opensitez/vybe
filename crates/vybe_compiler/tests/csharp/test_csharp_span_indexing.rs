//! `Span<T>` and range/index-from-end slicing of arrays and strings.
use super::helpers::run_csharp;

#[test]
fn span_from_array_slice_reads_correct_elements() {
    assert_eq!(
        run_csharp(
            r#"
int[] data = { 10, 20, 30, 40, 50 };
var span = new System.Span<int>(data, 1, 3);
Console.WriteLine(span[0]);
Console.WriteLine(span[2]);
"#
        ),
        &["20", "40"]
    );
}

#[test]
fn span_length_matches_requested_slice_count() {
    assert_eq!(
        run_csharp(
            r#"
int[] data = { 1, 2, 3, 4 };
var span = data.AsSpan(1, 2);
Console.WriteLine(span.Length);
"#
        ),
        &["2"]
    );
}

#[test]
fn span_write_mutates_backing_array() {
    assert_eq!(
        run_csharp(
            r#"
int[] data = { 1, 2, 3 };
var span = data.AsSpan();
span[1] = 99;
Console.WriteLine(data[1]);
"#
        ),
        &["99"]
    );
}

#[test]
fn memory_slice_reads_correct_element_via_span() {
    assert_eq!(
        run_csharp(
            r#"
var memory = new System.Memory<int>(new int[] { 5, 6, 7 });
Console.WriteLine(memory.Span[2]);
"#
        ),
        &["7"]
    );
}

#[test]
fn readonly_span_from_string_has_correct_length() {
    assert_eq!(
        run_csharp(
            r#"
System.ReadOnlySpan<char> span = "hello".AsSpan();
Console.WriteLine(span.Length);
"#
        ),
        &["5"]
    );
}
