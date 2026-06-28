//! `ReadOnlySpan<char>`, `Span<char>`, and span-based string operations.
use super::helpers::run_csharp;

#[test]
fn readonly_span_char_from_string_slice_reads_substring() {
    assert_eq!(
        run_csharp(
            r#"string s="hello world";
System.ReadOnlySpan<char> span=s.AsSpan(6,5);
Console.WriteLine(span.ToString());"#
        ),
        &["world"]
    );
}

#[test]
fn span_contains_finds_character_in_range() {
    assert_eq!(
        run_csharp(
            r#"System.ReadOnlySpan<char> span="hello".AsSpan();
Console.WriteLine(span.Contains('e'));"#
        ),
        &["True"]
    );
}

#[test]
fn span_of_int_slice_modifies_original_array() {
    assert_eq!(
        run_csharp(
            r#"int[] arr={1,2,3,4,5};
System.Span<int> s=arr.AsSpan(1,3);
s[0]=99;
Console.WriteLine(arr[1]);"#
        ),
        &["99"]
    );
}

#[test]
fn span_copy_to_writes_into_destination() {
    assert_eq!(
        run_csharp(
            r#"int[] src={1,2,3};
int[] dst=new int[3];
src.AsSpan().CopyTo(dst);
Console.WriteLine(dst[2]);"#
        ),
        &["3"]
    );
}

#[test]
fn memory_span_property_accesses_underlying_data() {
    assert_eq!(
        run_csharp(
            r#"System.Memory<int> m=new int[]{7,8,9};
Console.WriteLine(m.Span[1]);"#
        ),
        &["8"]
    );
}

#[test]
fn readonly_span_index_from_end_works() {
    assert_eq!(
        run_csharp(
            r#"System.ReadOnlySpan<char> s="hello".AsSpan();
Console.WriteLine(s[^1]);"#
        ),
        &["o"]
    );
}
