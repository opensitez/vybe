//! Advanced string operations: `String.Join` with IEnumerable, Format alignment, Concat spans.
use super::helpers::run_csharp;

#[test]
fn string_join_with_ienumerable_source() {
    assert_eq!(
        run_csharp(r#"var nums=Enumerable.Range(1,5);
Console.WriteLine(string.Join("-",nums));"#),
        &["1-2-3-4-5"]
    );
}

#[test]
fn string_concat_with_objects_uses_tostring() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.Concat("val=",42," ok=",true));"#),
        &["val=42 ok=True"]
    );
}

#[test]
fn string_format_right_align_pad_with_width() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.Format("{0,10}","hello"));"#),
        &["     hello"]
    );
}

#[test]
fn string_format_left_align_with_negative_width() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.Format("{0,-10}|","hello"));"#),
        &["hello     |"]
    );
}

#[test]
fn string_contains_with_string_comparison_case_insensitive() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine("Hello World".Contains("world",System.StringComparison.OrdinalIgnoreCase));"#),
        &["True"]
    );
}

#[test]
fn string_replace_specific_occurrence_via_stringbuilder() {
    assert_eq!(
        run_csharp(r#"string s="aababc";
var sb=new System.Text.StringBuilder(s);
int idx=s.IndexOf("ab",1);
sb.Remove(idx,2).Insert(idx,"XX");
Console.WriteLine(sb.ToString());"#),
        &["aaXXbc"]
    );
}
