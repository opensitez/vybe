//! `StringBuilder` advanced: Append repeated calls, large constructions, chaining.
use super::helpers::run_csharp;

#[test]
fn string_builder_chained_appends_produce_correct_result() {
    assert_eq!(
        run_csharp(
            r#"var sb=new System.Text.StringBuilder();
sb.Append("a").Append("b").Append("c");
Console.WriteLine(sb.ToString());"#
        ),
        &["abc"]
    );
}

#[test]
fn string_builder_append_line_adds_newline_separator() {
    assert_eq!(
        run_csharp(
            r#"var sb=new System.Text.StringBuilder();
sb.AppendLine("line1").AppendLine("line2");
Console.WriteLine(sb.ToString().Trim().Replace("\r\n","\n"));"#
        ),
        &["line1\nline2"]
    );
}

#[test]
fn string_builder_capacity_grows_automatically() {
    assert_eq!(
        run_csharp(
            r#"var sb=new System.Text.StringBuilder(4);
for(int i=0;i<100;i++) sb.Append('x');
Console.WriteLine(sb.Length);"#
        ),
        &["100"]
    );
}

#[test]
fn string_builder_index_access_reads_character() {
    assert_eq!(
        run_csharp(
            r#"var sb=new System.Text.StringBuilder("hello");
Console.WriteLine(sb[1]);"#
        ),
        &["e"]
    );
}

#[test]
fn string_builder_index_write_mutates_character() {
    assert_eq!(
        run_csharp(
            r#"var sb=new System.Text.StringBuilder("hello");
sb[0]='H';
Console.WriteLine(sb.ToString());"#
        ),
        &["Hello"]
    );
}

#[test]
fn string_builder_to_string_substring_overload() {
    assert_eq!(
        run_csharp(
            r#"var sb=new System.Text.StringBuilder("hello world");
Console.WriteLine(sb.ToString(6,5));"#
        ),
        &["world"]
    );
}
