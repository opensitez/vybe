//! `string.Format` composite format strings with indexed placeholders.
use super::helpers::run_csharp;

#[test]
fn string_format_single_placeholder_interpolates_argument() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.Format("Hello {0}!", "world"));"#),
        &["Hello world!"]
    );
}

#[test]
fn string_format_multiple_placeholders_map_to_positional_arguments() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.Format("{0} + {1} = {2}", 1, 2, 3));"#),
        &["1 + 2 = 3"]
    );
}

#[test]
fn string_format_placeholder_can_repeat_same_index() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.Format("{0} and {0}", "x"));"#),
        &["x and x"]
    );
}

#[test]
fn string_format_with_format_specifier_inside_placeholder() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.Format("{0:F1}", 3.14159));"#),
        &["3.1"]
    );
}

#[test]
fn string_format_with_alignment_right_pads_to_field_width() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.Format("{0,5}", 42));"#),
        &["   42"]
    );
}

#[test]
fn string_format_with_negative_alignment_left_pads_to_field_width() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.Format("{0,-5}|", "ab"));"#),
        &["ab   |"]
    );
}

#[test]
fn string_format_null_argument_renders_as_empty_string() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(string.Format("[{0}]", (object)null));"#),
        &["[]"]
    );
}
