//! `$"..."` interpolation only — holes, escapes, and expression forms.
use super::helpers::run_csharp;

#[test]
fn interpolated_string_embeds_local_variable_text() {
    assert_eq!(
        run_csharp(r#"var name = "Ada"; Console.WriteLine($"{name}");"#),
        &["Ada"]
    );
}

#[test]
fn interpolated_string_evaluates_arithmetic_inside_hole() {
    assert_eq!(
        run_csharp(r#"int a = 6; int b = 7; Console.WriteLine($"{a + b}");"#),
        &["13"]
    );
}

#[test]
fn interpolated_string_calls_method_on_expression_in_hole() {
    assert_eq!(
        run_csharp(r#"var text = "hi"; Console.WriteLine($"{text.ToUpper()}");"#),
        &["HI"]
    );
}

#[test]
fn interpolated_string_uses_ternary_expression_inside_hole() {
    assert_eq!(
        run_csharp(r#"int n = 4; Console.WriteLine($"{(n % 2 == 0 ? "even" : "odd")}");"#),
        &["even"]
    );
}

#[test]
fn interpolated_string_escapes_braces_as_literals() {
    assert_eq!(
        run_csharp(r#"int n = 3; Console.WriteLine($"{{count}}={n}");"#),
        &["{count}=3"]
    );
}

#[test]
fn interpolated_string_concatenated_with_plain_string() {
    assert_eq!(
        run_csharp(r#"var id = 7; Console.WriteLine("id=" + $"{id}");"#),
        &["id=7"]
    );
}

#[test]
fn interpolated_string_with_format_specifier_pads_numeric_output() {
    assert_eq!(
        run_csharp(r#"int n = 7; Console.WriteLine($"{n:D3}");"#),
        &["007"]
    );
}

#[test]
fn interpolated_string_multiple_holes_preserve_middle_literal_text() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine($"a{1}b{2}c");"#),
        &["a1b2c"]
    );
}

#[test]
fn interpolated_string_with_nullable_value_prints_empty_when_null() {
    assert_eq!(
        run_csharp(r#"int? value = null; Console.WriteLine($"[{value}]");"#),
        &["[]"]
    );
}

#[test]
fn interpolated_string_in_return_expression_from_local_function() {
    assert_eq!(
        run_csharp(
            r#"
string Label(int n) { return $"n={n}"; }
Console.WriteLine(Label(5));
"#
        ),
        &["n=5"]
    );
}

#[test]
fn interpolated_string_with_verbatim_literal_segment_beside_hole() {
    assert_eq!(
        run_csharp(r#"var drive = "C"; Console.WriteLine($@"{drive}\temp");"#),
        &[r"C\temp"]
    );
}

#[test]
fn interpolated_string_nested_object_member_access_in_hole() {
    assert_eq!(
        run_csharp(
            r#"
class Pair { public int A; public int B; }
var pair = new Pair { A = 2, B = 3 };
Console.WriteLine($"{pair.A + pair.B}");
"#
        ),
        &["5"]
    );
}
