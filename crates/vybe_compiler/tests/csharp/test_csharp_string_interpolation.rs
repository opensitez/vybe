//! String interpolation: method calls, conditionals, nesting, `$@` verbatim.
use super::helpers::run_csharp;

#[test]
fn simple_variable_interpolation_in_string() {
    assert_eq!(
        run_csharp(r#"string name="World"; Console.WriteLine($"Hello {name}!");"#),
        &["Hello World!"]
    );
}

#[test]
fn expression_evaluated_inside_interpolation() {
    assert_eq!(
        run_csharp(r#"int a=3,b=4; Console.WriteLine($"{a}+{b}={a+b}");"#),
        &["3+4=7"]
    );
}

#[test]
fn method_call_inside_interpolation() {
    assert_eq!(
        run_csharp(r#"string s="hello"; Console.WriteLine($"{s.ToUpper()}");"#),
        &["HELLO"]
    );
}

#[test]
fn ternary_operator_inside_interpolation() {
    assert_eq!(
        run_csharp(r#"int n=7; Console.WriteLine($"{(n%2==0?"even":"odd")}");"#),
        &["odd"]
    );
}

#[test]
fn alignment_specifier_pads_right_aligned() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine($"{"x",5}");"#),
        &["    x"]
    );
}

#[test]
fn format_specifier_in_interpolation_applies_number_format() {
    assert_eq!(
        run_csharp(r#"double d=1234.5; Console.WriteLine($"{d:N2}");"#),
        &["1,234.50"]
    );
}

#[test]
fn nested_braces_produce_literal_brace_in_output() {
    assert_eq!(
        run_csharp(r#"int n=5; Console.WriteLine($"{{n}}={n}");"#),
        &["{n}=5"]
    );
}

#[test]
fn verbatim_interpolated_string_preserves_backslash() {
    assert_eq!(
        run_csharp(r#"string dir="docs"; Console.WriteLine($@"C:\{dir}\file.txt");"#),
        &[r#"C:\docs\file.txt"#]
    );
}
