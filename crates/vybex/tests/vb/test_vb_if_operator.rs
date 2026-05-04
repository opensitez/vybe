use super::helpers::run_vb;

fn assert_vb_output_owned(src: String, expected: Vec<String>) {
    let out = run_vb(&src);
    assert_eq!(out, expected);
}

fn assert_if_ternary(condition: bool, when_true: i32, when_false: i32) {
    let condition_literal = if condition { "True" } else { "False" };
    let src = format!(r#"
Module M
    Sub Main()
        Console.WriteLine(If({condition}, {when_true}, {when_false}))
    End Sub
End Module
"#, condition = condition_literal, when_true = when_true, when_false = when_false);
    let expected = if condition { when_true } else { when_false };
    assert_vb_output_owned(src, vec![expected.to_string()]);
}

fn assert_if_coalesce(primary: Option<&str>, fallback: &str) {
    let primary_expr = match primary {
        Some(value) => format!("\"{}\"", value),
        None => "Nothing".to_string(),
    };
    let src = format!(r#"
Module M
    Sub Main()
        Dim value As String = {primary}
        Console.WriteLine(If(value, "{fallback}"))
    End Sub
End Module
"#, primary = primary_expr, fallback = fallback);
    let expected = primary.unwrap_or(fallback).to_string();
    assert_vb_output_owned(src, vec![expected]);
}

macro_rules! if_operator_cases {
    ($($name:ident => ternary($condition:expr, $when_true:expr, $when_false:expr)),* ; $($name2:ident => coalesce($primary:expr, $fallback:expr)),* $(,)?) => {
        $(#[test] fn $name() { assert_if_ternary($condition, $when_true, $when_false); })*
        $(#[test] fn $name2() { assert_if_coalesce($primary, $fallback); })*
    };
}

if_operator_cases! {
    if_operator_001 => ternary(true, 1, 9),
    if_operator_002 => ternary(false, 1, 9),
    if_operator_003 => ternary(true, 2, 8),
    if_operator_004 => ternary(false, 2, 8),
    if_operator_005 => ternary(true, 5, 15),
    if_operator_006 => ternary(false, 5, 15),
    if_operator_007 => ternary(true, 10, 20),
    if_operator_008 => ternary(false, 10, 20),
    if_operator_009 => ternary(true, 42, 24),
    if_operator_010 => ternary(false, 42, 24)
    ;
    if_operator_011 => coalesce(Some("alpha"), "fallback"),
    if_operator_012 => coalesce(None, "fallback"),
    if_operator_013 => coalesce(Some("vb"), "lang"),
    if_operator_014 => coalesce(None, "lang"),
    if_operator_015 => coalesce(Some("left"), "right"),
    if_operator_016 => coalesce(None, "right"),
    if_operator_017 => coalesce(Some("hello"), "world"),
    if_operator_018 => coalesce(None, "world"),
    if_operator_019 => coalesce(Some("primary"), "secondary"),
    if_operator_020 => coalesce(None, "secondary"),
}