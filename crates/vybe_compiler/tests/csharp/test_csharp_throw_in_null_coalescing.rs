//! `throw` expressions on the right of `??` abort before a default is supplied.
use super::helpers::run_csharp;

#[test]
fn null_coalescing_throw_expression_runs_when_left_is_null() {
    assert_eq!(
        run_csharp(
            r#"
string? missing = null;
try {
    string value = missing ?? throw new System.Exception("required");
    Console.WriteLine(value);
} catch (System.Exception e) {
    Console.WriteLine(e.Message);
}
"#
        ),
        &["required"]
    );
}

#[test]
fn null_coalescing_throw_expression_skipped_when_left_has_value() {
    assert_eq!(
        run_csharp(
            r#"
string? present = "ok";
string value = present ?? throw new System.Exception("fail");
Console.WriteLine(value);
"#
        ),
        &["ok"]
    );
}

#[test]
fn chained_null_coalescing_throw_only_evaluates_when_all_prior_operands_null() {
    assert_eq!(
        run_csharp(
            r#"
string? a = null;
string? b = null;
try {
    string value = a ?? b ?? throw new System.Exception("both-null");
    Console.WriteLine(value);
} catch (System.Exception) {
    Console.WriteLine("caught");
}
"#
        ),
        &["caught"]
    );
}
