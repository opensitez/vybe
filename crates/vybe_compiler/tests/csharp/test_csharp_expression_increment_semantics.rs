//! Prefix and postfix `++`/`--` return different values inside larger expressions
//! and only mutate storage after the subexpression value is produced.
use super::helpers::run_csharp;

#[test]
fn postfix_increment_returns_value_before_bump() {
    assert_eq!(
        run_csharp(
            r#"
int n = 4;
int read = n++;
Console.WriteLine(read);
Console.WriteLine(n);
"#
        ),
        &["4", "5"]
    );
}

#[test]
fn prefix_increment_returns_value_after_bump() {
    assert_eq!(
        run_csharp(
            r#"
int n = 4;
int read = ++n;
Console.WriteLine(read);
Console.WriteLine(n);
"#
        ),
        &["5", "5"]
    );
}

#[test]
fn postfix_decrement_in_expression_uses_original_value() {
    assert_eq!(
        run_csharp(
            r#"
int n = 3;
int total = n-- + n;
Console.WriteLine(total);
Console.WriteLine(n);
"#
        ),
        &["5", "2"]
    );
}

#[test]
fn prefix_decrement_in_expression_uses_updated_value() {
    assert_eq!(
        run_csharp(
            r#"
int n = 3;
int total = --n + n;
Console.WriteLine(total);
Console.WriteLine(n);
"#
        ),
        &["4", "2"]
    );
}

#[test]
fn increment_used_as_array_index_applies_after_index_is_read() {
    assert_eq!(
        run_csharp(
            r#"
var data = new[] { 10, 20, 30 };
int i = 0;
Console.WriteLine(data[i++]);
Console.WriteLine(i);
"#
        ),
        &["10", "1"]
    );
}
