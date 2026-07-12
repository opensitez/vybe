//! `checked` arithmetic throws `OverflowException` on out-of-range results.
use super::helpers::run_csharp;

#[test]
fn checked_add_on_int_max_value_throws_overflow_exception() {
    assert_eq!(
        run_csharp(
            r#"
string outcome = "ok";
try {
    checked {
        int value = int.MaxValue;
        value += 1;
    }
} catch (System.OverflowException) {
    outcome = "overflow";
}
Console.WriteLine(outcome);
"#
        ),
        &["overflow"]
    );
}
