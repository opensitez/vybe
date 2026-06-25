//! Boolean `&&`, `||`, and `!` must not evaluate the right-hand operand when the
//! result is already determined — observable via side-effect counters.
use super::helpers::run_csharp;

#[test]
fn logical_and_skips_right_operand_when_left_is_false() {
    assert_eq!(
        run_csharp(
            r#"
int calls = 0;
bool Right() { calls++; return true; }
bool result = false && Right();
Console.WriteLine(result ? "T" : "F");
Console.WriteLine(calls);
"#
        ),
        &["F", "0"]
    );
}

#[test]
fn logical_and_evaluates_right_operand_when_left_is_true() {
    assert_eq!(
        run_csharp(
            r#"
int calls = 0;
bool Right() { calls++; return false; }
bool result = true && Right();
Console.WriteLine(result ? "T" : "F");
Console.WriteLine(calls);
"#
        ),
        &["F", "1"]
    );
}

#[test]
fn logical_or_skips_right_operand_when_left_is_true() {
    assert_eq!(
        run_csharp(
            r#"
int calls = 0;
bool Right() { calls++; return true; }
bool result = true || Right();
Console.WriteLine(result ? "T" : "F");
Console.WriteLine(calls);
"#
        ),
        &["T", "0"]
    );
}

#[test]
fn logical_or_evaluates_right_operand_when_left_is_false() {
    assert_eq!(
        run_csharp(
            r#"
int calls = 0;
bool Right() { calls++; return true; }
bool result = false || Right();
Console.WriteLine(result ? "T" : "F");
Console.WriteLine(calls);
"#
        ),
        &["T", "1"]
    );
}

#[test]
fn and_short_circuits_before_or_evaluates_fallback_operand() {
    assert_eq!(
        run_csharp(
            r#"
int trace = 0;
bool A() { trace++; return false; }
bool B() { trace++; return true; }
bool C() { trace++; return true; }
bool value = A() && B() || C();
Console.WriteLine(value ? "T" : "F");
Console.WriteLine(trace);
"#
        ),
        &["T", "2"]
    );
}
