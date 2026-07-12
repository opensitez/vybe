//! Method groups convert to delegate types with matching signatures.
use super::helpers::run_csharp;

#[test]
fn method_group_converts_to_func_without_explicit_lambda_wrapper() {
    assert_eq!(
        run_csharp(
            r#"
static int Double(int n) => n * 2;
System.Func<int, int> fn = Double;
Console.WriteLine(fn(6));
"#
        ),
        &["12"]
    );
}

#[test]
fn method_group_converts_to_action_for_void_method() {
    assert_eq!(
        run_csharp(
            r#"
int total = 0;
void Bump() { total++; }
System.Action bump = Bump;
bump();
Console.WriteLine(total);
"#
        ),
        &["1"]
    );
}
