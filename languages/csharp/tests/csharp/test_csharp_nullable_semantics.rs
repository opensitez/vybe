//! Nullable value types `T?`: HasValue, Value, GetValueOrDefault, conversions.
use super::helpers::run_csharp;

#[test]
fn nullable_int_has_value_true_when_assigned() {
    assert_eq!(
        run_csharp(r#"int? n = 5; Console.WriteLine(n.HasValue);"#),
        &["True"]
    );
}

#[test]
fn nullable_int_has_value_false_when_null() {
    assert_eq!(
        run_csharp(r#"int? n = null; Console.WriteLine(n.HasValue);"#),
        &["False"]
    );
}

#[test]
fn value_property_retrieves_unwrapped_value() {
    assert_eq!(
        run_csharp(r#"int? n = 42; Console.WriteLine(n.Value);"#),
        &["42"]
    );
}

#[test]
fn get_value_or_default_returns_fallback_when_null() {
    assert_eq!(
        run_csharp(r#"int? n = null; Console.WriteLine(n.GetValueOrDefault(-1));"#),
        &["-1"]
    );
}

#[test]
fn null_coalescing_operator_returns_right_side_when_null() {
    assert_eq!(
        run_csharp(r#"int? n = null; Console.WriteLine(n ?? 99);"#),
        &["99"]
    );
}

#[test]
fn null_coalescing_returns_left_side_when_non_null() {
    assert_eq!(
        run_csharp(r#"int? n = 7; Console.WriteLine(n ?? 99);"#),
        &["7"]
    );
}

#[test]
fn null_coalescing_assign_only_sets_when_currently_null() {
    assert_eq!(
        run_csharp(r#"int? n = null; n ??= 5; Console.WriteLine(n);"#),
        &["5"]
    );
}

#[test]
fn arithmetic_on_two_nullable_ints_with_values_produces_nullable_result() {
    assert_eq!(
        run_csharp(r#"int? a=3, b=4; int? c=a+b; Console.WriteLine(c);"#),
        &["7"]
    );
}

#[test]
fn arithmetic_on_nullable_where_one_is_null_yields_null() {
    assert_eq!(
        run_csharp(r#"int? a=3, b=null; Console.WriteLine((a+b).HasValue);"#),
        &["False"]
    );
}

#[test]
fn nullable_value_type_cast_to_non_nullable_succeeds_with_value() {
    assert_eq!(
        run_csharp(r#"int? n = 10; int x = (int)n; Console.WriteLine(x);"#),
        &["10"]
    );
}

#[test]
fn nullable_bool_supports_three_state_logic() {
    assert_eq!(
        run_csharp(
            r#"bool? a = true, b = null;
Console.WriteLine(a == true);
Console.WriteLine(b == null);"#
        ),
        &["True", "True"]
    );
}

#[test]
fn comparing_two_nulls_returns_true_for_equality() {
    assert_eq!(
        run_csharp(r#"int? a=null, b=null; Console.WriteLine(a==b);"#),
        &["True"]
    );
}
