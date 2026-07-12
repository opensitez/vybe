//! Nullable value types (`T?`): lifted operators, HasValue, GetValueOrDefault,
//! and comparison semantics per C# spec.
use super::helpers::run_csharp;

#[test]
fn nullable_addition_with_both_operands_present_yields_sum() {
    assert_eq!(
        run_csharp(
            r#"
int? left = 4;
int? right = 6;
int? sum = left + right;
Console.WriteLine(sum.HasValue);
Console.WriteLine(sum.Value);
"#
        ),
        &["True", "10"]
    );
}

#[test]
fn nullable_addition_when_either_operand_null_yields_null() {
    assert_eq!(
        run_csharp(
            r#"
int? present = 5;
int? missing = null;
int? sum = present + missing;
Console.WriteLine(sum.HasValue);
"#
        ),
        &["False"]
    );
}

#[test]
fn nullable_equality_compares_values_not_references() {
    assert_eq!(
        run_csharp(
            r#"
int? a = 7;
int? b = 7;
Console.WriteLine(a == b);
int? c = null;
Console.WriteLine(a == c);
Console.WriteLine(c == null);
"#
        ),
        &["True", "False", "True"]
    );
}

#[test]
fn nullable_get_value_or_default_returns_zero_for_null_int() {
    assert_eq!(
        run_csharp(
            r#"
int? value = null;
Console.WriteLine(value.GetValueOrDefault());
Console.WriteLine(value.GetValueOrDefault(99));
"#
        ),
        &["0", "99"]
    );
}

#[test]
fn nullable_bool_logical_and_short_circuits_on_null() {
    assert_eq!(
        run_csharp(
            r#"
bool? t = true;
bool? n = null;
bool? f = false;
Console.WriteLine(t & n);
Console.WriteLine(n & f);
Console.WriteLine(f & t);
"#
        ),
        &["", "", "False"]
    );
}

#[test]
fn nullable_increment_operator_on_null_stays_null() {
    assert_eq!(
        run_csharp(
            r#"
int? value = null;
value++;
Console.WriteLine(value.HasValue);
"#
        ),
        &["False"]
    );
}

#[test]
fn nullable_increment_operator_on_value_increments_contents() {
    assert_eq!(
        run_csharp(
            r#"
int? value = 10;
value++;
Console.WriteLine(value);
"#
        ),
        &["11"]
    );
}

#[test]
fn nullable_coalesce_prefers_left_when_has_value() {
    assert_eq!(
        run_csharp(
            r#"
int? left = 8;
Console.WriteLine(left ?? 100);
"#
        ),
        &["8"]
    );
}

#[test]
fn nullable_coalesce_uses_right_when_left_is_null() {
    assert_eq!(
        run_csharp(
            r#"
int? left = null;
Console.WriteLine(left ?? 100);
"#
        ),
        &["100"]
    );
}

#[test]
fn nullable_chained_null_conditional_on_struct_member() {
    assert_eq!(
        run_csharp(
            r#"
struct Point { public int X; public int Y; }
Point? location = new Point { X = 2, Y = 3 };
Console.WriteLine(location?.X);
location = null;
Console.WriteLine(location?.X ?? -1);
"#
        ),
        &["2", "-1"]
    );
}

#[test]
fn nullable_compare_to_orders_values_with_null_as_smallest() {
    assert_eq!(
        run_csharp(
            r#"
int? low = 1;
int? high = 5;
int? missing = null;
Console.WriteLine(low.CompareTo(high));
Console.WriteLine(missing.CompareTo(low));
"#
        ),
        &["-1", "-1"]
    );
}

#[test]
fn nullable_conversion_from_literal_wraps_value_type() {
    assert_eq!(
        run_csharp(
            r#"
int? boxed = 42;
Console.WriteLine(boxed is int);
Console.WriteLine((int)boxed);
"#
        ),
        &["True", "42"]
    );
}
