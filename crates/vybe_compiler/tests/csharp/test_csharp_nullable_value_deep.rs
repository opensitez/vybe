//! Nullable value types (`T?`): `HasValue`, `Value`, `GetValueOrDefault`, `??`,
//! and lifted operators for `int?`, `decimal?`, and related primitives.
//! GAP: deep nullable semantics beyond basic smoke tests.

use crate::csharp_cases;

csharp_cases! {
    nullable_int_has_value_true_for_zero => {
        r#"int? n=0; Console.WriteLine(n.HasValue);"#,
        ["True"]
    };

    nullable_int_has_value_false_for_null_literal => {
        r#"int? n=null; Console.WriteLine(n.HasValue);"#,
        ["False"]
    };

    nullable_int_value_reads_stored_integer => {
        r#"int? n=123; Console.WriteLine(n.Value);"#,
        ["123"]
    };

    nullable_int_get_value_or_default_zero_when_null => {
        r#"int? n=null; Console.WriteLine(n.GetValueOrDefault());"#,
        ["0"]
    };

    nullable_int_get_value_or_default_custom_fallback => {
        r#"int? n=null; Console.WriteLine(n.GetValueOrDefault(88));"#,
        ["88"]
    };

    nullable_int_get_value_or_default_returns_value_when_present => {
        r#"int? n=15; Console.WriteLine(n.GetValueOrDefault(88));"#,
        ["15"]
    };

    nullable_int_null_coalescing_right_when_null => {
        r#"int? n=null; Console.WriteLine(n??42);"#,
        ["42"]
    };

    nullable_int_null_coalescing_left_when_present => {
        r#"int? n=7; Console.WriteLine(n??42);"#,
        ["7"]
    };

    nullable_int_null_coalescing_assignment_sets_when_null => {
        r#"int? n=null; n??=9; Console.WriteLine(n);"#,
        ["9"]
    };

    nullable_int_null_coalescing_assignment_skips_when_present => {
        r#"int? n=3; n??=9; Console.WriteLine(n);"#,
        ["3"]
    };

    nullable_int_lifted_addition_both_present => {
        r#"int? a=10; int? b=5; Console.WriteLine(a+b);"#,
        ["15"]
    };

    nullable_int_lifted_addition_one_null => {
        r#"int? a=10; int? b=null; Console.WriteLine((a+b).HasValue);"#,
        ["False"]
    };

    nullable_int_lifted_subtraction_both_present => {
        r#"int? a=20; int? b=6; Console.WriteLine(a-b);"#,
        ["14"]
    };

    nullable_int_lifted_subtraction_one_null => {
        r#"int? a=null; int? b=6; Console.WriteLine((a-b).HasValue);"#,
        ["False"]
    };

    nullable_int_lifted_multiplication_both_present => {
        r#"int? a=4; int? b=5; Console.WriteLine(a*b);"#,
        ["20"]
    };

    nullable_int_lifted_multiplication_one_null => {
        r#"int? a=4; int? b=null; Console.WriteLine((a*b).HasValue);"#,
        ["False"]
    };

    nullable_int_lifted_division_both_present => {
        r#"int? a=20; int? b=4; Console.WriteLine(a/b);"#,
        ["5"]
    };

    nullable_int_lifted_modulo_both_present => {
        r#"int? a=10; int? b=3; Console.WriteLine(a%b);"#,
        ["1"]
    };

    nullable_int_lifted_modulo_one_null => {
        r#"int? a=10; int? b=null; Console.WriteLine((a%b).HasValue);"#,
        ["False"]
    };

    nullable_int_lifted_unary_plus => {
        r#"int? n=8; Console.WriteLine(+n);"#,
        ["8"]
    };

    nullable_int_lifted_unary_minus => {
        r#"int? n=8; Console.WriteLine(-n);"#,
        ["-8"]
    };

    nullable_int_lifted_unary_minus_null => {
        r#"int? n=null; Console.WriteLine((-n).HasValue);"#,
        ["False"]
    };

    nullable_int_equality_same_values => {
        r#"int? a=5; int? b=5; Console.WriteLine(a==b);"#,
        ["True"]
    };

    nullable_int_equality_null_to_null => {
        r#"int? a=null; int? b=null; Console.WriteLine(a==b);"#,
        ["True"]
    };

    nullable_int_equality_value_to_null => {
        r#"int? a=5; int? b=null; Console.WriteLine(a==b);"#,
        ["False"]
    };

    nullable_int_inequality_value_to_null => {
        r#"int? a=5; int? b=null; Console.WriteLine(a!=b);"#,
        ["True"]
    };

    nullable_int_less_than_both_present => {
        r#"int? a=2; int? b=5; Console.WriteLine(a<b);"#,
        ["True"]
    };

    nullable_int_less_than_one_null => {
        r#"int? a=2; int? b=null; Console.WriteLine((a<b).HasValue);"#,
        ["False"]
    };

    nullable_int_greater_than_both_present => {
        r#"int? a=9; int? b=3; Console.WriteLine(a>b);"#,
        ["True"]
    };

    nullable_int_greater_or_equal_equal_values => {
        r#"int? a=4; int? b=4; Console.WriteLine(a>=b);"#,
        ["True"]
    };

    nullable_int_less_or_equal_equal_values => {
        r#"int? a=4; int? b=4; Console.WriteLine(a<=b);"#,
        ["True"]
    };

    nullable_int_increment_on_value => {
        r#"int? n=5; n++; Console.WriteLine(n);"#,
        ["6"]
    };

    nullable_int_increment_on_null => {
        r#"int? n=null; n++; Console.WriteLine(n.HasValue);"#,
        ["False"]
    };

    nullable_int_decrement_on_value => {
        r#"int? n=5; n--; Console.WriteLine(n);"#,
        ["4"]
    };

    nullable_int_decrement_on_null => {
        r#"int? n=null; n--; Console.WriteLine(n.HasValue);"#,
        ["False"]
    };

    nullable_int_explicit_cast_to_int => {
        r#"int? n=12; int x=(int)n; Console.WriteLine(x);"#,
        ["12"]
    };

    nullable_int_implicit_conversion_from_int => {
        r#"int x=33; int? n=x; Console.WriteLine(n);"#,
        ["33"]
    };

    nullable_decimal_has_value_when_assigned => {
        r#"decimal? d=1.5m; Console.WriteLine(d.HasValue);"#,
        ["True"]
    };

    nullable_decimal_has_value_false_when_null => {
        r#"decimal? d=null; Console.WriteLine(d.HasValue);"#,
        ["False"]
    };

    nullable_decimal_value_reads_fraction => {
        r#"decimal? d=2.25m; Console.WriteLine(d.Value);"#,
        ["2.25"]
    };

    nullable_decimal_get_value_or_default_zero => {
        r#"decimal? d=null; Console.WriteLine(d.GetValueOrDefault());"#,
        ["0"]
    };

    nullable_decimal_get_value_or_default_custom => {
        r#"decimal? d=null; Console.WriteLine(d.GetValueOrDefault(9.99m));"#,
        ["9.99"]
    };

    nullable_decimal_null_coalescing => {
        r#"decimal? d=null; Console.WriteLine(d??3.5m);"#,
        ["3.5"]
    };

    nullable_decimal_lifted_addition => {
        r#"decimal? a=0.1m; decimal? b=0.2m; Console.WriteLine(a+b);"#,
        ["0.3"]
    };

    nullable_decimal_lifted_subtraction => {
        r#"decimal? a=5.0m; decimal? b=2.0m; Console.WriteLine(a-b);"#,
        ["3.0"]
    };

    nullable_decimal_lifted_multiplication => {
        r#"decimal? a=2.5m; decimal? b=4m; Console.WriteLine(a*b);"#,
        ["10.0"]
    };

    nullable_decimal_lifted_division => {
        r#"decimal? a=7.5m; decimal? b=2.5m; Console.WriteLine(a/b);"#,
        ["3"]
    };

    nullable_decimal_lifted_addition_one_null => {
        r#"decimal? a=1m; decimal? b=null; Console.WriteLine((a+b).HasValue);"#,
        ["False"]
    };

    nullable_decimal_equality_exact_scale => {
        r#"decimal? a=1.0m; decimal? b=1.00m; Console.WriteLine(a==b);"#,
        ["True"]
    };

    nullable_decimal_comparison_less_than => {
        r#"decimal? a=1.2m; decimal? b=1.3m; Console.WriteLine(a<b);"#,
        ["True"]
    };

    nullable_bool_has_value_for_true => {
        r#"bool? flag=true; Console.WriteLine(flag.HasValue); Console.WriteLine(flag.Value);"#,
        ["True", "True"]
    };

    nullable_bool_has_value_false_for_null => {
        r#"bool? flag=null; Console.WriteLine(flag.HasValue);"#,
        ["False"]
    };

    nullable_bool_null_coalescing => {
        r#"bool? flag=null; Console.WriteLine(flag??true);"#,
        ["True"]
    };

    nullable_bool_lifted_and_both_true => {
        r#"bool? a=true; bool? b=true; Console.WriteLine(a&b);"#,
        ["True"]
    };

    nullable_bool_lifted_and_one_null => {
        r#"bool? a=true; bool? b=null; Console.WriteLine((a&b).HasValue);"#,
        ["False"]
    };

    nullable_bool_lifted_or_one_true => {
        r#"bool? a=true; bool? b=null; Console.WriteLine(a|b);"#,
        ["True"]
    };

    nullable_bool_lifted_or_both_null => {
        r#"bool? a=null; bool? b=null; Console.WriteLine((a|b).HasValue);"#,
        ["False"]
    };

    nullable_double_get_value_or_default => {
        r#"double? d=null; Console.WriteLine(d.GetValueOrDefault(3.14));"#,
        ["3.14"]
    };

    nullable_double_lifted_addition => {
        r#"double? a=1.5; double? b=2.5; Console.WriteLine(a+b);"#,
        ["4"]
    };

    nullable_long_null_coalescing => {
        r#"long? n=null; Console.WriteLine(n??100L);"#,
        ["100"]
    };

    nullable_int_to_string_via_value => {
        r#"int? n=55; Console.WriteLine(n.ToString());"#,
        ["55"]
    };

    nullable_int_to_string_null_prints_empty => {
        r#"int? n=null; Console.WriteLine(n.ToString().Length);"#,
        ["0"]
    };
}
