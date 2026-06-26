//! `TimeSpan` construction, `Add`/`Subtract`/`Negate`, factory methods,
//! `Total*` properties, and `CompareTo` ordering.
//! GAP: arithmetic-focused coverage beyond datetime/timespan smoke tests.

use crate::csharp_cases;

csharp_cases! {
    timespan_from_days_sets_days_component => {
        r#"var span=System.TimeSpan.FromDays(3); Console.WriteLine(span.Days);"#,
        ["3"]
    };

    timespan_from_days_total_days => {
        r#"var span=System.TimeSpan.FromDays(2.5); Console.WriteLine(span.TotalDays);"#,
        ["2.5"]
    };

    timespan_from_hours_sets_hours_component => {
        r#"var span=System.TimeSpan.FromHours(5); Console.WriteLine(span.Hours);"#,
        ["5"]
    };

    timespan_from_hours_total_hours => {
        r#"var span=System.TimeSpan.FromHours(2.5); Console.WriteLine(span.TotalHours);"#,
        ["2.5"]
    };

    timespan_from_minutes_total_minutes => {
        r#"var span=System.TimeSpan.FromMinutes(90); Console.WriteLine(span.TotalMinutes);"#,
        ["90"]
    };

    timespan_from_seconds_total_seconds => {
        r#"var span=System.TimeSpan.FromSeconds(45); Console.WriteLine(span.TotalSeconds);"#,
        ["45"]
    };

    timespan_from_milliseconds_total_milliseconds => {
        r#"var span=System.TimeSpan.FromMilliseconds(1500); Console.WriteLine(span.TotalMilliseconds);"#,
        ["1500"]
    };

    timespan_from_ticks_reads_ticks_property => {
        r#"var span=System.TimeSpan.FromTicks(10000000); Console.WriteLine(span.Ticks);"#,
        ["10000000"]
    };

    timespan_add_method_combines_durations => {
        r#"var a=System.TimeSpan.FromHours(1); var b=System.TimeSpan.FromMinutes(30); Console.WriteLine(a.Add(b).TotalMinutes);"#,
        ["90"]
    };

    timespan_subtract_method_returns_difference => {
        r#"var a=System.TimeSpan.FromMinutes(50); var b=System.TimeSpan.FromMinutes(20); Console.WriteLine(a.Subtract(b).TotalMinutes);"#,
        ["30"]
    };

    timespan_add_operator_two_positive_spans => {
        r#"var a=System.TimeSpan.FromDays(1); var b=System.TimeSpan.FromHours(12); Console.WriteLine((a+b).TotalHours);"#,
        ["36"]
    };

    timespan_subtract_operator_positive_result => {
        r#"var a=System.TimeSpan.FromHours(3); var b=System.TimeSpan.FromHours(1); Console.WriteLine((a-b).TotalHours);"#,
        ["2"]
    };

    timespan_negate_operator_flips_sign => {
        r#"var span=System.TimeSpan.FromMinutes(15); Console.WriteLine((-span).TotalMinutes);"#,
        ["-15"]
    };

    timespan_negate_method_matches_operator => {
        r#"var span=System.TimeSpan.FromSeconds(8); Console.WriteLine(span.Negate().TotalSeconds);"#,
        ["-8"]
    };

    timespan_double_negate_returns_original => {
        r#"var span=System.TimeSpan.FromSeconds(8); Console.WriteLine((-(-span)).TotalSeconds);"#,
        ["8"]
    };

    timespan_compare_to_shorter_is_negative => {
        r#"var left=System.TimeSpan.FromMinutes(1); var right=System.TimeSpan.FromMinutes(2); Console.WriteLine(left.CompareTo(right));"#,
        ["-1"]
    };

    timespan_compare_to_longer_is_positive => {
        r#"var left=System.TimeSpan.FromMinutes(3); var right=System.TimeSpan.FromMinutes(2); Console.WriteLine(left.CompareTo(right));"#,
        ["1"]
    };

    timespan_compare_to_equal_is_zero => {
        r#"var left=System.TimeSpan.FromSeconds(5); var right=System.TimeSpan.FromSeconds(5); Console.WriteLine(left.CompareTo(right));"#,
        ["0"]
    };

    timespan_static_compare_matches_instance => {
        r#"var left=System.TimeSpan.FromHours(1); var right=System.TimeSpan.FromHours(2); Console.WriteLine(System.TimeSpan.Compare(left,right));"#,
        ["-1"]
    };

    timespan_zero_total_seconds => {
        r#"Console.WriteLine(System.TimeSpan.Zero.TotalSeconds);"#,
        ["0"]
    };

    timespan_zero_is_neutral_for_addition => {
        r#"var span=System.TimeSpan.FromMinutes(10); Console.WriteLine(span.Add(System.TimeSpan.Zero).TotalMinutes);"#,
        ["10"]
    };

    timespan_constructor_days_hours_minutes => {
        r#"var span=new System.TimeSpan(1,2,30); Console.WriteLine(span.Days); Console.WriteLine(span.Hours); Console.WriteLine(span.Minutes);"#,
        ["1", "2", "30"]
    };

    timespan_constructor_with_seconds => {
        r#"var span=new System.TimeSpan(0,1,2,3); Console.WriteLine(span.Hours); Console.WriteLine(span.Minutes); Console.WriteLine(span.Seconds);"#,
        ["1", "2", "3"]
    };

    timespan_constructor_with_milliseconds => {
        r#"var span=new System.TimeSpan(0,0,0,0,250); Console.WriteLine(span.Milliseconds);"#,
        ["250"]
    };

    timespan_duration_of_negative_span => {
        r#"var span=System.TimeSpan.FromHours(-2); Console.WriteLine(span.Duration().TotalHours);"#,
        ["2"]
    };

    timespan_duration_of_positive_span_unchanged => {
        r#"var span=System.TimeSpan.FromHours(2); Console.WriteLine(span.Duration().TotalHours);"#,
        ["2"]
    };

    timespan_greater_than_operator => {
        r#"var left=System.TimeSpan.FromMinutes(5); var right=System.TimeSpan.FromMinutes(2); Console.WriteLine(left>right);"#,
        ["True"]
    };

    timespan_less_than_operator => {
        r#"var left=System.TimeSpan.FromMinutes(2); var right=System.TimeSpan.FromMinutes(5); Console.WriteLine(left<right);"#,
        ["True"]
    };

    timespan_equality_same_duration => {
        r#"var a=System.TimeSpan.FromMinutes(60); var b=System.TimeSpan.FromHours(1); Console.WriteLine(a==b);"#,
        ["True"]
    };

    timespan_inequality_different_duration => {
        r#"var a=System.TimeSpan.FromMinutes(60); var b=System.TimeSpan.FromHours(2); Console.WriteLine(a!=b);"#,
        ["True"]
    };

    timespan_add_negative_span_via_negate => {
        r#"var baseSpan=System.TimeSpan.FromHours(2); var delta=System.TimeSpan.FromMinutes(30); Console.WriteLine(baseSpan.Add(-delta).TotalMinutes);"#,
        ["90"]
    };

    timespan_subtract_negative_span_adds => {
        r#"var baseSpan=System.TimeSpan.FromHours(1); var delta=System.TimeSpan.FromMinutes(-30); Console.WriteLine(baseSpan.Subtract(delta).TotalMinutes);"#,
        ["90"]
    };

    timespan_from_hours_overflows_into_days => {
        r#"var span=System.TimeSpan.FromHours(25); Console.WriteLine(span.Days); Console.WriteLine(span.Hours);"#,
        ["1", "1"]
    };

    timespan_total_hours_from_minutes => {
        r#"var span=System.TimeSpan.FromMinutes(120); Console.WriteLine(span.TotalHours);"#,
        ["2"]
    };

    timespan_total_minutes_from_hours => {
        r#"var span=System.TimeSpan.FromHours(1.5); Console.WriteLine(span.TotalMinutes);"#,
        ["90"]
    };

    timespan_total_seconds_from_milliseconds => {
        r#"var span=System.TimeSpan.FromMilliseconds(2500); Console.WriteLine(span.TotalSeconds);"#,
        ["2.5"]
    };

    timespan_from_days_fractional_hours => {
        r#"var span=System.TimeSpan.FromDays(0.5); Console.WriteLine(span.TotalHours);"#,
        ["12"]
    };

    timespan_add_multiple_increments => {
        r#"var span=System.TimeSpan.Zero; span=span.Add(System.TimeSpan.FromMinutes(10)); span=span.Add(System.TimeSpan.FromMinutes(5)); Console.WriteLine(span.TotalMinutes);"#,
        ["15"]
    };

    timespan_subtract_to_zero => {
        r#"var span=System.TimeSpan.FromMinutes(10); Console.WriteLine(span.Subtract(span).TotalMinutes);"#,
        ["0"]
    };

    timespan_compare_negative_to_positive => {
        r#"var left=System.TimeSpan.FromMinutes(-5); var right=System.TimeSpan.FromMinutes(5); Console.WriteLine(left.CompareTo(right));"#,
        ["-1"]
    };

    timespan_negate_zero_is_zero => {
        r#"Console.WriteLine((-System.TimeSpan.Zero).TotalSeconds);"#,
        ["0"]
    };

    timespan_from_seconds_negative_total => {
        r#"var span=System.TimeSpan.FromSeconds(-30); Console.WriteLine(span.TotalSeconds);"#,
        ["-30"]
    };

    timespan_from_minutes_negative_total => {
        r#"var span=System.TimeSpan.FromMinutes(-2); Console.WriteLine(span.TotalMinutes);"#,
        ["-2"]
    };

    timespan_from_hours_negative_total => {
        r#"var span=System.TimeSpan.FromHours(-1); Console.WriteLine(span.TotalHours);"#,
        ["-1"]
    };

    timespan_from_days_negative_total => {
        r#"var span=System.TimeSpan.FromDays(-1); Console.WriteLine(span.TotalDays);"#,
        ["-1"]
    };

    timespan_max_value_is_positive => {
        r#"Console.WriteLine(System.TimeSpan.MaxValue.TotalDays>0);"#,
        ["True"]
    };

    timespan_min_value_is_negative => {
        r#"Console.WriteLine(System.TimeSpan.MinValue.TotalDays<0);"#,
        ["True"]
    };

    timespan_to_string_positive_hms => {
        r#"Console.WriteLine(System.TimeSpan.FromHours(1).Add(System.TimeSpan.FromMinutes(2)).Add(System.TimeSpan.FromSeconds(3)).ToString());"#,
        ["01:02:03"]
    };

    timespan_add_commutative_via_total_minutes => {
        r#"var a=System.TimeSpan.FromMinutes(10); var b=System.TimeSpan.FromMinutes(20); Console.WriteLine((a+b).TotalMinutes); Console.WriteLine((b+a).TotalMinutes);"#,
        ["30", "30"]
    };

    timespan_subtract_self_compare_zero => {
        r#"var span=System.TimeSpan.FromDays(1); Console.WriteLine(span.Subtract(span).CompareTo(System.TimeSpan.Zero));"#,
        ["0"]
    };

    timespan_from_ticks_zero => {
        r#"Console.WriteLine(System.TimeSpan.FromTicks(0).TotalSeconds);"#,
        ["0"]
    };

    timespan_milliseconds_component_after_from_seconds => {
        r#"var span=System.TimeSpan.FromSeconds(1.5); Console.WriteLine(span.Milliseconds);"#,
        ["500"]
    };
}
