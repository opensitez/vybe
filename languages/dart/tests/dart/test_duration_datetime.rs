//! Dart `Duration` and `DateTime`: construction, component getters, comparison,
//! arithmetic, and calendar operations.

dart_cases! {
    duration_milliseconds_constructor => {
        r#"void main() {
  var d = Duration(milliseconds: 500);
  print(d.inMilliseconds);
}"#,
        ["500"]
    };

    duration_seconds_constructor => {
        r#"void main() {
  var d = Duration(seconds: 30);
  print(d.inSeconds);
}"#,
        ["30"]
    };

    duration_minutes_constructor => {
        r#"void main() {
  var d = Duration(minutes: 5);
  print(d.inMinutes);
}"#,
        ["5"]
    };

    duration_hours_constructor => {
        r#"void main() {
  var d = Duration(hours: 2);
  print(d.inHours);
}"#,
        ["2"]
    };

    duration_days_constructor => {
        r#"void main() {
  var d = Duration(days: 1);
  print(d.inDays);
}"#,
        ["1"]
    };

    duration_zero_constant => {
        r#"void main() {
  var d = Duration.zero;
  print(d.inMilliseconds);
}"#,
        ["0"]
    };

    duration_multi_component_sum => {
        r#"void main() {
  var d = Duration(hours: 1, minutes: 30, seconds: 15);
  print(d.inMinutes);
}"#,
        ["90"]
    };

    duration_days_hours_minutes => {
        r#"void main() {
  var d = Duration(days: 1, hours: 2, minutes: 3);
  print(d.inHours);
}"#,
        ["26"]
    };

    duration_negative_milliseconds => {
        r#"void main() {
  var d = Duration(milliseconds: -250);
  print(d.inMilliseconds);
}"#,
        ["-250"]
    };

    duration_negative_seconds => {
        r#"void main() {
  var d = Duration(seconds: -10);
  print(d.inSeconds);
}"#,
        ["-10"]
    };

    duration_in_seconds_from_minutes => {
        r#"void main() {
  var d = Duration(minutes: 2);
  print(d.inSeconds);
}"#,
        ["120"]
    };

    duration_in_hours_from_minutes => {
        r#"void main() {
  var d = Duration(minutes: 90);
  print(d.inHours);
}"#,
        ["1"]
    };

    duration_in_days_from_hours => {
        r#"void main() {
  var d = Duration(hours: 48);
  print(d.inDays);
}"#,
        ["2"]
    };

    duration_fractional_minutes => {
        r#"void main() {
  var d = Duration(seconds: 90);
  print(d.inMinutes);
}"#,
        ["1"]
    };

    duration_compare_equal => {
        r#"void main() {
  var a = Duration(minutes: 5);
  var b = Duration(minutes: 5);
  print(a.compareTo(b));
}"#,
        ["0"]
    };

    duration_compare_less_than => {
        r#"void main() {
  var a = Duration(seconds: 10);
  var b = Duration(seconds: 20);
  print(a.compareTo(b));
}"#,
        ["-1"]
    };

    duration_compare_greater_than => {
        r#"void main() {
  var a = Duration(minutes: 10);
  var b = Duration(minutes: 5);
  print(a.compareTo(b));
}"#,
        ["1"]
    };

    duration_addition_combines_minutes => {
        r#"void main() {
  var a = Duration(minutes: 5);
  var b = Duration(minutes: 3);
  print((a + b).inMinutes);
}"#,
        ["8"]
    };

    duration_addition_hours_and_minutes => {
        r#"void main() {
  var a = Duration(hours: 1);
  var b = Duration(minutes: 30);
  print((a + b).inMinutes);
}"#,
        ["90"]
    };

    duration_subtraction_yields_remaining => {
        r#"void main() {
  var a = Duration(minutes: 45);
  var b = Duration(minutes: 15);
  print((a - b).inMinutes);
}"#,
        ["30"]
    };

    duration_subtraction_to_zero => {
        r#"void main() {
  var a = Duration(seconds: 10);
  var b = Duration(seconds: 10);
  print((a - b).inSeconds);
}"#,
        ["0"]
    };

    duration_abs_of_negative => {
        r#"void main() {
  var d = Duration(seconds: -9);
  print(d.abs().inSeconds);
}"#,
        ["9"]
    };

    duration_is_negative_true => {
        r#"void main() {
  var d = Duration(minutes: -1);
  print(d.isNegative);
}"#,
        ["true"]
    };

    duration_is_negative_false => {
        r#"void main() {
  var d = Duration(minutes: 1);
  print(d.isNegative);
}"#,
        ["false"]
    };

    duration_zero_is_not_negative => {
        r#"void main() {
  print(Duration.zero.isNegative);
}"#,
        ["false"]
    };

    duration_milliseconds_and_seconds_combo => {
        r#"void main() {
  var d = Duration(seconds: 1, milliseconds: 500);
  print(d.inMilliseconds);
}"#,
        ["1500"]
    };

    duration_days_to_minutes => {
        r#"void main() {
  var d = Duration(days: 2);
  print(d.inMinutes);
}"#,
        ["2880"]
    };

    duration_negate_flips_sign => {
        r#"void main() {
  var d = Duration(minutes: 4);
  print(d.negate().inMinutes);
}"#,
        ["-4"]
    };

    duration_double_negate_restores => {
        r#"void main() {
  var d = Duration(hours: 3);
  print(d.negate().negate().inHours);
}"#,
        ["3"]
    };

    duration_add_zero_is_identity => {
        r#"void main() {
  var d = Duration(minutes: 7);
  print((d + Duration.zero).inMinutes);
}"#,
        ["7"]
    };

    datetime_constructor_year_month_day => {
        r#"void main() {
  var dt = DateTime(2024, 6, 15);
  print(dt.year);
  print(dt.month);
  print(dt.day);
}"#,
        ["2024", "6", "15"]
    };

    datetime_constructor_with_time_components => {
        r#"void main() {
  var dt = DateTime(2024, 1, 1, 14, 30, 45);
  print(dt.hour);
  print(dt.minute);
  print(dt.second);
}"#,
        ["14", "30", "45"]
    };

    datetime_weekday_saturday => {
        r#"void main() {
  var dt = DateTime(2024, 6, 15);
  print(dt.weekday);
}"#,
        ["6"]
    };

    datetime_weekday_monday => {
        r#"void main() {
  var dt = DateTime(2024, 6, 3);
  print(dt.weekday);
}"#,
        ["1"]
    };

    datetime_add_duration_days => {
        r#"void main() {
  var dt = DateTime(2024, 1, 1);
  var later = dt.add(Duration(days: 10));
  print(later.day);
}"#,
        ["11"]
    };

    datetime_add_duration_hours => {
        r#"void main() {
  var dt = DateTime(2024, 1, 1, 10, 0, 0);
  var later = dt.add(Duration(hours: 5));
  print(later.hour);
}"#,
        ["15"]
    };

    datetime_subtract_duration_days => {
        r#"void main() {
  var dt = DateTime(2024, 3, 15);
  var earlier = dt.subtract(Duration(days: 5));
  print(earlier.day);
}"#,
        ["10"]
    };

    datetime_subtract_duration_hours => {
        r#"void main() {
  var dt = DateTime(2024, 1, 1, 12, 0, 0);
  var earlier = dt.subtract(Duration(hours: 2));
  print(earlier.hour);
}"#,
        ["10"]
    };

    datetime_difference_in_days => {
        r#"void main() {
  var start = DateTime(2024, 1, 1);
  var end = DateTime(2024, 1, 11);
  print(end.difference(start).inDays);
}"#,
        ["10"]
    };

    datetime_difference_in_hours => {
        r#"void main() {
  var start = DateTime(2024, 1, 1, 0, 0, 0);
  var end = DateTime(2024, 1, 1, 6, 0, 0);
  print(end.difference(start).inHours);
}"#,
        ["6"]
    };

    datetime_difference_reverse_is_negative => {
        r#"void main() {
  var later = DateTime(2024, 2, 1);
  var earlier = DateTime(2024, 1, 1);
  print(later.difference(earlier).inDays);
  print(earlier.difference(later).inDays);
}"#,
        ["31", "-31"]
    };

    datetime_is_before_true => {
        r#"void main() {
  var a = DateTime(2024, 1, 1);
  var b = DateTime(2024, 1, 2);
  print(a.isBefore(b));
}"#,
        ["true"]
    };

    datetime_is_before_false => {
        r#"void main() {
  var a = DateTime(2024, 1, 2);
  var b = DateTime(2024, 1, 1);
  print(a.isBefore(b));
}"#,
        ["false"]
    };

    datetime_is_after_true => {
        r#"void main() {
  var a = DateTime(2024, 3, 1);
  var b = DateTime(2024, 2, 1);
  print(a.isAfter(b));
}"#,
        ["true"]
    };

    datetime_is_after_false => {
        r#"void main() {
  var a = DateTime(2024, 1, 1);
  var b = DateTime(2024, 2, 1);
  print(a.isAfter(b));
}"#,
        ["false"]
    };

    datetime_is_at_same_moment_as_equal => {
        r#"void main() {
  var a = DateTime(2024, 5, 17, 8, 0, 0);
  var b = DateTime(2024, 5, 17, 8, 0, 0);
  print(a.isAtSameMomentAs(b));
}"#,
        ["true"]
    };

    datetime_is_at_same_moment_as_different => {
        r#"void main() {
  var a = DateTime(2024, 5, 17, 8, 0, 0);
  var b = DateTime(2024, 5, 17, 9, 0, 0);
  print(a.isAtSameMomentAs(b));
}"#,
        ["false"]
    };

    datetime_add_then_difference_roundtrip => {
        r#"void main() {
  var base = DateTime(2024, 7, 1);
  var span = Duration(days: 14);
  var target = base.add(span);
  print(target.difference(base).inDays);
}"#,
        ["14"]
    };

    datetime_month_boundary_add_days => {
        r#"void main() {
  var dt = DateTime(2024, 1, 30);
  var later = dt.add(Duration(days: 2));
  print(later.month);
  print(later.day);
}"#,
        ["2", "1"]
    };

    datetime_year_boundary_add_days => {
        r#"void main() {
  var dt = DateTime(2023, 12, 31);
  var later = dt.add(Duration(days: 1));
  print(later.year);
  print(later.month);
  print(later.day);
}"#,
        ["2024", "1", "1"]
    };

    datetime_leap_year_february_day => {
        r#"void main() {
  var dt = DateTime(2024, 2, 29);
  print(dt.day);
  print(dt.month);
}"#,
        ["29", "2"]
    };

    datetime_utc_constructor_fields => {
        r#"void main() {
  var dt = DateTime.utc(2024, 6, 15, 12, 0, 0);
  print(dt.year);
  print(dt.month);
  print(dt.isUtc);
}"#,
        ["2024", "6", "true"]
    };

    datetime_local_is_not_utc => {
        r#"void main() {
  var dt = DateTime(2024, 6, 15);
  print(dt.isUtc);
}"#,
        ["false"]
    };

    datetime_compare_chronological_order => {
        r#"void main() {
  var early = DateTime(2024, 1, 1);
  var late = DateTime(2024, 12, 31);
  print(early.compareTo(late));
  print(late.compareTo(early));
}"#,
        ["-1", "1"]
    };

    datetime_compare_equal_returns_zero => {
        r#"void main() {
  var a = DateTime(2024, 4, 4);
  var b = DateTime(2024, 4, 4);
  print(a.compareTo(b));
}"#,
        ["0"]
    };

    duration_compare_with_zero => {
        r#"void main() {
  var d = Duration(minutes: 1);
  print(d.compareTo(Duration.zero));
}"#,
        ["1"]
    };

    duration_add_to_datetime_minutes => {
        r#"void main() {
  var dt = DateTime(2024, 1, 1, 0, 0, 0);
  var later = dt.add(Duration(minutes: 45));
  print(later.minute);
}"#,
        ["45"]
    };
}
