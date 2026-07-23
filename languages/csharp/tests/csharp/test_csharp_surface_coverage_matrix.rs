use super::helpers::run_csharp;

const INT_CASES: &[(i64, i64)] = &[
    (1, 2),
    (2, 3),
    (3, 4),
    (4, 5),
    (5, 6),
    (6, 7),
    (7, 8),
    (8, 9),
    (9, 10),
    (10, 11),
    (11, 12),
    (12, 13),
    (13, 14),
    (14, 15),
    (15, 16),
    (16, 17),
    (17, 18),
    (18, 19),
    (19, 20),
    (20, 21),
    (21, 22),
];

fn assert_matrix_case(builder: impl Fn(i64, i64) -> String) {
    for &(left, right) in INT_CASES {
        let src = builder(left, right);
        assert_eq!(run_csharp(&src), vec!["True".to_string()]);
    }
}

macro_rules! matrix_case {
    ($name:ident, $builder:expr) => {
        #[test]
        fn $name() {
            assert_matrix_case($builder);
        }
    };
}

matrix_case!(matrix_arithmetic_add_inverse, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(left + right - right == left && right > 0);"
    )
});

matrix_case!(matrix_arithmetic_sub_identity, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(left - 0 == left && right >= 0);"
    )
});

matrix_case!(matrix_arithmetic_mul_identity, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(left * 1 == left && right >= 0);"
    )
});

matrix_case!(matrix_arithmetic_mul_left_one, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(1 * left == left && right >= 0);"
    )
});

matrix_case!(matrix_arithmetic_div_identity, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(left / 1 == left && right >= 0);"
    )
});

matrix_case!(matrix_arithmetic_mul_div_roundtrip, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(left * right / right == left && left > 0 && right > 0);"
    )
});

matrix_case!(matrix_arithmetic_negation_involution, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(-(-left) == left && right >= 0);"
    )
});

matrix_case!(matrix_arithmetic_abs_non_negative, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(System.Math.Abs(left) >= 0 && right >= 0);"
    )
});

matrix_case!(matrix_arithmetic_abs_identity, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(System.Math.Abs(left - left) == 0 && right >= 0);"
    )
});

matrix_case!(matrix_arithmetic_min_bound, |left, right| {
    format!(
        "long left = {left}; long right = {right}; long min = System.Math.Min(left, right); Console.WriteLine((min == left || min == right) && min <= left && min <= right);"
    )
});

matrix_case!(matrix_arithmetic_max_bound, |left, right| {
    format!(
        "long left = {left}; long right = {right}; long max = System.Math.Max(left, right); Console.WriteLine((max == left || max == right) && max >= left && max >= right);"
    )
});

matrix_case!(matrix_boolean_identity_combo, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(((left == left) && (right == right) && left != 0));"
    )
});

matrix_case!(matrix_arithmetic_increment_decrement, |left, right| {
    format!(
        "long left = {left}; long right = {right}; long next = left + 1; long back = next - 1; Console.WriteLine(back == left && right >= 0);"
    )
});

matrix_case!(matrix_bitwise_and_zero, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left & 0L) == 0L && right >= 0);"
    )
});

matrix_case!(matrix_bitwise_or_zero, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left | 0L) == left && right >= 0);"
    )
});

matrix_case!(matrix_bitwise_xor_zero, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left ^ 0L) == left && right >= 0);"
    )
});

matrix_case!(matrix_bitwise_not_twice, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((~(~left)) == left && right >= 0);"
    )
});

matrix_case!(matrix_bitwise_and_self, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left & left) == left && right >= 0);"
    )
});

matrix_case!(matrix_bitwise_or_self, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left | left) == left && right >= 0);"
    )
});

matrix_case!(matrix_bitwise_xor_self, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left ^ left) == 0L && right >= 0);"
    )
});

matrix_case!(matrix_bitwise_left_zero_shift, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left << 0) == left && right >= 0);"
    )
});

matrix_case!(matrix_bitwise_right_zero_shift, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left >> 0) == left && right >= 0);"
    )
});

matrix_case!(matrix_bitwise_xor_partner, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(((left ^ right) ^ right) == left && right >= 0);"
    )
});

matrix_case!(matrix_bitwise_or_commute, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left | right) == (right | left) && right >= 0);"
    )
});

matrix_case!(matrix_bitwise_and_commute, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left & right) == (right & left) && right >= 0);"
    )
});

matrix_case!(matrix_bitwise_xor_commute, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left ^ right) == (right ^ left) && right >= 0);"
    )
});

matrix_case!(matrix_compare_reflexive, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(left == left && right >= 0);"
    )
});

matrix_case!(matrix_compare_symmetric, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left == right) == (right == left));"
    )
});

matrix_case!(matrix_compare_dichotomy, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left == right) || (left < right) || (left > right));"
    )
});

matrix_case!(matrix_compare_total, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left < right) || (left >= right));"
    )
});

matrix_case!(matrix_compare_le_identity, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(left <= left + right && right >= 0);"
    )
});

matrix_case!(matrix_compare_ge_identity, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(left + 0 >= left && right >= 0);"
    )
});

matrix_case!(matrix_bool_tautology_true, |left, right| {
    format!("long left = {left}; long right = {right}; Console.WriteLine(left > 0 || left <= 0);")
});

matrix_case!(matrix_bool_tautology_false, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(!(left > 0 || left <= 0) == false && right >= 0);"
    )
});

matrix_case!(matrix_bool_not_involution, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left > 0) == !!(left > 0) && right >= 0);"
    )
});

matrix_case!(matrix_bool_and_positive, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left > 0 && right > 0) == true);"
    )
});

matrix_case!(matrix_bool_or_positive, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left > 0 || right > 0) == true);"
    )
});

matrix_case!(matrix_bool_xor_false, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine((left > 0) ^ false == (left > 0));"
    )
});

matrix_case!(matrix_bool_conditional_pos, |left, right| {
    format!("long left = {left}; long right = {right}; Console.WriteLine(left > 0 ? true : false);")
});

matrix_case!(matrix_bool_conditional_nested, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(left > 0 ? (right > 0 ? true : false) : false);"
    )
});

matrix_case!(matrix_string_to_string_roundtrip, |left, right| {
    format!(
        "long left = {left}; long right = {right}; string value = left.ToString(); Console.WriteLine(value == (left).ToString() && right >= 0);"
    )
});

matrix_case!(matrix_string_length_additivity, |left, right| {
    format!(
        "long left = {left}; long right = {right}; string leftText = left.ToString(); string rightText = right.ToString(); Console.WriteLine((leftText + rightText).Length == leftText.Length + rightText.Length);"
    )
});

matrix_case!(matrix_string_concat_contains, |left, right| {
    format!(
        "long left = {left}; long right = {right}; string leftText = left.ToString(); string rightText = right.ToString(); Console.WriteLine((leftText + rightText).Contains(leftText) && (leftText + rightText).Contains(rightText));"
    )
});

matrix_case!(matrix_string_length_stable, |left, right| {
    format!(
        r#"long left = {left}; long right = {right}; string text = "x" + left.ToString() + "x"; Console.WriteLine(text.Length > 1 && right >= 0);"#
    )
});

matrix_case!(matrix_string_starts_with, |left, right| {
    format!(
        "long left = {left}; long right = {right}; string text = left.ToString(); Console.WriteLine(text.StartsWith(text.Substring(0, 1)) && right >= 0);"
    )
});

matrix_case!(matrix_string_ends_with, |left, right| {
    format!(
        "long left = {left}; long right = {right}; string text = left.ToString(); Console.WriteLine(text.EndsWith(text.Substring(text.Length - 1, 1)) && right >= 0);"
    )
});

matrix_case!(matrix_string_replace_self, |left, right| {
    format!(
        "long left = {left}; long right = {right}; string text = left.ToString() + right.ToString(); string fixed = text.Replace(text, text); Console.WriteLine(fixed == text);"
    )
});

matrix_case!(matrix_string_trim_roundtrip, |left, right| {
    format!(
        r#"long left = {left}; long right = {right}; string text = "  " + left.ToString() + "  "; Console.WriteLine(text.Trim() == left.ToString() && right >= 0);"#
    )
});

matrix_case!(matrix_string_upper_identity, |left, right| {
    format!(
        "long left = {left}; long right = {right}; string text = left.ToString(); Console.WriteLine(text.ToUpper() == text && right >= 0);"
    )
});

matrix_case!(matrix_string_lower_identity, |left, right| {
    format!(
        "long left = {left}; long right = {right}; string text = left.ToString(); Console.WriteLine(text.ToLower() == text && right >= 0);"
    )
});

matrix_case!(matrix_string_join_contains, |left, right| {
    format!(
        r#"long left = {left}; long right = {right}; string text = string.Join(",", left.ToString(), right.ToString()); Console.WriteLine(text.Contains(",") && right >= 0);"#
    )
});

matrix_case!(matrix_string_index_of_comma, |left, right| {
    format!(
        r#"long left = {left}; long right = {right}; string text = left.ToString() + "," + right.ToString(); Console.WriteLine(text.IndexOf(",") >= 0);"#
    )
});

matrix_case!(matrix_string_substring_length, |left, right| {
    format!(
        "long left = {left}; long right = {right}; string text = left.ToString() + right.ToString(); Console.WriteLine(text.Substring(0, 1).Length == 1 && right >= 0);"
    )
});

matrix_case!(matrix_string_compare_to_self, |left, right| {
    format!(
        "long left = {left}; long right = {right}; string text = left.ToString(); Console.WriteLine(text.CompareTo(text) == 0 && right >= 0);"
    )
});

matrix_case!(matrix_stringbuilder_length, |left, right| {
    format!(
        "long left = {left}; long right = {right}; System.Text.StringBuilder builder = new System.Text.StringBuilder(); builder.Append(left.ToString()); builder.Append(right.ToString()); Console.WriteLine(builder.Length > 0);"
    )
});

matrix_case!(matrix_array_from_literals_count, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var array = new long[2]; array[0] = left; array[1] = right; Console.WriteLine(array.Length == 2);"
    )
});

matrix_case!(matrix_array_index_sum, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var array = new long[2]; array[0] = left; array[1] = right; Console.WriteLine(array[0] + array[1] == left + right);"
    )
});

matrix_case!(matrix_array_access_roundtrip, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var array = new long[2]; array[0] = left; array[1] = right; Console.WriteLine(array[array.Length - 1] == right);"
    )
});

matrix_case!(matrix_array_first_write, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var array = new long[1]; array[0] = left; array[0] = array[0] + right - right; Console.WriteLine(array[0] == left);"
    )
});

matrix_case!(matrix_array_clone_length, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var array = new long[2]; array[0] = left; array[1] = right; var clone = (long[])array.Clone(); Console.WriteLine(clone.Length == array.Length);"
    )
});

matrix_case!(matrix_list_add_remove_cycle, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var list = new System.Collections.Generic.List<long>(); list.Add(left); list.Add(right); list.Remove(left); Console.WriteLine(list.Count == 1);"
    )
});

matrix_case!(matrix_list_contains_added, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var list = new System.Collections.Generic.List<long>(); list.Add(left); Console.WriteLine(list.Contains(left) && right >= 0);"
    )
});

matrix_case!(matrix_list_insert_count, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var list = new System.Collections.Generic.List<long>(); list.Add(left); list.Insert(0, right); Console.WriteLine(list.Count == 2 && right >= 0);"
    )
});

matrix_case!(matrix_list_clear_zero, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var list = new System.Collections.Generic.List<long>(); list.Add(left); list.Clear(); Console.WriteLine(list.Count == 0);"
    )
});

matrix_case!(matrix_list_index_of_right, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var list = new System.Collections.Generic.List<long>(); list.Add(left); list.Add(right); Console.WriteLine(list.IndexOf(right) >= 0);"
    )
});

matrix_case!(matrix_dict_add_lookup, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var map = new System.Collections.Generic.Dictionary<long, long>(); map[left] = right; Console.WriteLine(map.ContainsKey(left));"
    )
});

matrix_case!(matrix_dict_value_roundtrip, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var map = new System.Collections.Generic.Dictionary<long, long>(); map[left] = left + right; Console.WriteLine(map[left] == left + right);"
    )
});

matrix_case!(matrix_dict_tryget, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var map = new System.Collections.Generic.Dictionary<long, long>(); map[left] = left + right; long found = -1; Console.WriteLine(map.TryGetValue(left, out found) && found == left + right);"
    )
});

matrix_case!(matrix_set_add_count, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var set = new System.Collections.Generic.HashSet<long>(); set.Add(left); set.Add(right); Console.WriteLine(set.Count >= 1 && right >= 0);"
    )
});

matrix_case!(matrix_set_remove_non_negative, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var set = new System.Collections.Generic.HashSet<long>(); set.Add(left); set.Add(right); set.Remove(right); Console.WriteLine(set.Count <= 2);"
    )
});

matrix_case!(matrix_set_contains_added, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var set = new System.Collections.Generic.HashSet<long>(); set.Add(left); Console.WriteLine(set.Contains(left) && right >= 0);"
    )
});

matrix_case!(matrix_tuple_item_one, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var pair = System.ValueTuple.Create(left, right); Console.WriteLine(pair.Item1 == left && right >= 0);"
    )
});

matrix_case!(matrix_tuple_item_two, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var pair = System.ValueTuple.Create(left, right); Console.WriteLine(pair.Item2 == right && right >= 0);"
    )
});

matrix_case!(matrix_tuple_equals_clone, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var pair = System.ValueTuple.Create(left, right); Console.WriteLine(System.ValueTuple.Create(left, right).Equals(pair) && right >= 0);"
    )
});

matrix_case!(matrix_tuple_nested_item, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var nested = System.ValueTuple.Create(left, System.ValueTuple.Create(right, right)); Console.WriteLine(nested.Item2.Item1 == right && right >= 0);"
    )
});

matrix_case!(matrix_cast_long_to_int_roundtrip, |left, right| {
    format!(
        "long left = {left}; long right = {right}; int narrowed = (int)left; long widened = (long)narrowed; Console.WriteLine(widened == left && right >= 0);"
    )
});

matrix_case!(matrix_cast_int_to_long_expand, |left, right| {
    format!(
        "long left = {left}; long right = {right}; int narrowed = (int)left; long widened = (long)narrowed; Console.WriteLine(widened == narrowed && right >= 0);"
    )
});

matrix_case!(matrix_cast_to_string_and_parse, |left, right| {
    format!(
        "long left = {left}; long right = {right}; long parsed = long.Parse(left.ToString()); Console.WriteLine(parsed == left && right >= 0);"
    )
});

matrix_case!(matrix_parse_double_safety, |left, right| {
    format!(
        "long left = {left}; long right = {right}; double value = (double)left; Console.WriteLine(value >= 0 || value < 0);"
    )
});

matrix_case!(matrix_nullable_has_value, |left, right| {
    format!(
        "long left = {left}; long right = {right}; long? value = left; Console.WriteLine(value.HasValue && value.GetValueOrDefault() == left && right >= 0);"
    )
});

matrix_case!(matrix_nullable_coalesce_self, |left, right| {
    format!(
        "long left = {left}; long right = {right}; long? value = left; Console.WriteLine((value ?? (long)0) == left && right >= 0);"
    )
});

matrix_case!(matrix_nullable_default_fallback, |left, right| {
    format!(
        "long left = {left}; long right = {right}; long? value = left > 0 ? left : (long?)null; Console.WriteLine((value ?? 0L) == left && left > 0 && right >= 0);"
    )
});

matrix_case!(matrix_nullable_conditional_member, |left, right| {
    format!(
        "long left = {left}; long right = {right}; long? value = left; Console.WriteLine((value?.ToString() == left.ToString()) == true);"
    )
});

matrix_case!(matrix_nullable_value_or_default, |left, right| {
    format!(
        "long left = {left}; long right = {right}; long? value = left; Console.WriteLine(value.GetValueOrDefault() == left && right >= 0);"
    )
});

matrix_case!(matrix_generic_list_simple, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var list = new System.Collections.Generic.List<long>(); list.Add(left); list.Add(right); Console.WriteLine(list.Count == 2);"
    )
});

matrix_case!(matrix_generic_dict_simple, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var map = new System.Collections.Generic.Dictionary<long, long>(); map[left] = right; Console.WriteLine(map.ContainsKey(left) && map[left] == right);"
    )
});

matrix_case!(matrix_boxed_identity, |left, right| {
    format!(
        "long left = {left}; long right = {right}; object value = left; Console.WriteLine(value != null && right >= 0);"
    )
});

matrix_case!(matrix_boxed_unbox, |left, right| {
    format!(
        "long left = {left}; long right = {right}; object value = left; Console.WriteLine((long)value == left && right >= 0);"
    )
});

matrix_case!(matrix_random_next_bound, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var rand = new System.Random((int)left); int value = rand.Next((int)right, (int)right + 2); Console.WriteLine(value >= right && value < right + 2);"
    )
});

matrix_case!(matrix_random_double_range, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var rand = new System.Random((int)left); double value = rand.NextDouble(); Console.WriteLine(value >= 0.0 && value < 1.0);"
    )
});

matrix_case!(matrix_object_to_string, |left, right| {
    format!(
        "long left = {left}; long right = {right}; object value = left; Console.WriteLine(value.ToString() == left.ToString() && right >= 0);"
    )
});

matrix_case!(matrix_pattern_type_match, |left, right| {
    format!(
        "long left = {left}; long right = {right}; object value = left; Console.WriteLine(value is long && right >= 0);"
    )
});

matrix_case!(matrix_typeof_similar, |left, right| {
    format!(
        r#"long left = {left}; long right = {right}; object value = left; Console.WriteLine((value.GetType().Name == "Int64") && right >= 0);"#
    )
});

matrix_case!(matrix_enum_identity, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(System.DayOfWeek.Monday == System.DayOfWeek.Monday && right >= 0);"
    )
});

matrix_case!(matrix_enum_parse_to_string, |left, right| {
    format!(
        "long left = {left}; long right = {right}; Console.WriteLine(System.DayOfWeek.Saturday.ToString().Length > 0 && right >= 0);"
    )
});

matrix_case!(matrix_datetime_year_guard, |left, right| {
    format!(
        "long left = {left}; long right = {right}; long year = System.DateTime.UtcNow.Year; Console.WriteLine(year > 0 && right >= 0);"
    )
});

matrix_case!(matrix_datetime_parse_fixed, |left, right| {
    format!(
        r#"long left = {left}; long right = {right}; var parsed = System.DateTime.Parse("2020-01-01"); Console.WriteLine(parsed.Year == 2020 && right >= 0);"#
    )
});

matrix_case!(matrix_timespan_from_minutes, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var span = System.TimeSpan.FromMinutes(left); Console.WriteLine(span.TotalMinutes >= left && right >= 0);"
    )
});

matrix_case!(matrix_timespan_ticks_positive, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var span = System.TimeSpan.FromSeconds(left); Console.WriteLine(span.Ticks > 0 && right >= 0);"
    )
});

matrix_case!(matrix_task_from_result, |left, right| {
    format!(
        "long left = {left}; long right = {right}; var t = System.Threading.Tasks.Task.FromResult(left); Console.WriteLine(t.Result == left && right >= 0);"
    )
});
