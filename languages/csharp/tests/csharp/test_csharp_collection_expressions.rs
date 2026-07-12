//! Collection expressions `[..]` and spread `..items` (C# 12).

csharp_cases! {
    collection_expression_list_int_count => {
        r#"System.Collections.Generic.List<int> list = [1, 2, 3];
Console.WriteLine(list.Count);"#,
        ["3"]
    };

    collection_expression_list_int_middle_element => {
        r#"System.Collections.Generic.List<int> list = [10, 20, 30];
Console.WriteLine(list[1]);"#,
        ["20"]
    };

    collection_expression_array_length => {
        r#"int[] arr = [10, 20, 30];
Console.WriteLine(arr.Length);"#,
        ["3"]
    };

    collection_expression_array_first_element => {
        r#"int[] arr = [7, 8, 9];
Console.WriteLine(arr[0]);"#,
        ["7"]
    };

    collection_expression_array_last_element => {
        r#"int[] arr = [7, 8, 9];
Console.WriteLine(arr[2]);"#,
        ["9"]
    };

    collection_expression_spread_merges_two_arrays => {
        r#"int[] a = [1, 2, 3];
int[] b = [4, 5, 6];
int[] c = [..a, ..b];
Console.WriteLine(c.Length); Console.WriteLine(c[3]);"#,
        ["6", "4"]
    };

    collection_expression_spread_preserves_order => {
        r#"int[] a = [1, 2];
int[] b = [3];
int[] c = [..a, ..b];
Console.WriteLine(c[0]); Console.WriteLine(c[2]);"#,
        ["1", "3"]
    };

    collection_expression_empty_array_zero_length => {
        r#"int[] empty = [];
Console.WriteLine(empty.Length);"#,
        ["0"]
    };

    collection_expression_span_int_length => {
        r#"System.Span<int> s = [1, 2, 3];
Console.WriteLine(s.Length);"#,
        ["3"]
    };

    collection_expression_span_first_element => {
        r#"System.Span<int> s = [5, 6];
Console.WriteLine(s[0]);"#,
        ["5"]
    };

    collection_expression_spread_single_element_array => {
        r#"int[] one = [42];
int[] two = [..one, 99];
Console.WriteLine(two.Length); Console.WriteLine(two[1]);"#,
        ["2", "99"]
    };

    collection_expression_spread_at_start => {
        r#"int[] tail = [3, 4];
int[] all = [..tail, 1, 2];
Console.WriteLine(all[0]); Console.WriteLine(all[2]);"#,
        ["3", "1"]
    };

    collection_expression_spread_at_end => {
        r#"int[] head = [1, 2];
int[] all = [9, ..head];
Console.WriteLine(all[0]); Console.WriteLine(all[2]);"#,
        ["9", "2"]
    };

    collection_expression_spread_in_middle => {
        r#"int[] mid = [2, 3];
int[] all = [1, ..mid, 4];
Console.WriteLine(all[1]); Console.WriteLine(all[3]);"#,
        ["2", "4"]
    };

    collection_expression_triple_spread_merge => {
        r#"int[] a = [1]; int[] b = [2]; int[] c = [3];
int[] all = [..a, ..b, ..c];
Console.WriteLine(all.Length); Console.WriteLine(all[2]);"#,
        ["3", "3"]
    };

    collection_expression_literal_then_spread => {
        r#"int[] rest = [2, 3];
int[] all = [1, ..rest];
Console.WriteLine(all.Length);"#,
        ["3"]
    };

    collection_expression_string_array_elements => {
        r#"string[] words = ["a", "b", "c"];
Console.WriteLine(words[1]);"#,
        ["b"]
    };

    collection_expression_char_array_elements => {
        r#"char[] chars = ['x', 'y'];
Console.WriteLine(chars[0]);"#,
        ["x"]
    };

    collection_expression_double_array_length => {
        r#"double[] vals = [1.5, 2.5];
Console.WriteLine(vals.Length);"#,
        ["2"]
    };

    collection_expression_bool_array_values => {
        r#"bool[] flags = [true, false, true];
Console.WriteLine(flags[0]); Console.WriteLine(flags[1]);"#,
        ["True", "False"]
    };

    collection_expression_long_array_sum_loop => {
        r#"long[] nums = [10000000000L, 20000000000L];
long total = 0;
foreach (var n in nums) total += n;
Console.WriteLine(total);"#,
        ["30000000000"]
    };

    collection_expression_byte_array_length => {
        r#"byte[] data = [1, 2, 3, 4];
Console.WriteLine(data.Length);"#,
        ["4"]
    };

    collection_expression_single_element => {
        r#"int[] one = [99];
Console.WriteLine(one[0]);"#,
        ["99"]
    };

    collection_expression_two_elements => {
        r#"int[] pair = [4, 5];
Console.WriteLine(pair[0] + pair[1]);"#,
        ["9"]
    };

    collection_expression_with_zero_values => {
        r#"int[] zeros = [0, 0, 0];
Console.WriteLine(zeros[1]);"#,
        ["0"]
    };

    collection_expression_with_negative_numbers => {
        r#"int[] nums = [-1, -2];
Console.WriteLine(nums[0] + nums[1]);"#,
        ["-3"]
    };

    collection_expression_with_expression_elements => {
        r#"int[] nums = [1 + 1, 2 + 2, 3 + 3];
Console.WriteLine(nums[2]);"#,
        ["6"]
    };

    collection_expression_spread_empty_array_adds_nothing => {
        r#"int[] empty = [];
int[] all = [1, ..empty, 2];
Console.WriteLine(all.Length);"#,
        ["2"]
    };

    collection_expression_spread_copy_via_self => {
        r#"int[] src = [1, 2];
int[] copy = [..src];
Console.WriteLine(copy[1]);"#,
        ["2"]
    };

    collection_expression_list_from_spread_arrays => {
        r#"int[] a = [1, 2];
int[] b = [3];
System.Collections.Generic.List<int> list = [..a, ..b];
Console.WriteLine(list.Count); Console.WriteLine(list[2]);"#,
        ["3", "3"]
    };

    collection_expression_foreach_prints_count => {
        r#"int[] arr = [1, 2, 3];
int count = 0;
foreach (var _ in arr) count++;
Console.WriteLine(count);"#,
        ["3"]
    };

    collection_expression_index_access_second => {
        r#"int[] arr = [10, 11, 12];
Console.WriteLine(arr[1]);"#,
        ["11"]
    };

    collection_expression_target_typed_list_add_after => {
        r#"System.Collections.Generic.List<int> list = [1, 2];
list.Add(3);
Console.WriteLine(list[2]);"#,
        ["3"]
    };

    collection_expression_duplicate_literals_allowed => {
        r#"int[] arr = [7, 7, 7];
Console.WriteLine(arr[2]);"#,
        ["7"]
    };

    collection_expression_in_method_argument => {
        r#"int Sum(int[] data) { int t = 0; foreach (var n in data) t += n; return t; }
Console.WriteLine(Sum([1, 2, 3]));"#,
        ["6"]
    };

    collection_expression_returned_from_method => {
        r#"int[] Make() => [4, 5, 6];
Console.WriteLine(Make()[1]);"#,
        ["5"]
    };

    collection_expression_nested_spread_flatten => {
        r#"int[] a = [1]; int[] b = [2]; int[] c = [3];
int[] all = [..a, ..b, ..c];
Console.WriteLine(string.Join(",", all));"#,
        ["1,2,3"]
    };

    collection_expression_spread_different_lengths => {
        r#"int[] small = [1];
int[] big = [2, 3, 4, 5];
int[] all = [..small, ..big];
Console.WriteLine(all.Length);"#,
        ["5"]
    };

    collection_expression_modifying_copy_not_source => {
        r#"int[] src = [1, 2];
int[] copy = [..src];
copy[0] = 9;
Console.WriteLine(src[0]); Console.WriteLine(copy[0]);"#,
        ["1", "9"]
    };

    collection_expression_string_join => {
        r#"string[] parts = ["a", "b", "c"];
Console.WriteLine(string.Join("-", parts));"#,
        ["a-b-c"]
    };

    collection_expression_linq_count => {
        r#"int[] arr = [1, 2, 3, 4];
Console.WriteLine(System.Linq.Enumerable.Count(arr));"#,
        ["4"]
    };

    collection_expression_short_array => {
        r#"short[] vals = [10, 20];
Console.WriteLine(vals[1]);"#,
        ["20"]
    };

    collection_expression_float_array => {
        r#"float[] vals = [1.1f, 2.2f];
Console.WriteLine(vals.Length);"#,
        ["2"]
    };

    collection_expression_decimal_array => {
        r#"decimal[] vals = [1.5m, 2.5m];
Console.WriteLine(vals[0] + vals[1]);"#,
        ["4.0"]
    };

    collection_expression_spread_list_into_array => {
        r#"System.Collections.Generic.List<int> list = new() { 1, 2 };
int[] arr = [..list, 3];
Console.WriteLine(arr[2]);"#,
        ["3"]
    };

    collection_expression_multiple_spreads_with_literals => {
        r#"int[] a = [1, 2]; int[] b = [3];
int[] c = [0, ..a, ..b, 4];
Console.WriteLine(c[0]); Console.WriteLine(c[4]);"#,
        ["0", "4"]
    };

    collection_expression_readonly_span_length => {
        r#"System.ReadOnlySpan<int> s = [9, 8, 7];
Console.WriteLine(s.Length);"#,
        ["3"]
    };

    collection_expression_array_is_array_type => {
        r#"int[] arr = [1, 2];
Console.WriteLine(arr.GetType().IsArray);"#,
        ["True"]
    };

    collection_expression_list_first_element => {
        r#"System.Collections.Generic.List<string> list = ["x", "y"];
Console.WriteLine(list[0]);"#,
        ["x"]
    };

    collection_expression_spread_preserves_source_after_merge => {
        r#"int[] a = [1, 2];
int[] b = [..a, 3];
Console.WriteLine(a.Length); Console.WriteLine(b.Length);"#,
        ["2", "3"]
    };

    collection_expression_large_count_via_loop => {
        r#"int[] arr = [1, 2, 3, 4, 5];
int sum = 0;
for (int i = 0; i < arr.Length; i++) sum += arr[i];
Console.WriteLine(sum);"#,
        ["15"]
    };

    collection_expression_mixed_spread_and_literal_sum => {
        r#"int[] mid = [2, 3];
int[] all = [1, ..mid, 4];
Console.WriteLine(all[0] + all[3]);"#,
        ["5"]
    };

    collection_expression_empty_list_via_target_type => {
        r#"System.Collections.Generic.List<int> list = [];
Console.WriteLine(list.Count);"#,
        ["0"]
    };

    collection_expression_spread_into_new_list_count => {
        r#"int[] data = [5, 6, 7];
System.Collections.Generic.List<int> list = [..data];
Console.WriteLine(list[2]);"#,
        ["7"]
    };

    collection_expression_int_array_from_nested_addition => {
        r#"int[] arr = [1 + 2, 3 + 4];
Console.WriteLine(arr[1]);"#,
        ["7"]
    };
}
