use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(string_comparer_ordinal_ignore_case_treats_text_as_equal, r#"Console.WriteLine(System.StringComparer.OrdinalIgnoreCase.Equals("abc", "ABC"));"#, ["True"]);
csharp_case!(string_comparer_ordinal_reports_negative_for_smaller_text, r#"Console.WriteLine(System.StringComparer.Ordinal.Compare("a", "b"));"#, ["-1"]);
csharp_case!(default_comparer_orders_smaller_int_before_larger_int, r#"Console.WriteLine(System.Collections.Generic.Comparer<int>.Default.Compare(2, 5));"#, ["-1"]);
csharp_case!(default_equality_comparer_reports_equal_values, r#"Console.WriteLine(System.Collections.Generic.EqualityComparer<int>.Default.Equals(4, 4));"#, ["True"]);
csharp_case!(list_sort_orders_strings_alphabetically, r#"using System.Collections.Generic; var list = new List<string> { "c", "a", "b" }; list.Sort(); foreach (var value in list) Console.WriteLine(value);"#, ["a", "b", "c"]);
csharp_case!(list_sort_with_custom_comparison_orders_descending, r#"using System.Collections.Generic; var list = new List<int> { 1, 3, 2 }; list.Sort((left, right) => right.CompareTo(left)); foreach (var value in list) Console.WriteLine(value);"#, ["3", "2", "1"]);
csharp_case!(array_sort_with_parallel_keys_reorders_items_by_key, r#"var keys = new[] { 2, 1 }; var items = new[] { "b", "a" }; System.Array.Sort(keys, items); foreach (var value in items) Console.WriteLine(value);"#, ["a", "b"]);
csharp_case!(sequence_equal_reports_true_for_matching_arrays, r#"using System.Linq; Console.WriteLine(new[] { 1, 2 }.SequenceEqual(new[] { 1, 2 }));"#, ["True"]);
csharp_case!(sequence_equal_reports_false_for_different_arrays, r#"using System.Linq; Console.WriteLine(new[] { 1, 2 }.SequenceEqual(new[] { 2, 1 }));"#, ["False"]);
csharp_case!(hashset_with_case_insensitive_comparer_merges_text_variants, r#"using System.Collections.Generic; var set = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase); set.Add("A"); set.Add("a"); Console.WriteLine(set.Count);"#, ["1"]);
csharp_case!(dictionary_with_case_insensitive_comparer_finds_key_variant, r#"using System.Collections.Generic; var map = new Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase) { ["Key"] = 3 }; Console.WriteLine(map.ContainsKey("key"));"#, ["True"]);
csharp_case!(icomparable_implementation_drives_default_sort_order, r#"using System.Collections.Generic; class Rank : System.IComparable<Rank> { public int Value; public int CompareTo(Rank other) { return Value.CompareTo(other.Value); } } var list = new List<Rank> { new Rank { Value = 3 }, new Rank { Value = 1 } }; list.Sort(); foreach (var item in list) Console.WriteLine(item.Value);"#, ["1", "3"]);
csharp_case!(icomparable_direct_compareto_invocation_returns_positive_value, r#"class Rank : System.IComparable<Rank> { public int Value; public int CompareTo(Rank other) { return Value.CompareTo(other.Value); } } var left = new Rank(); left.Value = 3; var right = new Rank(); right.Value = 1; Console.WriteLine(left.CompareTo(right));"#, ["1"]);
csharp_case!(icomparable_compareto_after_list_indexing_returns_positive_value, r#"using System.Collections.Generic; class Rank : System.IComparable<Rank> { public int Value; public Rank(int value) { Value = value; } public int CompareTo(Rank other) { return Value.CompareTo(other.Value); } } var list = new List<Rank> { new Rank(3), new Rank(1) }; Console.WriteLine(list[0].CompareTo(list[1]));"#, ["1"]);
csharp_case!(compareto_on_string_reports_zero_for_same_text, r#"Console.WriteLine("abc".CompareTo("abc"));"#, ["0"]);
csharp_case!(compareto_on_integer_reports_positive_for_larger_value, r#"Console.WriteLine(9.CompareTo(3));"#, ["1"]);
csharp_case!(order_by_with_key_projection_sorts_by_length, r#"using System.Linq; var values = new[] { "bbb", "a", "cc" }.OrderBy(text => text.Length); foreach (var value in values) Console.WriteLine(value);"#, ["a", "cc", "bbb"]);
csharp_case!(then_by_breaks_ties_after_primary_sort_key, r#"using System.Linq; var values = new[] { "ba", "aa", "c" }.OrderBy(text => text.Length).ThenBy(text => text); foreach (var value in values) Console.WriteLine(value);"#, ["c", "aa", "ba"]);
csharp_case!(max_returns_largest_value_from_sequence, r#"using System.Linq; Console.WriteLine(new[] { 2, 9, 4 }.Max());"#, ["9"]);
csharp_case!(min_returns_smallest_value_from_sequence, r#"using System.Linq; Console.WriteLine(new[] { 2, 9, 4 }.Min());"#, ["2"]);
csharp_case!(string_compare_with_ignore_case_reports_equality, r#"Console.WriteLine(string.Compare("abc", "ABC", true));"#, ["0"]);
csharp_case!(array_sort_with_string_comparer_can_ignore_case, r#"var values = new[] { "b", "A", "c" }; System.Array.Sort(values, System.StringComparer.OrdinalIgnoreCase); foreach (var value in values) Console.WriteLine(value);"#, ["A", "b", "c"]);