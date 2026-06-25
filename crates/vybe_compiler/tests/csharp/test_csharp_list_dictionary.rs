use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

// ═══════════════════════════════════════════════════════════
// List<T> — construction, mutation, lookup, ordering, iteration
// ═══════════════════════════════════════════════════════════

csharp_case!(
    list_empty_initial_count_is_zero,
    r#"using System.Collections.Generic; var list = new List<int>(); Console.WriteLine(list.Count);"#,
    ["0"]
);
csharp_case!(
    list_add_single_integer_sets_count_one,
    r#"using System.Collections.Generic; var list = new List<int>(); list.Add(42); Console.WriteLine(list.Count);"#,
    ["1"]
);
csharp_case!(
    list_add_two_integers_preserves_insertion_order,
    r#"using System.Collections.Generic; var list = new List<int>(); list.Add(10); list.Add(20); Console.WriteLine(list[0]); Console.WriteLine(list[1]);"#,
    ["10", "20"]
);
csharp_case!(
    list_add_string_appends_text_element,
    r#"using System.Collections.Generic; var list = new List<string>(); list.Add("hello"); Console.WriteLine(list[0]);"#,
    ["hello"]
);
csharp_case!(
    list_add_bool_stores_true_literal,
    r#"using System.Collections.Generic; var list = new List<bool>(); list.Add(true); Console.WriteLine(list[0]);"#,
    ["True"]
);
csharp_case!(
    list_add_double_stores_fractional_number,
    r#"using System.Collections.Generic; var list = new List<double>(); list.Add(2.5); Console.WriteLine(list[0]);"#,
    ["2.5"]
);
csharp_case!(
    list_add_zero_as_valid_element,
    r#"using System.Collections.Generic; var list = new List<int>(); list.Add(0); Console.WriteLine(list[0]); Console.WriteLine(list.Contains(0));"#,
    ["0", "True"]
);
csharp_case!(
    list_indexer_get_reads_first_added_item,
    r#"using System.Collections.Generic; var list = new List<int>(); list.Add(5); list.Add(6); Console.WriteLine(list[0]);"#,
    ["5"]
);
csharp_case!(
    list_indexer_get_reads_last_added_item,
    r#"using System.Collections.Generic; var list = new List<int>(); list.Add(5); list.Add(6); Console.WriteLine(list[1]);"#,
    ["6"]
);
csharp_case!(
    list_indexer_set_overwrites_middle_element,
    r#"using System.Collections.Generic; var list = new List<int> { 1, 2, 3 }; list[1] = 99; Console.WriteLine(list[1]);"#,
    ["99"]
);
csharp_case!(
    list_indexer_set_replaces_first_element,
    r#"using System.Collections.Generic; var list = new List<string> { "old", "keep" }; list[0] = "new"; Console.WriteLine(list[0]);"#,
    ["new"]
);
csharp_case!(
    list_insert_at_zero_prefixes_new_item,
    r#"using System.Collections.Generic; var list = new List<int> { 2, 3 }; list.Insert(0, 1); foreach (var x in list) Console.WriteLine(x);"#,
    ["1", "2", "3"]
);
csharp_case!(
    list_insert_at_middle_splits_sequence,
    r#"using System.Collections.Generic; var list = new List<string> { "a", "c" }; list.Insert(1, "b"); foreach (var s in list) Console.WriteLine(s);"#,
    ["a", "b", "c"]
);
csharp_case!(
    list_insert_at_end_appends_like_add,
    r#"using System.Collections.Generic; var list = new List<int> { 1, 2 }; list.Insert(2, 3); Console.WriteLine(list[2]);"#,
    ["3"]
);
csharp_case!(
    list_insert_increases_count_by_one,
    r#"using System.Collections.Generic; var list = new List<int> { 1 }; list.Insert(0, 0); Console.WriteLine(list.Count);"#,
    ["2"]
);
csharp_case!(
    list_remove_existing_integer_returns_true,
    r#"using System.Collections.Generic; var list = new List<int> { 1, 2, 3 }; Console.WriteLine(list.Remove(2));"#,
    ["True"]
);
csharp_case!(
    list_remove_missing_integer_returns_false,
    r#"using System.Collections.Generic; var list = new List<int> { 1, 2 }; Console.WriteLine(list.Remove(9));"#,
    ["False"]
);
csharp_case!(
    list_remove_existing_string_element,
    r#"using System.Collections.Generic; var list = new List<string> { "a", "b" }; list.Remove("a"); Console.WriteLine(list[0]);"#,
    ["b"]
);
csharp_case!(
    list_remove_shrinks_count_by_one,
    r#"using System.Collections.Generic; var list = new List<int> { 1, 2, 3 }; list.Remove(2); Console.WriteLine(list.Count);"#,
    ["2"]
);
csharp_case!(
    list_removeat_zero_drops_first_item,
    r#"using System.Collections.Generic; var list = new List<int> { 10, 20, 30 }; list.RemoveAt(0); Console.WriteLine(list[0]);"#,
    ["20"]
);
csharp_case!(
    list_removeat_middle_drops_center_item,
    r#"using System.Collections.Generic; var list = new List<int> { 10, 20, 30 }; list.RemoveAt(1); foreach (var x in list) Console.WriteLine(x);"#,
    ["10", "30"]
);
csharp_case!(
    list_removeat_last_drops_final_item,
    r#"using System.Collections.Generic; var list = new List<int> { 10, 20, 30 }; list.RemoveAt(2); Console.WriteLine(list.Count);"#,
    ["2"]
);
csharp_case!(
    list_removeat_shrinks_count_by_one,
    r#"using System.Collections.Generic; var list = new List<int> { 1, 2 }; list.RemoveAt(0); Console.WriteLine(list.Count);"#,
    ["1"]
);
csharp_case!(
    list_clear_empties_all_elements,
    r#"using System.Collections.Generic; var list = new List<int> { 1, 2, 3 }; list.Clear(); Console.WriteLine(list.Count);"#,
    ["0"]
);
csharp_case!(
    list_clear_allows_subsequent_add,
    r#"using System.Collections.Generic; var list = new List<int> { 1 }; list.Clear(); list.Add(7); Console.WriteLine(list[0]);"#,
    ["7"]
);
csharp_case!(
    list_contains_locates_present_integer,
    r#"using System.Collections.Generic; var list = new List<int> { 3, 4 }; Console.WriteLine(list.Contains(4));"#,
    ["True"]
);
csharp_case!(
    list_contains_reports_absent_integer,
    r#"using System.Collections.Generic; var list = new List<int> { 3, 4 }; Console.WriteLine(list.Contains(9));"#,
    ["False"]
);
csharp_case!(
    list_contains_locates_present_string,
    r#"using System.Collections.Generic; var list = new List<string> { "cat" }; Console.WriteLine(list.Contains("cat"));"#,
    ["True"]
);
csharp_case!(
    list_contains_false_after_successful_remove,
    r#"using System.Collections.Generic; var list = new List<int> { 5 }; list.Remove(5); Console.WriteLine(list.Contains(5));"#,
    ["False"]
);
csharp_case!(
    list_indexof_finds_first_matching_index,
    r#"using System.Collections.Generic; var list = new List<int> { 10, 20, 30 }; Console.WriteLine(list.IndexOf(20));"#,
    ["1"]
);
csharp_case!(
    list_indexof_returns_negative_one_when_absent,
    r#"using System.Collections.Generic; var list = new List<int> { 1, 2 }; Console.WriteLine(list.IndexOf(99));"#,
    ["-1"]
);
csharp_case!(
    list_indexof_selects_first_duplicate_match,
    r#"using System.Collections.Generic; var list = new List<int> { 7, 3, 7 }; Console.WriteLine(list.IndexOf(7));"#,
    ["0"]
);
csharp_case!(
    list_indexof_zero_after_head_insert,
    r#"using System.Collections.Generic; var list = new List<int> { 2 }; list.Insert(0, 1); Console.WriteLine(list.IndexOf(1));"#,
    ["0"]
);
csharp_case!(
    list_sort_orders_numbers_ascending,
    r#"using System.Collections.Generic; var list = new List<int> { 3, 1, 2 }; list.Sort(); foreach (var x in list) Console.WriteLine(x);"#,
    ["1", "2", "3"]
);
csharp_case!(
    list_sort_orders_words_alphabetically,
    r#"using System.Collections.Generic; var list = new List<string> { "c", "a", "b" }; list.Sort(); foreach (var s in list) Console.WriteLine(s);"#,
    ["a", "b", "c"]
);
csharp_case!(
    list_sort_leaves_already_sorted_unchanged,
    r#"using System.Collections.Generic; var list = new List<int> { 1, 2, 3 }; list.Sort(); Console.WriteLine(list[0]); Console.WriteLine(list[2]);"#,
    ["1", "3"]
);
csharp_case!(
    list_indexof_after_sort_finds_reordered_item,
    r#"using System.Collections.Generic; var list = new List<int> { 3, 1, 2 }; list.Sort(); Console.WriteLine(list.IndexOf(2));"#,
    ["1"]
);
csharp_case!(
    list_reverse_inverts_element_order,
    r#"using System.Collections.Generic; var list = new List<int> { 1, 2, 3 }; list.Reverse(); foreach (var x in list) Console.WriteLine(x);"#,
    ["3", "2", "1"]
);
csharp_case!(
    list_reverse_leaves_singleton_unchanged,
    r#"using System.Collections.Generic; var list = new List<int> { 42 }; list.Reverse(); Console.WriteLine(list[0]);"#,
    ["42"]
);
csharp_case!(
    list_reverse_twice_restores_original_order,
    r#"using System.Collections.Generic; var list = new List<int> { 1, 2, 3 }; list.Reverse(); list.Reverse(); Console.WriteLine(list[0]); Console.WriteLine(list[2]);"#,
    ["1", "3"]
);
csharp_case!(
    list_foreach_sums_all_integers,
    r#"using System.Collections.Generic; var list = new List<int> { 1, 2, 3 }; int sum = 0; foreach (var x in list) sum += x; Console.WriteLine(sum);"#,
    ["6"]
);
csharp_case!(
    list_foreach_emits_each_string_line,
    r#"using System.Collections.Generic; var list = new List<string> { "x", "y" }; foreach (var s in list) Console.WriteLine(s);"#,
    ["x", "y"]
);
csharp_case!(
    list_foreach_over_empty_emits_nothing,
    r#"using System.Collections.Generic; var list = new List<int>(); foreach (var x in list) Console.WriteLine(x); Console.WriteLine("done");"#,
    ["done"]
);
csharp_case!(
    list_single_element_reports_count_one,
    r#"using System.Collections.Generic; var list = new List<int> { 99 }; Console.WriteLine(list.Count);"#,
    ["1"]
);
csharp_case!(
    list_single_element_indexer_returns_value,
    r#"using System.Collections.Generic; var list = new List<string> { "solo" }; Console.WriteLine(list[0]);"#,
    ["solo"]
);
csharp_case!(
    list_addrange_from_int_array_extends_tail,
    r#"using System.Collections.Generic; var list = new List<int> { 1, 2 }; list.AddRange(new int[] { 3, 4 }); Console.WriteLine(list[3]);"#,
    ["4"]
);
csharp_case!(
    list_addrange_from_sibling_list_merges_items,
    r#"using System.Collections.Generic; var list = new List<int> { 1 }; var extra = new List<int> { 2, 3 }; list.AddRange(extra); Console.WriteLine(list.Count);"#,
    ["3"]
);
csharp_case!(
    list_addrange_empty_collection_keeps_count,
    r#"using System.Collections.Generic; var list = new List<int> { 1 }; list.AddRange(new int[] { }); Console.WriteLine(list.Count);"#,
    ["1"]
);
csharp_case!(
    list_addrange_updates_count_by_batch_size,
    r#"using System.Collections.Generic; var list = new List<int>(); list.AddRange(new int[] { 4, 5, 6 }); Console.WriteLine(list.Count);"#,
    ["3"]
);
csharp_case!(
    list_initializer_syntax_populates_elements,
    r#"using System.Collections.Generic; var list = new List<int> { 8, 9 }; Console.WriteLine(list[1]);"#,
    ["9"]
);
csharp_case!(
    list_add_after_remove_restores_length,
    r#"using System.Collections.Generic; var list = new List<int> { 1, 2 }; list.Remove(1); list.Add(3); Console.WriteLine(list.Count);"#,
    ["2"]
);
csharp_case!(
    list_remove_sole_item_yields_empty_list,
    r#"using System.Collections.Generic; var list = new List<int> { 7 }; list.Remove(7); Console.WriteLine(list.Count);"#,
    ["0"]
);
csharp_case!(
    list_count_tracks_clear_then_three_adds,
    r#"using System.Collections.Generic; var list = new List<int> { 1, 2 }; list.Clear(); list.Add(1); list.Add(2); list.Add(3); Console.WriteLine(list.Count);"#,
    ["3"]
);
csharp_case!(
    list_nested_outer_list_holds_two_inner_lists,
    r#"using System.Collections.Generic; var outer = new List<List<int>>(); outer.Add(new List<int> { 1 }); outer.Add(new List<int> { 2, 3 }); Console.WriteLine(outer.Count);"#,
    ["2"]
);
csharp_case!(
    list_nested_indexer_reaches_inner_list,
    r#"using System.Collections.Generic; var outer = new List<List<int>> { new List<int> { 10, 20 } }; Console.WriteLine(outer[0][1]);"#,
    ["20"]
);
csharp_case!(
    list_nested_foreach_counts_each_inner_length,
    r#"using System.Collections.Generic; var outer = new List<List<int>> { new List<int> { 1, 2 }, new List<int> { 3 } }; foreach (var inner in outer) Console.WriteLine(inner.Count);"#,
    ["2", "1"]
);
csharp_case!(
    list_nested_three_deep_reaches_innermost_value,
    r#"using System.Collections.Generic; var outer = new List<List<List<int>>>(); var mid = new List<List<int>>(); var inner = new List<int> { 5 }; mid.Add(inner); outer.Add(mid); Console.WriteLine(outer[0][0][0]);"#,
    ["5"]
);

// ═══════════════════════════════════════════════════════════
// Dictionary<K,V> — add, lookup, mutation, iteration
// ═══════════════════════════════════════════════════════════

csharp_case!(
    dictionary_empty_initial_count_is_zero,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); Console.WriteLine(map.Count);"#,
    ["0"]
);
csharp_case!(
    dictionary_add_string_key_stores_integer,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); map.Add("one", 1); Console.WriteLine(map["one"]);"#,
    ["1"]
);
csharp_case!(
    dictionary_add_second_distinct_string_key,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); map.Add("a", 1); map.Add("b", 2); Console.WriteLine(map["b"]);"#,
    ["2"]
);
csharp_case!(
    dictionary_add_integer_key_stores_string,
    r#"using System.Collections.Generic; var map = new Dictionary<int, string>(); map.Add(1, "one"); Console.WriteLine(map[1]);"#,
    ["one"]
);
csharp_case!(
    dictionary_indexer_reads_string_key_entry,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["k"] = 4 }; Console.WriteLine(map["k"]);"#,
    ["4"]
);
csharp_case!(
    dictionary_indexer_reads_integer_key_entry,
    r#"using System.Collections.Generic; var map = new Dictionary<int, string> { [2] = "two" }; Console.WriteLine(map[2]);"#,
    ["two"]
);
csharp_case!(
    dictionary_indexer_inserts_new_string_key,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); map["new"] = 9; Console.WriteLine(map["new"]);"#,
    ["9"]
);
csharp_case!(
    dictionary_indexer_overwrites_existing_entry,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1 }; map["a"] = 5; Console.WriteLine(map["a"]);"#,
    ["5"]
);
csharp_case!(
    dictionary_containskey_true_for_existing_string_key,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["x"] = 3 }; Console.WriteLine(map.ContainsKey("x"));"#,
    ["True"]
);
csharp_case!(
    dictionary_containskey_false_for_missing_string_key,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["x"] = 3 }; Console.WriteLine(map.ContainsKey("z"));"#,
    ["False"]
);
csharp_case!(
    dictionary_containskey_then_indexer_reads_value,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 7 }; if (map.ContainsKey("a")) Console.WriteLine(map["a"]);"#,
    ["7"]
);
csharp_case!(
    dictionary_containskey_guard_skips_missing_key,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); if (map.ContainsKey("missing")) Console.WriteLine(map["missing"]); else Console.WriteLine("absent");"#,
    ["absent"]
);
csharp_case!(
    dictionary_remove_deletes_existing_string_key,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["b"] = 2 }; map.Remove("a"); Console.WriteLine(map.ContainsKey("a"));"#,
    ["False"]
);
csharp_case!(
    dictionary_remove_absent_key_returns_false,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1 }; Console.WriteLine(map.Remove("z"));"#,
    ["False"]
);
csharp_case!(
    dictionary_count_reflects_two_add_operations,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); map.Add("a", 1); map.Add("b", 2); Console.WriteLine(map.Count);"#,
    ["2"]
);
csharp_case!(
    dictionary_foreach_keys_yields_inserted_string_keys,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["b"] = 2 }; foreach (var key in map.Keys) Console.WriteLine(key);"#,
    ["a", "b"]
);
csharp_case!(
    dictionary_foreach_values_yields_stored_integers,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 10, ["b"] = 20 }; foreach (var val in map.Values) Console.WriteLine(val);"#,
    ["10", "20"]
);
csharp_case!(
    dictionary_string_keys_support_multi_character_names,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["alpha"] = 1, ["beta"] = 2 }; Console.WriteLine(map["alpha"]);"#,
    ["1"]
);
csharp_case!(
    dictionary_integer_keys_map_to_string_values,
    r#"using System.Collections.Generic; var map = new Dictionary<int, string> { [10] = "ten", [20] = "twenty" }; Console.WriteLine(map[20]);"#,
    ["twenty"]
);
csharp_case!(
    dictionary_indexer_addition_increments_count,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); map["k"] = 1; Console.WriteLine(map.Count);"#,
    ["1"]
);
csharp_case!(
    dictionary_overwrite_preserves_entry_count,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1 }; map["a"] = 9; Console.WriteLine(map.Count);"#,
    ["1"]
);
csharp_case!(
    dictionary_remove_reduces_count_by_one,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["b"] = 2 }; map.Remove("a"); Console.WriteLine(map.Count);"#,
    ["1"]
);
csharp_case!(
    dictionary_clear_resets_entry_count,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1 }; map.Clear(); Console.WriteLine(map.Count);"#,
    ["0"]
);
csharp_case!(
    dictionary_multiple_int_keys_hold_distinct_values,
    r#"using System.Collections.Generic; var map = new Dictionary<int, int> { [1] = 100, [2] = 200 }; Console.WriteLine(map[1]); Console.WriteLine(map[2]);"#,
    ["100", "200"]
);
csharp_case!(
    dictionary_string_key_can_be_read_twice,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["k"] = 3 }; Console.WriteLine(map["k"]); Console.WriteLine(map["k"]);"#,
    ["3", "3"]
);
csharp_case!(
    dictionary_indexer_updates_integer_key_value,
    r#"using System.Collections.Generic; var map = new Dictionary<int, string> { [1] = "old" }; map[1] = "new"; Console.WriteLine(map[1]);"#,
    ["new"]
);
csharp_case!(
    dictionary_remove_erases_integer_key_mapping,
    r#"using System.Collections.Generic; var map = new Dictionary<int, string> { [5] = "five" }; map.Remove(5); Console.WriteLine(map.ContainsKey(5));"#,
    ["False"]
);
csharp_case!(
    dictionary_foreach_values_reads_string_payloads,
    r#"using System.Collections.Generic; var map = new Dictionary<int, string> { [1] = "a", [2] = "b" }; foreach (var val in map.Values) Console.WriteLine(val);"#,
    ["a", "b"]
);
csharp_case!(
    dictionary_containskey_still_true_after_overwrite,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["k"] = 1 }; map["k"] = 2; Console.WriteLine(map.ContainsKey("k"));"#,
    ["True"]
);
csharp_case!(
    dictionary_keys_collection_reports_count,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["b"] = 2, ["c"] = 3 }; Console.WriteLine(map.Keys.Count);"#,
    ["3"]
);
csharp_case!(
    dictionary_values_collection_reports_count,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["b"] = 2 }; Console.WriteLine(map.Values.Count);"#,
    ["2"]
);
csharp_case!(
    dictionary_foreach_pairs_prints_key_colon_value,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["b"] = 2, ["a"] = 1 }; foreach (var pair in map) Console.WriteLine(pair.Key + ":" + pair.Value);"#,
    ["b:2", "a:1"]
);
csharp_case!(
    dictionary_int_key_overwrite_via_indexer,
    r#"using System.Collections.Generic; var map = new Dictionary<int, int> { [1] = 10 }; map[1] = 99; Console.WriteLine(map[1]);"#,
    ["99"]
);
csharp_case!(
    dictionary_empty_containskey_always_false,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); Console.WriteLine(map.ContainsKey("any"));"#,
    ["False"]
);
csharp_case!(
    dictionary_remove_then_readd_same_key,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["k"] = 1 }; map.Remove("k"); map["k"] = 2; Console.WriteLine(map["k"]);"#,
    ["2"]
);
csharp_case!(
    dictionary_string_keys_with_different_lengths,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["longer"] = 2 }; Console.WriteLine(map["longer"]);"#,
    ["2"]
);
csharp_case!(
    dictionary_int_keys_store_negative_numbers,
    r#"using System.Collections.Generic; var map = new Dictionary<int, int> { [-1] = 100 }; Console.WriteLine(map[-1]);"#,
    ["100"]
);
csharp_case!(
    dictionary_add_via_add_method_increases_count,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); map.Add("solo", 1); Console.WriteLine(map.Count);"#,
    ["1"]
);
csharp_case!(
    dictionary_containskey_after_clear_is_false,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1 }; map.Clear(); Console.WriteLine(map.ContainsKey("a"));"#,
    ["False"]
);
csharp_case!(
    dictionary_foreach_values_after_overwrite,
    r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1 }; map["a"] = 9; foreach (var val in map.Values) Console.WriteLine(val);"#,
    ["9"]
);
csharp_case!(
    dictionary_indexer_sets_brand_new_int_key,
    r#"using System.Collections.Generic; var map = new Dictionary<int, string>(); map[42] = "answer"; Console.WriteLine(map[42]);"#,
    ["answer"]
);
