//! SortedDictionary<K,V> and SortedSet<T> ordering, views, and comparison semantics.

csharp_cases! {
    sorted_dictionary_keys_enumerate_in_ascending_order => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [3] = "c", [1] = "a", [2] = "b" }; foreach (var k in sd.Keys) Console.WriteLine(k);"#,
        ["1", "2", "3"]
    };

    sorted_dictionary_values_follow_key_order => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [3] = "c", [1] = "a", [2] = "b" }; foreach (var v in sd.Values) Console.WriteLine(v);"#,
        ["a", "b", "c"]
    };

    sorted_dictionary_foreach_yields_sorted_pairs => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<string, int> { ["b"] = 2, ["a"] = 1, ["c"] = 3 }; foreach (var p in sd) Console.WriteLine(p.Key + ":" + p.Value);"#,
        ["a:1", "b:2", "c:3"]
    };

    sorted_dictionary_indexer_reads_by_key => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<int, string>(); sd[5] = "five"; Console.WriteLine(sd[5]);"#,
        ["five"]
    };

    sorted_dictionary_indexer_overwrites_existing_value => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [1] = "old" }; sd[1] = "new"; Console.WriteLine(sd[1]);"#,
        ["new"]
    };

    sorted_dictionary_containskey_finds_present_key => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [2] = "two" }; Console.WriteLine(sd.ContainsKey(2));"#,
        ["True"]
    };

    sorted_dictionary_containskey_false_for_absent => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [1] = "a" }; Console.WriteLine(sd.ContainsKey(9));"#,
        ["False"]
    };

    sorted_dictionary_count_reflects_insertions => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<int, int> { [1] = 10, [2] = 20, [3] = 30 }; Console.WriteLine(sd.Count);"#,
        ["3"]
    };

    sorted_dictionary_remove_deletes_entry => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [1] = "a", [2] = "b" }; sd.Remove(1); Console.WriteLine(sd.Count);"#,
        ["1"]
    };

    sorted_dictionary_clear_resets_count => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [1] = "a" }; sd.Clear(); Console.WriteLine(sd.Count);"#,
        ["0"]
    };

    sorted_dictionary_first_key_is_minimum => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [10] = "ten", [2] = "two", [7] = "seven" }; int first = 0; foreach (var k in sd.Keys) { first = k; break; } Console.WriteLine(first);"#,
        ["2"]
    };

    sorted_dictionary_last_key_is_maximum => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [10] = "ten", [2] = "two", [7] = "seven" }; int last = 0; foreach (var k in sd.Keys) last = k; Console.WriteLine(last);"#,
        ["10"]
    };

    sorted_dictionary_string_keys_sort_lexicographically => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<string, int> { ["zebra"] = 1, ["apple"] = 2, ["mango"] = 3 }; foreach (var k in sd.Keys) Console.WriteLine(k);"#,
        ["apple", "mango", "zebra"]
    };

    sorted_dictionary_trygetvalue_on_hit => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<string, int> { ["k"] = 4 }; Console.WriteLine(sd.TryGetValue("k", out int v)); Console.WriteLine(v);"#,
        ["True", "4"]
    };

    sorted_dictionary_trygetvalue_on_miss => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<string, int>(); Console.WriteLine(sd.TryGetValue("x", out int v));"#,
        ["False"]
    };

    sorted_set_enumerates_in_ascending_order => {
        r#"using System.Collections.Generic; var ss = new SortedSet<int> { 5, 1, 3, 4, 2 }; foreach (var x in ss) Console.WriteLine(x);"#,
        ["1", "2", "3", "4", "5"]
    };

    sorted_set_rejects_duplicate_elements => {
        r#"using System.Collections.Generic; var ss = new SortedSet<int> { 1, 2, 2, 3, 1 }; Console.WriteLine(ss.Count);"#,
        ["3"]
    };

    sorted_set_min_returns_smallest_element => {
        r#"using System.Collections.Generic; var ss = new SortedSet<int> { 8, 2, 5 }; Console.WriteLine(ss.Min);"#,
        ["2"]
    };

    sorted_set_max_returns_largest_element => {
        r#"using System.Collections.Generic; var ss = new SortedSet<int> { 8, 2, 5 }; Console.WriteLine(ss.Max);"#,
        ["8"]
    };

    sorted_set_add_new_element_increases_count => {
        r#"using System.Collections.Generic; var ss = new SortedSet<int> { 1, 2 }; Console.WriteLine(ss.Add(3)); Console.WriteLine(ss.Count);"#,
        ["True", "3"]
    };

    sorted_set_add_duplicate_returns_false => {
        r#"using System.Collections.Generic; var ss = new SortedSet<int> { 1, 2 }; Console.WriteLine(ss.Add(2));"#,
        ["False"]
    };

    sorted_set_remove_existing_element => {
        r#"using System.Collections.Generic; var ss = new SortedSet<int> { 1, 2, 3 }; ss.Remove(2); Console.WriteLine(ss.Contains(2));"#,
        ["False"]
    };

    sorted_set_contains_finds_member => {
        r#"using System.Collections.Generic; var ss = new SortedSet<string> { "a", "b" }; Console.WriteLine(ss.Contains("b"));"#,
        ["True"]
    };

    sorted_set_get_view_between_inclusive_range => {
        r#"using System.Collections.Generic; var ss = new SortedSet<int> { 1, 2, 3, 4, 5 }; var view = ss.GetViewBetween(2, 4); Console.WriteLine(view.Count);"#,
        ["3"]
    };

    sorted_set_get_view_between_single_element => {
        r#"using System.Collections.Generic; var ss = new SortedSet<int> { 10, 20, 30 }; var view = ss.GetViewBetween(20, 20); Console.WriteLine(view.Min);"#,
        ["20"]
    };

    sorted_set_subset_of_larger_sorted_set => {
        r#"using System.Collections.Generic; var small = new SortedSet<int> { 2, 3 }; var big = new SortedSet<int> { 1, 2, 3, 4 }; Console.WriteLine(small.IsSubsetOf(big));"#,
        ["True"]
    };

    sorted_set_proper_subset_when_strictly_smaller => {
        r#"using System.Collections.Generic; var small = new SortedSet<int> { 2 }; var big = new SortedSet<int> { 1, 2, 3 }; Console.WriteLine(small.IsProperSubsetOf(big));"#,
        ["True"]
    };

    sorted_set_set_equals_with_same_elements => {
        r#"using System.Collections.Generic; var a = new SortedSet<int> { 3, 1, 2 }; var b = new SortedSet<int> { 1, 2, 3 }; Console.WriteLine(a.SetEquals(b));"#,
        ["True"]
    };

    sorted_set_overlaps_detects_shared_element => {
        r#"using System.Collections.Generic; var a = new SortedSet<int> { 1, 2 }; var b = new SortedSet<int> { 2, 9 }; Console.WriteLine(a.Overlaps(b));"#,
        ["True"]
    };

    sorted_set_overlaps_false_when_disjoint => {
        r#"using System.Collections.Generic; var a = new SortedSet<int> { 1, 2 }; var b = new SortedSet<int> { 5, 6 }; Console.WriteLine(a.Overlaps(b));"#,
        ["False"]
    };

    sorted_set_clear_empties_all_members => {
        r#"using System.Collections.Generic; var ss = new SortedSet<int> { 1, 2, 3 }; ss.Clear(); Console.WriteLine(ss.Count);"#,
        ["0"]
    };

    sorted_set_reverse_enumeration_not_guaranteed_but_min_max_hold => {
        r#"using System.Collections.Generic; var ss = new SortedSet<int> { 4, 1, 7 }; Console.WriteLine(ss.Min); Console.WriteLine(ss.Max);"#,
        ["1", "7"]
    };

    sorted_dictionary_insert_out_of_order_still_sorts => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<int, int>(); sd[30] = 3; sd[10] = 1; sd[20] = 2; int sum = 0; foreach (var p in sd) sum += p.Key; Console.WriteLine(sum);"#,
        ["60"]
    };

    sorted_set_view_min_max_match_bounds => {
        r#"using System.Collections.Generic; var ss = new SortedSet<int> { 1, 2, 3, 4, 5, 6 }; var view = ss.GetViewBetween(2, 5); Console.WriteLine(view.Min); Console.WriteLine(view.Max);"#,
        ["2", "5"]
    };

    sorted_dictionary_negative_keys_sort_numerically => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [-1] = "neg", [0] = "zero", [1] = "pos" }; int first = 0; foreach (var k in sd.Keys) { first = k; break; } Console.WriteLine(first);"#,
        ["-1"]
    };

    sorted_set_remove_min_then_new_min => {
        r#"using System.Collections.Generic; var ss = new SortedSet<int> { 1, 2, 3 }; ss.Remove(1); Console.WriteLine(ss.Min);"#,
        ["2"]
    };

    sorted_dictionary_add_via_indexer_increments_count => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<string, int>(); sd["x"] = 1; sd["y"] = 2; Console.WriteLine(sd.Count);"#,
        ["2"]
    };

    sorted_set_union_with_combines_sorted_uniques => {
        r#"using System.Collections.Generic; var a = new SortedSet<int> { 1, 3 }; a.UnionWith(new[] { 2, 3, 4 }); Console.WriteLine(a.Count); Console.WriteLine(a.Min); Console.WriteLine(a.Max);"#,
        ["4", "1", "4"]
    };

    sorted_set_intersect_with_keeps_sorted_overlap => {
        r#"using System.Collections.Generic; var a = new SortedSet<int> { 1, 2, 3, 4 }; a.IntersectWith(new[] { 3, 4, 5 }); Console.WriteLine(a.Count); Console.WriteLine(a.Min);"#,
        ["2", "3"]
    };

    sorted_dictionary_get_value_or_default_on_miss => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<string, int> { ["a"] = 1 }; Console.WriteLine(sd.GetValueOrDefault("z", -1));"#,
        ["-1"]
    };

    sorted_set_except_with_removes_matching_elements => {
        r#"using System.Collections.Generic; var ss = new SortedSet<int> { 1, 2, 3, 4 }; ss.ExceptWith(new[] { 2, 4 }); Console.WriteLine(ss.Count); Console.WriteLine(ss.Contains(1));"#,
        ["2", "True"]
    };

    sorted_dictionary_keys_count_matches_entries => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [1] = "a", [2] = "b" }; Console.WriteLine(sd.Keys.Count);"#,
        ["2"]
    };

    sorted_set_singleton_min_equals_max => {
        r#"using System.Collections.Generic; var ss = new SortedSet<int> { 42 }; Console.WriteLine(ss.Min); Console.WriteLine(ss.Max);"#,
        ["42", "42"]
    };

    sorted_dictionary_values_count_matches_entries => {
        r#"using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [1] = "a", [2] = "b", [3] = "c" }; Console.WriteLine(sd.Values.Count);"#,
        ["3"]
    };

    sorted_set_symmetric_except_with_produces_sorted_result => {
        r#"using System.Collections.Generic; var ss = new SortedSet<int> { 1, 2, 3 }; ss.SymmetricExceptWith(new[] { 2, 3, 4 }); Console.WriteLine(ss.Contains(1)); Console.WriteLine(ss.Contains(4)); Console.WriteLine(ss.Contains(2));"#,
        ["True", "True", "False"]
    };
}
