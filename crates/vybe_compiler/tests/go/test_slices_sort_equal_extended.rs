//! Go 1.21+ `slices` sorting, comparison, search, and Replace — distinct from
//! `test_slices_maps_stdlib.rs` (Contains/Clone/Compact/Delete/Insert/Grow) and
//! `test_slices_delete_insert.rs` (basic Delete/Insert/Replace smoke).

go_run_cases! {
    slices_sort_ints_ascending => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{5, 1, 4, 2, 3}; slices.Sort(s); fmt.Println(s[0]); fmt.Println(s[4]) }",
        vec!["1", "5"]
    ),
    slices_sort_ints_already_sorted => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1, 2, 3}; slices.Sort(s); fmt.Println(slices.IsSorted(s)) }",
        vec!["true"]
    ),
    slices_sort_ints_single_element => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{7}; slices.Sort(s); fmt.Println(s[0]) }",
        vec!["7"]
    ),
    slices_sort_ints_empty_slice => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{}; slices.Sort(s); fmt.Println(len(s)) }",
        vec!["0"]
    ),
    slices_sort_strings_lexicographic => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []string{\"cherry\", \"apple\", \"banana\"}; slices.Sort(s); fmt.Println(s[0]); fmt.Println(s[2]) }",
        vec!["apple", "cherry"]
    ),
    slices_sort_float64s_ascending => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []float64{3.3, 1.1, 2.2}; slices.Sort(s); fmt.Println(s[0]); fmt.Println(s[2]) }",
        vec!["1.1", "3.3"]
    ),
    slices_sort_func_by_absolute_value => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{-3, 1, -2, 4}; slices.SortFunc(s, func(a, b int) int { aa, bb := a, b; if aa < 0 { aa = -aa }; if bb < 0 { bb = -bb }; if aa < bb { return -1 }; if aa > bb { return 1 }; return 0 }); fmt.Println(s[0]); fmt.Println(s[3]) }",
        vec!["1", "-3"]
    ),
    slices_sort_func_by_string_length => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []string{\"go\", \"vybe\", \"a\"}; slices.SortFunc(s, func(a, b string) int { if len(a) < len(b) { return -1 }; if len(a) > len(b) { return 1 }; return 0 }); fmt.Println(s[0]); fmt.Println(s[2]) }",
        vec!["a", "vybe"]
    ),
    slices_sort_func_descending_ints => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1, 3, 2}; slices.SortFunc(s, func(a, b int) int { if a > b { return -1 }; if a < b { return 1 }; return 0 }); fmt.Println(s[0]); fmt.Println(s[2]) }",
        vec!["3", "1"]
    ),
    slices_sort_stable_func_preserves_equal_order => (
        "package main; import \"fmt\"; import \"slices\"; func main() { type pair struct { k, ord int }; s := []pair{{1, 0}, {2, 0}, {1, 1}}; slices.SortStableFunc(s, func(a, b pair) int { if a.k < b.k { return -1 }; if a.k > b.k { return 1 }; return 0 }); fmt.Println(s[1].ord); fmt.Println(s[2].ord) }",
        vec!["0", "1"]
    ),
    slices_sort_stable_func_stable_on_equal_keys => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []string{\"bb\", \"aa\", \"ab\"}; slices.SortStableFunc(s, func(a, b string) int { la, lb := len(a), len(b); if la < lb { return -1 }; if la > lb { return 1 }; return 0 }); fmt.Println(s[0]); fmt.Println(s[2]) }",
        vec!["aa", "bb"]
    ),
    slices_is_sorted_true_ascending => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.IsSorted([]int{1, 2, 3, 4})) }",
        vec!["true"]
    ),
    slices_is_sorted_false_unsorted => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.IsSorted([]int{2, 1})) }",
        vec!["false"]
    ),
    slices_is_sorted_true_single => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.IsSorted([]int{9})) }",
        vec!["true"]
    ),
    slices_is_sorted_true_empty => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.IsSorted([]int{})) }",
        vec!["true"]
    ),
    slices_is_sorted_false_duplicate_breaks_strict => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.IsSorted([]int{1, 1, 0})) }",
        vec!["false"]
    ),
    slices_is_sorted_func_custom_order => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{3, 2, 1}; fmt.Println(slices.IsSortedFunc(s, func(a, b int) int { if a > b { return -1 }; if a < b { return 1 }; return 0 })) }",
        vec!["true"]
    ),
    slices_is_sorted_func_false_for_natural_order => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{3, 2, 1}; fmt.Println(slices.IsSortedFunc(s, func(a, b int) int { if a < b { return -1 }; if a > b { return 1 }; return 0 })) }",
        vec!["false"]
    ),
    slices_equal_identical_int_slices => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Equal([]int{1, 2, 3}, []int{1, 2, 3})) }",
        vec!["true"]
    ),
    slices_equal_different_lengths => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Equal([]int{1, 2}, []int{1, 2, 3})) }",
        vec!["false"]
    ),
    slices_equal_different_elements => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Equal([]int{1, 2, 3}, []int{1, 9, 3})) }",
        vec!["false"]
    ),
    slices_equal_both_empty => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Equal([]string{}, []string{})) }",
        vec!["true"]
    ),
    slices_equal_both_nil => (
        "package main; import \"fmt\"; import \"slices\"; func main() { var a, b []int; fmt.Println(slices.Equal(a, b)) }",
        vec!["true"]
    ),
    slices_equal_nil_vs_empty => (
        "package main; import \"fmt\"; import \"slices\"; func main() { var a []int; b := []int{}; fmt.Println(slices.Equal(a, b)) }",
        vec!["true"]
    ),
    slices_compare_less_lexicographic => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Compare([]int{1, 2}, []int{1, 3})) }",
        vec!["-1"]
    ),
    slices_compare_equal_prefix => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Compare([]int{1, 2, 3}, []int{1, 2, 3})) }",
        vec!["0"]
    ),
    slices_compare_greater_shorter_vs_longer => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Compare([]int{1, 2, 3}, []int{1, 2})) }",
        vec!["1"]
    ),
    slices_compare_shorter_less_than_longer_prefix => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Compare([]int{1, 2}, []int{1, 2, 3})) }",
        vec!["-1"]
    ),
    slices_compare_strings_lexicographic => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Compare([]string{\"a\", \"b\"}, []string{\"a\", \"c\"})) }",
        vec!["-1"]
    ),
    slices_index_int_found => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Index([]int{10, 20, 30}, 20)) }",
        vec!["1"]
    ),
    slices_index_int_not_found => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Index([]int{1, 2, 3}, 9)) }",
        vec!["-1"]
    ),
    slices_index_string_found => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Index([]string{\"go\", \"vybe\"}, \"vybe\")) }",
        vec!["1"]
    ),
    slices_index_empty_slice => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Index([]int{}, 1)) }",
        vec!["-1"]
    ),
    slices_index_func_first_even => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.IndexFunc([]int{1, 3, 4, 5}, func(v int) bool { return v%2 == 0 })) }",
        vec!["2"]
    ),
    slices_index_func_no_match => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.IndexFunc([]int{1, 3, 5}, func(v int) bool { return v > 10 })) }",
        vec!["-1"]
    ),
    slices_index_func_string_prefix => (
        "package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.IndexFunc([]string{\"foo\", \"bar\", \"baz\"}, func(s string) bool { return s[0] == 'b' })) }",
        vec!["1"]
    ),
    slices_binary_search_found_exact => (
        "package main; import \"fmt\"; import \"slices\"; func main() { i, ok := slices.BinarySearch([]int{1, 3, 5, 7}, 5); fmt.Println(i); fmt.Println(ok) }",
        vec!["2", "true"]
    ),
    slices_binary_search_not_found_insert_point => (
        "package main; import \"fmt\"; import \"slices\"; func main() { i, ok := slices.BinarySearch([]int{1, 3, 5, 7}, 4); fmt.Println(i); fmt.Println(ok) }",
        vec!["2", "false"]
    ),
    slices_binary_search_before_first => (
        "package main; import \"fmt\"; import \"slices\"; func main() { i, ok := slices.BinarySearch([]int{2, 4, 6}, 0); fmt.Println(i); fmt.Println(ok) }",
        vec!["0", "false"]
    ),
    slices_binary_search_after_last => (
        "package main; import \"fmt\"; import \"slices\"; func main() { i, ok := slices.BinarySearch([]int{2, 4, 6}, 9); fmt.Println(i); fmt.Println(ok) }",
        vec!["3", "false"]
    ),
    slices_binary_search_strings => (
        "package main; import \"fmt\"; import \"slices\"; func main() { i, ok := slices.BinarySearch([]string{\"a\", \"c\", \"e\"}, \"c\"); fmt.Println(i); fmt.Println(ok) }",
        vec!["1", "true"]
    ),
    slices_binary_search_func_custom_order => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{3, 2, 1}; i, ok := slices.BinarySearchFunc(s, 2, func(a, b int) int { if a > b { return -1 }; if a < b { return 1 }; return 0 }); fmt.Println(i); fmt.Println(ok) }",
        vec!["1", "true"]
    ),
    slices_binary_search_func_not_found_desc => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{9, 7, 5}; i, ok := slices.BinarySearchFunc(s, 6, func(a, b int) int { if a > b { return -1 }; if a < b { return 1 }; return 0 }); fmt.Println(i); fmt.Println(ok) }",
        vec!["2", "false"]
    ),
    slices_replace_middle_with_single => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1, 2, 3, 4, 5}; t := slices.Replace(s, 1, 4, 99); fmt.Println(len(t)); fmt.Println(t[1]); fmt.Println(t[2]) }",
        vec!["3", "99", "5"]
    ),
    slices_replace_start_range => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1, 2, 3}; t := slices.Replace(s, 0, 1, 9); fmt.Println(t[0]); fmt.Println(t[1]) }",
        vec!["9", "2"]
    ),
    slices_replace_end_range => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1, 2, 3}; t := slices.Replace(s, 2, 3, 9); fmt.Println(t[1]); fmt.Println(t[2]) }",
        vec!["2", "9"]
    ),
    slices_replace_zero_width_inserts => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1, 3}; t := slices.Replace(s, 1, 1, 2); fmt.Println(len(t)); fmt.Println(t[1]) }",
        vec!["3", "2"]
    ),
    slices_replace_multiple_insert_values => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1, 5}; t := slices.Replace(s, 1, 2, 2, 3, 4); fmt.Println(len(t)); fmt.Println(t[2]); fmt.Println(t[4]) }",
        vec!["5", "3", "5"]
    ),
    slices_replace_delete_only_no_insert => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1, 2, 3, 4}; t := slices.Replace(s, 1, 3); fmt.Println(len(t)); fmt.Println(t[0]); fmt.Println(t[1]) }",
        vec!["2", "1", "4"]
    ),
    slices_sort_after_equal_compare => (
        "package main; import \"fmt\"; import \"slices\"; func main() { a := []int{3, 1, 2}; b := []int{3, 1, 2}; slices.Sort(a); slices.Sort(b); fmt.Println(slices.Equal(a, b)); fmt.Println(slices.Compare(a, b)) }",
        vec!["true", "0"]
    ),
}

go_compile_cases! {
    slices_sort_byte_slice => "package main; import \"slices\"; func main() { s := []byte{'z', 'a', 'm'}; slices.Sort(s) }",
    slices_sort_func_rune_slice => "package main; import \"slices\"; func main() { s := []rune{'日', 'a', 'b'}; slices.SortFunc(s, func(a, b rune) int { if a < b { return -1 }; if a > b { return 1 }; return 0 }) }",
    slices_sort_stable_func_struct_slice => "package main; import \"slices\"; type Item struct { N int }; func main() { s := []Item{{2}, {1}, {2}}; slices.SortStableFunc(s, func(a, b Item) int { if a.N < b.N { return -1 }; if a.N > b.N { return 1 }; return 0 }) }",
    slices_is_sorted_uint_slice => "package main; import \"slices\"; func main() { _ = slices.IsSorted([]uint{1, 2, 3}) }",
    slices_is_sorted_func_float64 => "package main; import \"slices\"; func main() { s := []float64{1.0, 2.0}; _ = slices.IsSortedFunc(s, func(a, b float64) int { if a < b { return -1 }; if a > b { return 1 }; return 0 }) }",
    slices_equal_struct_elements => "package main; import \"slices\"; type P struct { X int }; func main() { _ = slices.Equal([]P{{1}}, []P{{1}}) }",
    slices_compare_empty_vs_nil => "package main; import \"slices\"; func main() { var a []int; _ = slices.Compare(a, []int{}) }",
    slices_index_bool_slice => "package main; import \"slices\"; func main() { _ = slices.Index([]bool{true, false}, false) }",
    slices_index_func_pointer_elements => "package main; import \"slices\"; func main() { s := []*int{new(int)}; _ = slices.IndexFunc(s, func(p *int) bool { return p != nil }) }",
    slices_binary_search_empty_slice => "package main; import \"slices\"; func main() { _, _ = slices.BinarySearch([]int{}, 1) }",
    slices_binary_search_func_empty => "package main; import \"slices\"; func main() { _, _ = slices.BinarySearchFunc([]int{}, 1, func(a, b int) int { return a - b }) }",
    slices_replace_string_elements => "package main; import \"slices\"; func main() { s := []string{\"a\", \"c\"}; _ = slices.Replace(s, 1, 2, \"b\") }",
    slices_replace_full_slice => "package main; import \"slices\"; func main() { s := []int{1, 2, 3}; _ = slices.Replace(s, 0, 3, 9, 8) }",
    slices_sort_func_three_way_tie => "package main; import \"slices\"; func main() { s := []int{2, 2, 2}; slices.SortFunc(s, func(a, b int) int { return 0 }) }",
    slices_binary_search_duplicate_elements => "package main; import \"slices\"; func main() { _, _ = slices.BinarySearch([]int{1, 2, 2, 3}, 2) }",
}
