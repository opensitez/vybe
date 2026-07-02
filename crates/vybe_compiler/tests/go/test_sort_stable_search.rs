//! sort package: Ints/Float64s/Strings, Slice/SliceStable, Search, SliceIsSorted,
//! Reverse — distinct from `test_sort_package.rs` (basic smoke) and
//! `test_sort_slice_find.rs` (Find helper, minimal SliceIsSorted).


go_run_cases! {
    sort_ints_negative_values => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{-3, -1, -2, 0, 1}; sort.Ints(a); fmt.Println(a[0]); fmt.Println(a[4]) }",
        vec!["-3", "1"]
    ),
    sort_ints_all_equal => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{5, 5, 5}; sort.Ints(a); fmt.Println(a[0]); fmt.Println(len(a)) }",
        vec!["5", "3"]
    ),
    sort_ints_already_sorted => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{1, 2, 3, 4}; sort.Ints(a); fmt.Println(a[0]); fmt.Println(a[3]) }",
        vec!["1", "4"]
    ),
    sort_ints_reverse_input => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{9, 7, 5, 3, 1}; sort.Ints(a); fmt.Println(a[0]); fmt.Println(a[4]) }",
        vec!["1", "9"]
    ),
    sort_ints_single_element => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{42}; sort.Ints(a); fmt.Println(a[0]) }",
        vec!["42"]
    ),
    sort_ints_empty_slice => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{}; sort.Ints(a); fmt.Println(len(a)) }",
        vec!["0"]
    ),
    sort_ints_two_element_swap => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{2, 1}; sort.Ints(a); fmt.Println(a[0]); fmt.Println(a[1]) }",
        vec!["1", "2"]
    ),
    sort_ints_duplicates_clustered => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{2, 1, 2, 1, 3}; sort.Ints(a); fmt.Println(a[0]); fmt.Println(a[4]) }",
        vec!["1", "3"]
    ),
    sort_ints_mixed_positive_negative => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{10, -10, 0, 5, -5}; sort.Ints(a); fmt.Println(a[0]); fmt.Println(a[2]); fmt.Println(a[4]) }",
        vec!["-10", "0", "10"]
    ),
    sort_ints_are_sorted_after_sort => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{4, 2, 3, 1}; sort.Ints(a); fmt.Println(sort.IntsAreSorted(a)) }",
        vec!["true"]
    ),
    sort_ints_are_sorted_false_before => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{3, 1, 2}; fmt.Println(sort.IntsAreSorted(a)) }",
        vec!["false"]
    ),
    sort_ints_are_sorted_empty => (
        "package main; import \"fmt\"; import \"sort\"; func main() { fmt.Println(sort.IntsAreSorted([]int{})) }",
        vec!["true"]
    ),

    sort_float64s_negative_values => (
        "package main; import \"fmt\"; import \"sort\"; func main() { f := []float64{-1.5, -0.5, -2.0}; sort.Float64s(f); fmt.Println(f[0]); fmt.Println(f[2]) }",
        vec!["-2", "-0.5"]
    ),
    sort_float64s_zeros => (
        "package main; import \"fmt\"; import \"sort\"; func main() { f := []float64{0.0, -0.0, 0.0}; sort.Float64s(f); fmt.Println(len(f)); fmt.Println(f[0]) }",
        vec!["3", "0"]
    ),
    sort_float64s_mixed_signs => (
        "package main; import \"fmt\"; import \"sort\"; func main() { f := []float64{1.1, -1.1, 0.0, 2.2, -2.2}; sort.Float64s(f); fmt.Println(f[0]); fmt.Println(f[4]) }",
        vec!["-2.2", "2.2"]
    ),
    sort_float64s_already_sorted => (
        "package main; import \"fmt\"; import \"sort\"; func main() { f := []float64{0.1, 0.2, 0.3}; sort.Float64s(f); fmt.Println(f[0]); fmt.Println(f[2]) }",
        vec!["0.1", "0.3"]
    ),
    sort_float64s_single => (
        "package main; import \"fmt\"; import \"sort\"; func main() { f := []float64{3.14}; sort.Float64s(f); fmt.Println(f[0]) }",
        vec!["3.14"]
    ),
    sort_float64s_empty => (
        "package main; import \"fmt\"; import \"sort\"; func main() { f := []float64{}; sort.Float64s(f); fmt.Println(len(f)) }",
        vec!["0"]
    ),
    sort_float64s_are_sorted_true => (
        "package main; import \"fmt\"; import \"sort\"; func main() { fmt.Println(sort.Float64sAreSorted([]float64{1.0, 2.0, 3.0})) }",
        vec!["true"]
    ),
    sort_float64s_are_sorted_false => (
        "package main; import \"fmt\"; import \"sort\"; func main() { fmt.Println(sort.Float64sAreSorted([]float64{2.0, 1.0})) }",
        vec!["false"]
    ),
    sort_float64s_duplicates => (
        "package main; import \"fmt\"; import \"sort\"; func main() { f := []float64{1.0, 1.0, 0.5, 1.0}; sort.Float64s(f); fmt.Println(f[0]); fmt.Println(f[3]) }",
        vec!["0.5", "1"]
    ),
    sort_float64s_large_and_small => (
        "package main; import \"fmt\"; import \"sort\"; func main() { f := []float64{1e10, 1e-10, 1.0}; sort.Float64s(f); fmt.Println(f[0] < f[1]); fmt.Println(f[2]) }",
        vec!["true", "10000000000"]
    ),

    sort_strings_case_sensitive => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []string{\"Banana\", \"apple\", \"Cherry\"}; sort.Strings(s); fmt.Println(s[0]); fmt.Println(s[2]) }",
        vec!["Banana", "apple"]
    ),
    sort_strings_empty_element => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []string{\"b\", \"\", \"a\"}; sort.Strings(s); fmt.Println(s[0]); fmt.Println(s[2]) }",
        vec!["", "b"]
    ),
    sort_strings_single => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []string{\"solo\"}; sort.Strings(s); fmt.Println(s[0]) }",
        vec!["solo"]
    ),
    sort_strings_empty_slice => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []string{}; sort.Strings(s); fmt.Println(len(s)) }",
        vec!["0"]
    ),
    sort_strings_prefix_order => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []string{\"foo\", \"foobar\", \"fo\"}; sort.Strings(s); fmt.Println(s[0]); fmt.Println(s[2]) }",
        vec!["fo", "foobar"]
    ),
    sort_strings_are_sorted_true => (
        "package main; import \"fmt\"; import \"sort\"; func main() { fmt.Println(sort.StringsAreSorted([]string{\"a\", \"b\", \"c\"})) }",
        vec!["true"]
    ),
    sort_strings_are_sorted_false => (
        "package main; import \"fmt\"; import \"sort\"; func main() { fmt.Println(sort.StringsAreSorted([]string{\"c\", \"a\"})) }",
        vec!["false"]
    ),
    sort_strings_duplicates => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []string{\"z\", \"a\", \"z\", \"m\"}; sort.Strings(s); fmt.Println(s[0]); fmt.Println(s[3]) }",
        vec!["a", "z"]
    ),
    sort_strings_numeric_strings => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []string{\"10\", \"2\", \"1\"}; sort.Strings(s); fmt.Println(s[0]); fmt.Println(s[2]) }",
        vec!["1", "2"]
    ),

    sort_slice_modulo_three => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{7, 2, 5, 8, 1}; sort.Slice(a, func(i, j int) bool { return a[i]%3 < a[j]%3 }); fmt.Println(a[0]%3); fmt.Println(a[4]%3) }",
        vec!["1", "2"]
    ),
    sort_slice_struct_by_field => (
        "package main; import \"fmt\"; import \"sort\"; type pair struct { k, v int }; func main() { s := []pair{{3, 30}, {1, 10}, {2, 20}}; sort.Slice(s, func(i, j int) bool { return s[i].k < s[j].k }); fmt.Println(s[0].v); fmt.Println(s[2].v) }",
        vec!["10", "30"]
    ),
    sort_slice_descending_ints => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{1, 5, 3, 2, 4}; sort.Slice(a, func(i, j int) bool { return a[i] > a[j] }); fmt.Println(a[0]); fmt.Println(a[4]) }",
        vec!["5", "1"]
    ),
    sort_slice_by_string_length => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []string{\"go\", \"vybe\", \"a\", \"lang\"}; sort.Slice(s, func(i, j int) bool { return len(s[i]) < len(s[j]) }); fmt.Println(s[0]); fmt.Println(s[3]) }",
        vec!["a", "vybe"]
    ),
    sort_slice_bool_custom => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{1, 2, 3, 4, 5, 6}; sort.Slice(a, func(i, j int) bool { return a[i]%2 > a[j]%2 }); fmt.Println(a[0]%2); fmt.Println(a[5]%2) }",
        vec!["1", "0"]
    ),
    sort_slice_float64_abs => (
        "package main; import \"fmt\"; import \"sort\"; func main() { f := []float64{-3.0, 1.0, -2.0, 4.0}; sort.Slice(f, func(i, j int) bool { ai, aj := f[i], f[j]; if ai < 0 { ai = -ai }; if aj < 0 { aj = -aj }; return ai < aj }); fmt.Println(f[0]); fmt.Println(f[3]) }",
        vec!["1", "-3"]
    ),
    sort_slice_empty => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{}; sort.Slice(a, func(i, j int) bool { return a[i] < a[j] }); fmt.Println(len(a)) }",
        vec!["0"]
    ),
    sort_slice_single => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{99}; sort.Slice(a, func(i, j int) bool { return a[i] < a[j] }); fmt.Println(a[0]) }",
        vec!["99"]
    ),

    stable_sort_preserves_equal_key_order_first => (
        "package main; import \"fmt\"; import \"sort\"; type tagged struct { key, ord int }; func main() { s := []tagged{{1, 0}, {2, 0}, {1, 1}, {1, 2}}; sort.SliceStable(s, func(i, j int) bool { return s[i].key < s[j].key }); fmt.Println(s[1].ord); fmt.Println(s[2].ord); fmt.Println(s[3].ord) }",
        vec!["0", "1", "2"]
    ),
    stable_sort_preserves_equal_key_order_second => (
        "package main; import \"fmt\"; import \"sort\"; type tagged struct { key, ord int }; func main() { s := []tagged{{2, 0}, {1, 0}, {2, 1}, {2, 2}}; sort.SliceStable(s, func(i, j int) bool { return s[i].key < s[j].key }); fmt.Println(s[1].ord); fmt.Println(s[2].ord); fmt.Println(s[3].ord) }",
        vec!["0", "1", "2"]
    ),
    stable_sort_three_equal_groups => (
        "package main; import \"fmt\"; import \"sort\"; type rec struct { g, seq int }; func main() { s := []rec{{0, 0}, {1, 0}, {0, 1}, {1, 1}, {0, 2}}; sort.SliceStable(s, func(i, j int) bool { return s[i].g < s[j].g }); fmt.Println(s[0].seq); fmt.Println(s[1].seq); fmt.Println(s[2].seq) }",
        vec!["0", "1", "2"]
    ),
    stable_sort_by_mod_bucket_order => (
        "package main; import \"fmt\"; import \"sort\"; type item struct { v, id int }; func main() { s := []item{{3, 0}, {1, 1}, {4, 2}, {2, 3}, {5, 4}}; sort.SliceStable(s, func(i, j int) bool { return s[i].v%2 < s[j].v%2 }); fmt.Println(s[0].id); fmt.Println(s[1].id); fmt.Println(s[4].id) }",
        vec!["1", "3", "4"]
    ),
    stable_sort_strings_equal_length => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []string{\"bb\", \"aa\", \"cc\", \"dd\"}; sort.SliceStable(s, func(i, j int) bool { return len(s[i]) < len(s[j]) }); fmt.Println(s[0]); fmt.Println(s[3]) }",
        vec!["aa", "dd"]
    ),
    stable_sort_already_sorted_stable => (
        "package main; import \"fmt\"; import \"sort\"; type kv struct { k, ord int }; func main() { s := []kv{{1, 0}, {2, 1}, {3, 2}}; sort.SliceStable(s, func(i, j int) bool { return s[i].k < s[j].k }); fmt.Println(s[0].ord); fmt.Println(s[2].ord) }",
        vec!["0", "2"]
    ),
    stable_sort_all_equal_keys => (
        "package main; import \"fmt\"; import \"sort\"; type tagged struct { key, ord int }; func main() { s := []tagged{{5, 0}, {5, 1}, {5, 2}}; sort.SliceStable(s, func(i, j int) bool { return s[i].key < s[j].key }); fmt.Println(s[0].ord); fmt.Println(s[2].ord) }",
        vec!["0", "2"]
    ),
    stable_sort_reverse_equal_keys => (
        "package main; import \"fmt\"; import \"sort\"; type tagged struct { key, ord int }; func main() { s := []tagged{{1, 2}, {1, 1}, {1, 0}}; sort.SliceStable(s, func(i, j int) bool { return s[i].key < s[j].key }); fmt.Println(s[0].ord); fmt.Println(s[2].ord) }",
        vec!["2", "0"]
    ),

    search_ints_at_first_index => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{2, 4, 6, 8}; fmt.Println(sort.SearchInts(a, 2)) }",
        vec!["0"]
    ),
    search_ints_at_last_index => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{2, 4, 6, 8}; fmt.Println(sort.SearchInts(a, 8)) }",
        vec!["3"]
    ),
    search_ints_before_first => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{10, 20, 30}; fmt.Println(sort.SearchInts(a, 5)) }",
        vec!["0"]
    ),
    search_ints_after_last => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{10, 20, 30}; fmt.Println(sort.SearchInts(a, 99)) }",
        vec!["3"]
    ),
    search_ints_empty_slice => (
        "package main; import \"fmt\"; import \"sort\"; func main() { fmt.Println(sort.SearchInts([]int{}, 1)) }",
        vec!["0"]
    ),
    search_ints_duplicate_insert_point => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{1, 2, 2, 2, 5}; fmt.Println(sort.SearchInts(a, 2)); fmt.Println(sort.SearchInts(a, 3)) }",
        vec!["1", "4"]
    ),
    search_strings_at_middle => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []string{\"alpha\", \"beta\", \"gamma\"}; fmt.Println(sort.SearchStrings(s, \"beta\")) }",
        vec!["1"]
    ),
    search_strings_before_all => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []string{\"m\", \"n\", \"o\"}; fmt.Println(sort.SearchStrings(s, \"a\")) }",
        vec!["0"]
    ),
    search_strings_after_all => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []string{\"m\", \"n\", \"o\"}; fmt.Println(sort.SearchStrings(s, \"z\")) }",
        vec!["3"]
    ),
    search_generic_int_slice => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{10, 20, 30, 40}; i := sort.Search(len(a), func(k int) bool { return a[k] >= 25 }); fmt.Println(i); fmt.Println(a[i]) }",
        vec!["2", "30"]
    ),
    search_generic_first_match => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := []int{1, 3, 5, 7}; i := sort.Search(len(a), func(k int) bool { return a[k] >= 1 }); fmt.Println(i) }",
        vec!["0"]
    ),
    search_float64s_middle => (
        "package main; import \"fmt\"; import \"sort\"; func main() { f := []float64{1.0, 2.5, 4.0}; fmt.Println(sort.SearchFloat64s(f, 2.5)) }",
        vec!["1"]
    ),

    slice_is_sorted_ascending_true => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []int{1, 2, 3, 4}; fmt.Println(sort.SliceIsSorted(s, func(i, j int) bool { return s[i] < s[j] })) }",
        vec!["true"]
    ),
    slice_is_sorted_ascending_false => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []int{1, 3, 2}; fmt.Println(sort.SliceIsSorted(s, func(i, j int) bool { return s[i] < s[j] })) }",
        vec!["false"]
    ),
    slice_is_sorted_descending_true => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []int{5, 4, 3}; fmt.Println(sort.SliceIsSorted(s, func(i, j int) bool { return s[i] > s[j] })) }",
        vec!["true"]
    ),
    slice_is_sorted_descending_false => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []int{5, 3, 4}; fmt.Println(sort.SliceIsSorted(s, func(i, j int) bool { return s[i] > s[j] })) }",
        vec!["false"]
    ),
    slice_is_sorted_empty => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []int{}; fmt.Println(sort.SliceIsSorted(s, func(i, j int) bool { return s[i] < s[j] })) }",
        vec!["true"]
    ),
    slice_is_sorted_single => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []int{7}; fmt.Println(sort.SliceIsSorted(s, func(i, j int) bool { return s[i] < s[j] })) }",
        vec!["true"]
    ),
    slice_is_sorted_by_length => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []string{\"a\", \"bb\", \"ccc\"}; fmt.Println(sort.SliceIsSorted(s, func(i, j int) bool { return len(s[i]) < len(s[j]) })) }",
        vec!["true"]
    ),
    slice_is_sorted_equal_breaks_strict => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := []int{1, 1, 0}; fmt.Println(sort.SliceIsSorted(s, func(i, j int) bool { return s[i] < s[j] })) }",
        vec!["false"]
    ),

    reverse_int_slice_min_max => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := sort.IntSlice{1, 2, 3, 4, 5}; sort.Sort(sort.Reverse(a)); fmt.Println(a[0]); fmt.Println(a[4]) }",
        vec!["5", "1"]
    ),
    reverse_string_slice => (
        "package main; import \"fmt\"; import \"sort\"; func main() { s := sort.StringSlice{\"a\", \"b\", \"c\"}; sort.Sort(sort.Reverse(s)); fmt.Println(s[0]); fmt.Println(s[2]) }",
        vec!["c", "a"]
    ),
    reverse_float64_slice => (
        "package main; import \"fmt\"; import \"sort\"; func main() { f := sort.Float64Slice{1.1, 2.2, 3.3}; sort.Sort(sort.Reverse(f)); fmt.Println(f[0]); fmt.Println(f[2]) }",
        vec!["3.3", "1.1"]
    ),
    reverse_then_search => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := sort.IntSlice{1, 3, 5, 7}; sort.Sort(sort.Reverse(a)); fmt.Println(a[0]); fmt.Println(sort.SearchInts(a, 5)) }",
        vec!["7", "2"]
    ),
    reverse_two_elements => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := sort.IntSlice{10, 20}; sort.Sort(sort.Reverse(a)); fmt.Println(a[0]); fmt.Println(a[1]) }",
        vec!["20", "10"]
    ),
    reverse_single_element => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := sort.IntSlice{42}; sort.Sort(sort.Reverse(a)); fmt.Println(a[0]) }",
        vec!["42"]
    ),
    reverse_empty_slice => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := sort.IntSlice{}; sort.Sort(sort.Reverse(a)); fmt.Println(len(a)) }",
        vec!["0"]
    ),
    reverse_duplicates => (
        "package main; import \"fmt\"; import \"sort\"; func main() { a := sort.IntSlice{1, 2, 2, 3}; sort.Sort(sort.Reverse(a)); fmt.Println(a[0]); fmt.Println(a[3]) }",
        vec!["3", "1"]
    ),
}

go_compile_cases! {
    sort_stable_by_struct_field => "package main; import \"sort\"; type node struct { val, pri int }; func main() { s := []node{{2, 1}, {1, 2}}; sort.SliceStable(s, func(i, j int) bool { return s[i].pri < s[j].pri }) }",
    sort_search_float64s_insert => "package main; import \"sort\"; func main() { _ = sort.SearchFloat64s([]float64{0.5, 1.5}, 1.0) }",
    sort_interface_reverse => "package main; import \"sort\"; type ints []int; func (p ints) Len() int { return len(p) }; func (p ints) Less(i, j int) bool { return p[i] < p[j] }; func (p ints) Swap(i, j int) { p[i], p[j] = p[j], p[i] }; func main() { sort.Sort(sort.Reverse(ints{3, 1, 2})) }",
}
