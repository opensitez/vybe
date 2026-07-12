//! sort package: Ints, Float64s, Strings, Search, IsSorted.

go_run_cases! {
    sort_ints_ascending => ("package main; import \"fmt\"; import \"sort\"; func main() { a := []int{3,1,2}; sort.Ints(a); fmt.Println(a[0]); fmt.Println(a[2]) }", vec!["1", "3"]),
    sort_strings_lexicographic => ("package main; import \"fmt\"; import \"sort\"; func main() { s := []string{\"b\",\"a\",\"c\"}; sort.Strings(s); fmt.Println(s[0]); fmt.Println(s[2]) }", vec!["a", "c"]),
    sort_float64s_ascending => ("package main; import \"fmt\"; import \"sort\"; func main() { f := []float64{2.5,1.1,3.3}; sort.Float64s(f); fmt.Println(f[0]); fmt.Println(f[2]) }", vec!["1.1", "3.3"]),
    sort_search_ints_found => ("package main; import \"fmt\"; import \"sort\"; func main() { a := []int{1,3,5}; fmt.Println(sort.SearchInts(a, 3)) }", vec!["1"]),
    sort_search_ints_insert_point => ("package main; import \"fmt\"; import \"sort\"; func main() { a := []int{1,3,5}; fmt.Println(sort.SearchInts(a, 4)) }", vec!["2"]),
    sort_is_sorted_true => ("package main; import \"fmt\"; import \"sort\"; func main() { fmt.Println(sort.IntsAreSorted([]int{1,2,3})) }", vec!["true"]),
    sort_is_sorted_false => ("package main; import \"fmt\"; import \"sort\"; func main() { fmt.Println(sort.IntsAreSorted([]int{2,1})) }", vec!["false"]),
    sort_slice_stable_by_len => ("package main; import \"fmt\"; import \"sort\"; func main() { items := []string{\"aaa\",\"b\",\"cc\"}; sort.SliceStable(items, func(i,j int) bool { return len(items[i]) < len(items[j]) }); fmt.Println(items[0]); fmt.Println(items[2]) }", vec!["b", "aaa"]),
    sort_slice_unstable_desc => ("package main; import \"fmt\"; import \"sort\"; func main() { a := []int{1,3,2}; sort.Slice(a, func(i,j int) bool { return a[i] > a[j] }); fmt.Println(a[0]); fmt.Println(a[2]) }", vec!["3", "1"]),
    sort_search_strings => ("package main; import \"fmt\"; import \"sort\"; func main() { s := []string{\"a\",\"c\",\"e\"}; fmt.Println(sort.SearchStrings(s, \"c\")) }", vec!["1"]),
}

go_compile_cases! {
    sort_reverse_ints => "package main; import \"sort\"; func main() { a := []int{1,2,3}; sort.Sort(sort.Reverse(sort.IntSlice(a))) }",
    sort_interface_sort => "package main; import \"sort\"; type ints []int; func (p ints) Len() int { return len(p) }; func (p ints) Less(i,j int) bool { return p[i] < p[j] }; func (p ints) Swap(i,j int) { p[i], p[j] = p[j], p[i] }; func main() { sort.Sort(ints{2,1}) }",
    sort_search_float64s => "package main; import \"sort\"; func main() { _ = sort.SearchFloat64s([]float64{1.0,2.0}, 1.5) }",
    sort_strings_are_sorted => "package main; import \"sort\"; func main() { _ = sort.StringsAreSorted([]string{\"a\",\"b\"}) }",
    sort_float64s_are_sorted => "package main; import \"sort\"; func main() { _ = sort.Float64sAreSorted([]float64{1.0}) }",
}
