//! sort.Slice, SliceStable, SliceIsSorted and search helpers.


go_run_cases! {
    sort_slice_ints => ("package main; import \"fmt\"; import \"sort\"; func main() { s := []int{3,1,2}; sort.Slice(s, func(i,j int) bool { return s[i] < s[j] }); fmt.Println(s) }", vec!["[1 2 3]"]),
    sort_slice_stable_strings => ("package main; import \"fmt\"; import \"sort\"; func main() { s := []string{\"b\",\"a\",\"a\"}; sort.SliceStable(s, func(i,j int) bool { return s[i] < s[j] }); fmt.Println(s) }", vec!["[a a b]"]),
    sort_slice_is_sorted_true => ("package main; import \"fmt\"; import \"sort\"; func main() { s := []int{1,2,3}; fmt.Println(sort.SliceIsSorted(s, func(i,j int) bool { return s[i] < s[j] })) }", vec!["true"]),
    sort_slice_is_sorted_false => ("package main; import \"fmt\"; import \"sort\"; func main() { s := []int{2,1}; fmt.Println(sort.SliceIsSorted(s, func(i,j int) bool { return s[i] < s[j] })) }", vec!["false"]),
    sort_search_strings => ("package main; import \"fmt\"; import \"sort\"; func main() { s := []string{\"a\",\"c\",\"f\"}; i, ok := sort.Find(s, 1, func(i int) int { return int(s[i][0]) }); fmt.Println(i, ok) }", vec!["1 true"]),
}

go_compile_cases! {
    sort_slice_float64 => "package main; import \"sort\"; func main() { s := []float64{1.2,0.1}; sort.Slice(s, func(i,j int) bool { return s[i] < s[j] }) }",
    sort_slice_stable_struct => "package main; import \"sort\"; type P struct { k int }; func main() { s := []P{{1},{2}}; sort.SliceStable(s, func(i,j int) bool { return s[i].k < s[j].k }) }",
    sort_find_not_found => "package main; import \"sort\"; func main() { s := []int{1,3,5}; _, _ = sort.Find(s, 0, func(i int) int { return s[i] }) }",
}
