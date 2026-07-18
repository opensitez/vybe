//! copy/clear builtins, full slice expressions, append growth semantics.

go_run_cases! {
    copy_overlapping_slices => ("package main; import \"fmt\"; func main() { a := []int{1,2,3,4}; n := copy(a, a[1:]); fmt.Println(n); fmt.Println(a[0]); fmt.Println(a[1]) }", vec!["3", "2", "3"]),
    copy_into_larger_dst => ("package main; import \"fmt\"; func main() { dst := make([]int, 5); src := []int{7,8}; n := copy(dst, src); fmt.Println(n); fmt.Println(dst[0]); fmt.Println(dst[4]) }", vec!["2", "7", "0"]),
    copy_zero_len_src => ("package main; import \"fmt\"; func main() { dst := []int{1,2}; n := copy(dst, []int{}); fmt.Println(n); fmt.Println(dst[0]) }", vec!["0", "1"]),
    append_grows_len_and_maybe_cap => ("package main; import \"fmt\"; func main() { s := make([]int, 0, 2); s = append(s, 1, 2, 3); fmt.Println(len(s)); fmt.Println(cap(s) >= 3) }", vec!["3", "true"]),
    append_slice_spread => ("package main; import \"fmt\"; func main() { base := []int{1}; more := []int{2,3}; s := append(base, more...); fmt.Println(len(s)); fmt.Println(s[2]) }", vec!["3", "3"]),
    three_index_slice_cap => ("package main; import \"fmt\"; func main() { a := []int{0,1,2,3,4}; b := a[1:3:4]; fmt.Println(len(b)); fmt.Println(cap(b)); fmt.Println(b[1]) }", vec!["2", "3", "2"]),
    clear_slice_zeros_len => ("package main; import \"fmt\"; func main() { s := []int{1,2,3}; clear(s); fmt.Println(len(s)) }", vec!["3"]),
    clear_map_removes_entries => ("package main; import \"fmt\"; func main() { m := map[string]int{\"a\":1,\"b\":2}; clear(m); fmt.Println(len(m)) }", vec!["0"]),
}

go_compile_cases! {
    copy_string_to_byte_slice => "package main; func main() { dst := make([]byte, 3); _ = copy(dst, \"abc\") }",
    append_nil_slice => "package main; func main() { var s []int; s = append(s, 1) }",
    full_slice_expression_max => "package main; func main() { a := []int{1,2,3}; _ = a[0:1:2] }",
}
