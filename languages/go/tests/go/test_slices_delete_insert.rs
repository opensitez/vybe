//! slices.Delete, Insert, Replace, Compact (Go 1.21+ slices package extras).

go_run_cases! {
    slices_delete_middle => ("package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1,2,3,4}; t := slices.Delete(s, 1, 3); fmt.Println(t) }", vec!["[1 4]"]),
    slices_insert => ("package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1,3}; t := slices.Insert(s, 1, 2); fmt.Println(t) }", vec!["[1 2 3]"]),
    slices_replace => ("package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1,2,3,4}; t := slices.Replace(s, 1, 3, 9); fmt.Println(t) }", vec!["[1 9 4]"]),
    slices_compact => ("package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1,0,2,0,3}; t := slices.Compact(s); fmt.Println(t) }", vec!["[1 0 2 0 3]"]),
    slices_grow => ("package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1}; t := slices.Grow(s, 2); fmt.Println(len(t), cap(t) >= 3) }", vec!["1 true"]),
    slices_clip => ("package main; import \"fmt\"; import \"slices\"; func main() { s := make([]int, 3, 10); t := slices.Clip(s); fmt.Println(len(t), cap(t)) }", vec!["3 3"]) }

go_compile_cases! {
    slices_delete_func => "package main; import \"slices\"; func main() { s := []int{1,2,3}; _ = slices.DeleteFunc(s, func(v int) bool { return v == 2 }) }",
    slices_compact_func => "package main; import \"slices\"; func main() { s := []string{\"a\",\"\",\"b\"}; _ = slices.CompactFunc(s, func(a string) bool { return a == \"\" }) }",
    slices_replace_large => "package main; import \"slices\"; func main() { s := []int{1,2,3,4,5}; _ = slices.Replace(s, 1, 4, 9, 8) }",
    slices_insert_slice => "package main; import \"slices\"; func main() { s := []int{1,4}; _ = slices.Insert(s, 1, []int{2,3}...) }" }
