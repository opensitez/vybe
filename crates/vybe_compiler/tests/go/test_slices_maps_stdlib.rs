//! Go 1.21+ `slices` (Contains, Clone, Compact, Delete, Insert, Grow) and
//! `maps` (Clone, Copy, DeleteFunc) package functions.

use crate::helpers::*;

go_run_cases! {
    slices_contains_int_found => ("package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Contains([]int{1,2,3}, 2)) }", vec!["true"]),
    slices_contains_int_missing => ("package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Contains([]int{1,2,3}, 9)) }", vec!["false"]),
    slices_contains_empty_slice => ("package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Contains([]int{}, 1)) }", vec!["false"]),
    slices_contains_string_element => ("package main; import \"fmt\"; import \"slices\"; func main() { fmt.Println(slices.Contains([]string{\"go\",\"vybe\"}, \"vybe\")) }", vec!["true"]),
    slices_clone_mutation_isolated => ("package main; import \"fmt\"; import \"slices\"; func main() { orig := []int{1,2,3}; cp := slices.Clone(orig); cp[0] = 99; fmt.Println(orig[0]); fmt.Println(cp[0]) }", vec!["1", "99"]),
    slices_clone_empty_preserves_len => ("package main; import \"fmt\"; import \"slices\"; func main() { cp := slices.Clone([]int{}); fmt.Println(len(cp)) }", vec!["0"]),
    slices_clone_nil_slice => ("package main; import \"fmt\"; import \"slices\"; func main() { var s []int; cp := slices.Clone(s); fmt.Println(cp == nil); fmt.Println(len(cp)) }", vec!["true", "0"]),
    slices_compact_strips_consecutive_dupes => ("package main; import \"fmt\"; import \"slices\"; func main() { s := []int{0,0,1,1,2,3,3}; s = slices.Compact(s); fmt.Println(len(s)); fmt.Println(s[0]); fmt.Println(s[3]) }", vec!["4", "0", "3"]),
    slices_compact_keeps_non_adjacent_dupes => ("package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1,2,1,2}; s = slices.Compact(s); fmt.Println(len(s)); fmt.Println(s[2]) }", vec!["4", "1"]),
    slices_compact_all_identical => ("package main; import \"fmt\"; import \"slices\"; func main() { s := []int{7,7,7,7}; s = slices.Compact(s); fmt.Println(len(s)); fmt.Println(s[0]) }", vec!["1", "7"]),
    slices_delete_middle_range => ("package main; import \"fmt\"; import \"slices\"; func main() { s := []int{0,1,2,3,4}; s = slices.Delete(s, 1, 3); fmt.Println(len(s)); fmt.Println(s[0]); fmt.Println(s[1]); fmt.Println(s[2]) }", vec!["3", "0", "3", "4"]),
    slices_delete_from_start => ("package main; import \"fmt\"; import \"slices\"; func main() { s := []int{9,8,7}; s = slices.Delete(s, 0, 1); fmt.Println(len(s)); fmt.Println(s[0]) }", vec!["2", "8"]),
    slices_delete_from_end => ("package main; import \"fmt\"; import \"slices\"; func main() { s := []int{9,8,7}; s = slices.Delete(s, 2, 3); fmt.Println(len(s)); fmt.Println(s[1]) }", vec!["2", "8"]),
    slices_delete_zero_width_noop => ("package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1,2,3}; s = slices.Delete(s, 1, 1); fmt.Println(len(s)); fmt.Println(s[1]) }", vec!["3", "2"]),
    slices_insert_at_beginning => ("package main; import \"fmt\"; import \"slices\"; func main() { s := []int{2,3}; s = slices.Insert(s, 0, 1); fmt.Println(len(s)); fmt.Println(s[0]); fmt.Println(s[2]) }", vec!["3", "1", "3"]),
    slices_insert_at_middle => ("package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1,3}; s = slices.Insert(s, 1, 2); fmt.Println(len(s)); fmt.Println(s[1]) }", vec!["3", "2"]),
    slices_insert_at_end => ("package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1,2}; s = slices.Insert(s, 2, 3); fmt.Println(len(s)); fmt.Println(s[2]) }", vec!["3", "3"]),
    slices_insert_multiple_values => ("package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1,4}; s = slices.Insert(s, 1, 2, 3); fmt.Println(len(s)); fmt.Println(s[1]); fmt.Println(s[2]) }", vec!["4", "2", "3"]),
    slices_insert_into_nil_slice => ("package main; import \"fmt\"; import \"slices\"; func main() { var s []int; s = slices.Insert(s, 0, 5, 6); fmt.Println(len(s)); fmt.Println(s[1]) }", vec!["2", "6"]),
    slices_grow_expands_capacity => ("package main; import \"fmt\"; import \"slices\"; func main() { s := make([]int, 2, 2); s[0], s[1] = 1, 2; s = slices.Grow(s, 3); fmt.Println(len(s)); fmt.Println(cap(s) >= 5); fmt.Println(s[0]) }", vec!["2", "true", "1"]),
    slices_grow_zero_len_slice => ("package main; import \"fmt\"; import \"slices\"; func main() { s := make([]int, 0, 1); s = slices.Grow(s, 2); fmt.Println(len(s)); fmt.Println(cap(s) >= 3) }", vec!["0", "true"]),
    maps_clone_mutation_isolated => ("package main; import \"fmt\"; import \"maps\"; func main() { orig := map[string]int{\"a\": 1}; cp := maps.Clone(orig); cp[\"a\"] = 9; fmt.Println(orig[\"a\"]); fmt.Println(cp[\"a\"]) }", vec!["1", "9"]),
    maps_clone_empty_map => ("package main; import \"fmt\"; import \"maps\"; func main() { cp := maps.Clone(map[string]int{}); fmt.Println(len(cp)) }", vec!["0"]),
    maps_clone_nil_map => ("package main; import \"fmt\"; import \"maps\"; func main() { var m map[string]int; cp := maps.Clone(m); fmt.Println(cp == nil); fmt.Println(len(cp)) }", vec!["true", "0"]),
    maps_copy_adds_missing_keys => ("package main; import \"fmt\"; import \"maps\"; func main() { dst := map[string]int{\"a\": 1}; src := map[string]int{\"b\": 2}; n := maps.Copy(dst, src); fmt.Println(n); fmt.Println(dst[\"b\"]) }", vec!["1", "2"]),
    maps_copy_overwrites_existing => ("package main; import \"fmt\"; import \"maps\"; func main() { dst := map[string]int{\"a\": 1}; src := map[string]int{\"a\": 9}; maps.Copy(dst, src); fmt.Println(dst[\"a\"]) }", vec!["9"]),
    maps_copy_returns_new_key_count => ("package main; import \"fmt\"; import \"maps\"; func main() { dst := map[string]int{\"x\": 1}; src := map[string]int{\"x\": 2, \"y\": 3}; n := maps.Copy(dst, src); fmt.Println(n); fmt.Println(len(dst)) }", vec!["1", "2"]),
    maps_deletefunc_by_value => ("package main; import \"fmt\"; import \"maps\"; func main() { m := map[int]string{1: \"keep\", 2: \"drop\", 3: \"drop\"}; maps.DeleteFunc(m, func(k int, v string) bool { return v == \"drop\" }); fmt.Println(len(m)); fmt.Println(m[1]) }", vec!["1", "keep"]),
    maps_deletefunc_by_key_prefix => ("package main; import \"fmt\"; import \"maps\"; func main() { m := map[string]int{\"pre1\": 1, \"pre2\": 2, \"other\": 3}; maps.DeleteFunc(m, func(k string, v int) bool { return len(k) >= 5 }); fmt.Println(len(m)); fmt.Println(m[\"pre1\"]) }", vec!["2", "1"]),
    maps_deletefunc_clears_all_entries => ("package main; import \"fmt\"; import \"maps\"; func main() { m := map[int]int{1: 1, 2: 2}; maps.DeleteFunc(m, func(k int, v int) bool { return true }); fmt.Println(len(m)) }", vec!["0"]),
    maps_deletefunc_no_predicate_match => ("package main; import \"fmt\"; import \"maps\"; func main() { m := map[int]int{1: 10, 2: 20}; maps.DeleteFunc(m, func(k int, v int) bool { return v > 100 }); fmt.Println(len(m)); fmt.Println(m[2]) }", vec!["2", "20"]),
}

go_compile_cases! {
    slices_insert_string_slice => "package main; import \"slices\"; func main() { s := []string{\"a\"}; _ = slices.Insert(s, 1, \"b\", \"c\") }",
    slices_grow_byte_slice => "package main; import \"slices\"; func main() { s := make([]byte, 1, 1); _ = slices.Grow(s, 4) }",
    maps_copy_disjoint_key_sets => "package main; import \"maps\"; func main() { dst := map[int]bool{}; src := map[int]bool{1: true, 2: false}; _ = maps.Copy(dst, src) }",
    maps_deletefunc_struct_values => "package main; import \"maps\"; type Pair struct { N int }; func main() { m := map[string]Pair{\"a\": {N: 1}}; maps.DeleteFunc(m, func(k string, v Pair) bool { return v.N == 0 }) }",
}
