use crate::helpers::*;

// ── Slice tests ──────────────────────────────────────────────────────────────

#[test] fn slice_empty_len() {
    let out = run_prints("package main; import \"fmt\"; func main() { s := []int{}; fmt.Println(len(s)); }");
    assert_eq!(out, vec!["0"]);
}
#[test] fn slice_index_first() {
    let out = run_prints("package main; import \"fmt\"; func main() { s := []int{10, 20, 30}; fmt.Println(s[0]); }");
    assert_eq!(out, vec!["10"]);
}
#[test] fn slice_index_last() {
    let out = run_prints("package main; import \"fmt\"; func main() { s := []int{10, 20, 30}; fmt.Println(s[2]); }");
    assert_eq!(out, vec!["30"]);
}
#[test] fn slice_assign_element() {
    let out = run_prints("package main; import \"fmt\"; func main() { s := []int{1, 2, 3}; s[1] = 99; fmt.Println(s[1]); }");
    assert_eq!(out, vec!["99"]);
}
#[test] fn slice_append_one() {
    let out = run_prints("package main; import \"fmt\"; func main() { s := []int{1, 2}; s = append(s, 3); fmt.Println(len(s)); }");
    assert_eq!(out, vec!["3"]);
}
#[test] fn slice_append_multiple() {
    let out = run_prints("package main; import \"fmt\"; func main() { s := []int{}; s = append(s, 10); s = append(s, 20); s = append(s, 30); fmt.Println(len(s)); }");
    assert_eq!(out, vec!["3"]);
}
#[test] fn slice_range_values() {
    let out = run_prints("package main; import \"fmt\"; func main() { s := []int{5, 6, 7}; for _, v := range s { fmt.Println(v); } }");
    assert_eq!(out, vec!["5", "6", "7"]);
}
#[test] fn slice_range_indices() {
    let out = run_prints("package main; import \"fmt\"; func main() { s := []int{100, 200, 300}; for i, _ := range s { fmt.Println(i); } }");
    assert_eq!(out, vec!["0", "1", "2"]);
}
#[test] fn slice_string_elements() {
    let out = run_prints("package main; import \"fmt\"; func main() { s := []string{\"a\", \"b\", \"c\"}; for _, v := range s { fmt.Println(v); } }");
    assert_eq!(out, vec!["a", "b", "c"]);
}
#[test] fn slice_sum() {
    let out = run_prints("package main; import \"fmt\"; func main() { nums := []int{1, 2, 3, 4, 5}; total := 0; for _, n := range nums { total = total + n }; fmt.Println(total); }");
    assert_eq!(out, vec!["15"]);
}
#[test] fn slice_passed_to_func() {
    let out = run_prints("package main; import \"fmt\"; func sumSlice(s []int) int { t := 0; for _, v := range s { t = t + v }; return t } func main() { fmt.Println(sumSlice([]int{1, 2, 3})); }");
    assert_eq!(out, vec!["6"]);
}
#[test] fn slice_returned_from_func() {
    let out = run_prints("package main; import \"fmt\"; func makeSlice() []int { return []int{7, 8, 9} } func main() { s := makeSlice(); fmt.Println(s[0]); }");
    assert_eq!(out, vec!["7"]);
}
#[test] fn slice_nested_loop() {
    let out = run_prints("package main; import \"fmt\"; func main() { matrix := [][]int{{1,2},{3,4}}; fmt.Println(matrix[0][1]); fmt.Println(matrix[1][0]); }");
    assert_eq!(out, vec!["2", "3"]);
}
#[test] fn slice_bool_elements() {
    let out = run_prints("package main; import \"fmt\"; func main() { flags := []bool{true, false, true}; fmt.Println(flags[0]); fmt.Println(flags[1]); }");
    assert_eq!(out, vec!["true", "false"]);
}
#[test] fn slice_max_element() {
    let out = run_prints("package main; import \"fmt\"; func main() { s := []int{3, 1, 4, 1, 5, 9}; m := s[0]; for _, v := range s { if v > m { m = v } }; fmt.Println(m); }");
    assert_eq!(out, vec!["9"]);
}

// ── Map tests ────────────────────────────────────────────────────────────────

#[test] fn map_empty_len() {
    let out = run_prints("package main; import \"fmt\"; func main() { m := map[string]int{}; fmt.Println(len(m)); }");
    assert_eq!(out, vec!["0"]);
}
#[test] fn map_string_int_set_get() {
    let out = run_prints("package main; import \"fmt\"; func main() { m := map[string]int{}; m[\"age\"] = 25; fmt.Println(m[\"age\"]); }");
    assert_eq!(out, vec!["25"]);
}
#[test] fn map_literal_get() {
    let out = run_prints("package main; import \"fmt\"; func main() { m := map[string]string{\"lang\": \"go\"}; fmt.Println(m[\"lang\"]); }");
    assert_eq!(out, vec!["go"]);
}
#[test] fn map_overwrite_key() {
    let out = run_prints("package main; import \"fmt\"; func main() { m := map[string]int{\"x\": 1}; m[\"x\"] = 42; fmt.Println(m[\"x\"]); }");
    assert_eq!(out, vec!["42"]);
}
#[test] fn map_delete_key() {
    let out = run_prints("package main; import \"fmt\"; func main() { m := map[string]int{\"a\": 1, \"b\": 2}; delete(m, \"a\"); fmt.Println(len(m)); }");
    assert_eq!(out, vec!["1"]);
}
#[test] fn map_int_key() {
    let out = run_prints("package main; import \"fmt\"; func main() { m := map[int]string{1: \"one\", 2: \"two\"}; fmt.Println(m[1]); fmt.Println(m[2]); }");
    assert_eq!(out, vec!["one", "two"]);
}
#[test] fn map_count_words() {
    let out = run_prints("package main; import \"fmt\"; func main() { words := []string{\"a\", \"b\", \"a\", \"c\", \"a\"}; freq := map[string]int{}; for _, w := range words { freq[w] = freq[w] + 1 }; fmt.Println(freq[\"a\"]); }");
    assert_eq!(out, vec!["3"]);
}
#[test] fn map_bool_values() {
    let out = run_prints("package main; import \"fmt\"; func main() { seen := map[string]bool{\"go\": true, \"rust\": false}; fmt.Println(seen[\"go\"]); fmt.Println(seen[\"rust\"]); }");
    assert_eq!(out, vec!["true", "false"]);
}
#[test] fn map_multiple_keys_set() {
    let out = run_prints("package main; import \"fmt\"; func main() { m := map[string]int{}; m[\"a\"] = 1; m[\"b\"] = 2; m[\"c\"] = 3; fmt.Println(len(m)); }");
    assert_eq!(out, vec!["3"]);
}
#[test] fn map_used_as_set() {
    let out = run_prints("package main; import \"fmt\"; func main() { s := map[int]bool{1: true, 2: true, 3: true}; fmt.Println(s[2]); fmt.Println(s[5]); }");
    assert_eq!(out, vec!["true", "false"]);
}
#[test] fn map_accumulate_values() {
    let out = run_prints("package main; import \"fmt\"; func main() { scores := map[string]int{\"alice\": 10, \"bob\": 20, \"carol\": 30}; total := 0; for _, v := range scores { total = total + v }; fmt.Println(total); }");
    assert_eq!(out, vec!["60"]);
}
