use crate::helpers::*;

#[test]
fn map_nested() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { m := map[string]map[string]int{}; m[\"a\"] = map[string]int{\"b\": 2}; fmt.Println(m[\"a\"][\"b\"]); }",
    );
    assert_eq!(out, vec!["2"]);
}
#[test]
fn map_slice_values() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { m := map[string][]int{}; m[\"evens\"] = []int{2, 4, 6}; fmt.Println(m[\"evens\"][1]); }",
    );
    assert_eq!(out, vec!["4"]);
}
#[test]
fn map_of_structs() {
    let out = run_prints(
        "package main; import \"fmt\"; type Point struct { X int; Y int }; func main() { m := map[string]Point{\"p1\": {X: 1, Y: 2}}; fmt.Println(m[\"p1\"].X); }",
    );
    assert_eq!(out, vec!["1"]);
}
#[test]
fn map_keys_iteration_order() {
    compile_ok(
        "package main; import \"fmt\"; func main() { m := map[int]int{1: 1, 2: 2}; for k, _ := range m { _ = k } }",
    );
}
#[test]
fn map_delete_non_existent() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { m := map[int]int{1: 1}; delete(m, 2); fmt.Println(len(m)); }",
    );
    assert_eq!(out, vec!["1"]);
}
#[test]
fn map_clear() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { m := map[int]int{1: 1, 2: 2}; m = map[int]int{}; fmt.Println(len(m)); }",
    );
    assert_eq!(out, vec!["0"]);
}
#[test]
fn map_ok_idiom_exists() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { m := map[string]int{\"a\": 1}; v, ok := m[\"a\"]; fmt.Println(v); fmt.Println(ok); }",
    );
    assert_eq!(out, vec!["1", "true"]);
}
#[test]
fn map_ok_idiom_not_exists() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { m := map[string]int{\"a\": 1}; v, ok := m[\"b\"]; fmt.Println(v); fmt.Println(ok); }",
    );
    assert_eq!(out, vec!["0", "false"]);
}
#[test]
fn map_nil_read() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { var m map[string]int; fmt.Println(m[\"a\"]); }",
    );
    assert_eq!(out, vec!["0"]);
}
#[test]
fn map_pointer_values() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 5; m := map[string]*int{\"a\": &x}; fmt.Println(*m[\"a\"]); }",
    );
    assert_eq!(out, vec!["5"]);
}
#[test]
fn map_func_values() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { m := map[string]func() int{\"five\": func() int { return 5 }}; fmt.Println(m[\"five\"]()); }",
    );
    assert_eq!(out, vec!["5"]);
}
#[test]
fn map_bool_keys() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { m := map[bool]string{true: \"yes\", false: \"no\"}; fmt.Println(m[true]); }",
    );
    assert_eq!(out, vec!["yes"]);
}
#[test]
fn slice_append_slice() {
    compile_ok(
        "package main; import \"fmt\"; func main() { s1 := []int{1, 2}; s2 := []int{3, 4}; s1 = append(s1, s2...) }",
    );
}
#[test]
fn slice_copy() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { src := []int{1, 2, 3}; dst := make([]int, 3); copy(dst, src); fmt.Println(dst[1]); }",
    );
    assert_eq!(out, vec!["2"]);
}
#[test]
fn slice_make_len() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := make([]int, 5); fmt.Println(len(s)); }",
    );
    assert_eq!(out, vec!["5"]);
}
#[test]
fn slice_make_len_cap() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := make([]int, 5, 10); fmt.Println(len(s)); fmt.Println(cap(s)); }",
    );
    assert_eq!(out, vec!["5", "10"]);
}
#[test]
fn slice_capacity() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := []int{1, 2, 3}; fmt.Println(cap(s)); }",
    );
    assert_eq!(out, vec!["3"]);
}
#[test]
fn slice_reslice() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := []int{0, 1, 2, 3, 4}; s = s[1:3]; fmt.Println(s[0]); fmt.Println(s[1]); }",
    );
    assert_eq!(out, vec!["1", "2"]);
}
#[test]
fn slice_reslice_end() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := []int{0, 1, 2, 3, 4}; s = s[2:]; fmt.Println(s[0]); }",
    );
    assert_eq!(out, vec!["2"]);
}
#[test]
fn slice_reslice_start() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := []int{0, 1, 2, 3, 4}; s = s[:2]; fmt.Println(len(s)); }",
    );
    assert_eq!(out, vec!["2"]);
}
#[test]
fn slice_pass_by_reference_like() {
    let out = run_prints(
        "package main; import \"fmt\"; func modify(s []int) { s[0] = 99 }; func main() { s := []int{1, 2, 3}; modify(s); fmt.Println(s[0]); }",
    );
    assert_eq!(out, vec!["99"]);
}
