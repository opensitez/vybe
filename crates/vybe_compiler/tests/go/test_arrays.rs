use crate::helpers::*;

#[test]
fn array_literal_len() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { arr := [3]int{1, 2, 3}; fmt.Println(len(arr)); }",
    );
    assert_eq!(out, vec!["3"]);
}
#[test]
fn array_index_access() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { arr := [4]int{10, 20, 30, 40}; fmt.Println(arr[2]); }",
    );
    assert_eq!(out, vec!["30"]);
}
#[test]
fn array_assign_element() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { arr := [2]int{1, 2}; arr[0] = 99; fmt.Println(arr[0]); }",
    );
    assert_eq!(out, vec!["99"]);
}
#[test]
fn array_iteration() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { arr := [3]int{5, 6, 7}; for _, v := range arr { fmt.Println(v); } }",
    );
    assert_eq!(out, vec!["5", "6", "7"]);
}
#[test]
fn array_sum() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { arr := [5]int{1, 2, 3, 4, 5}; sum := 0; for _, v := range arr { sum += v }; fmt.Println(sum); }",
    );
    assert_eq!(out, vec!["15"]);
}
#[test]
fn array_strings() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { arr := [2]string{\"hello\", \"world\"}; fmt.Println(arr[0]); fmt.Println(arr[1]); }",
    );
    assert_eq!(out, vec!["hello", "world"]);
}
#[test]
fn array_of_structs() {
    let out = run_prints(
        "package main; import \"fmt\"; type Point struct { X int; Y int }; func main() { arr := [2]Point{{X: 1, Y: 2}, {X: 3, Y: 4}}; fmt.Println(arr[1].X); }",
    );
    assert_eq!(out, vec!["3"]);
}
#[test]
fn array_multi_dimensional() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { var arr [2][2]int; arr[0][0] = 1; arr[1][1] = 4; fmt.Println(arr[0][0]); fmt.Println(arr[1][1]); }",
    );
    assert_eq!(out, vec!["1", "4"]);
}
#[test]
fn array_implicit_length() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { arr := [...]int{10, 20, 30}; fmt.Println(len(arr)); }",
    );
    assert_eq!(out, vec!["3"]);
}
#[test]
fn array_passed_by_value() {
    let out = run_prints(
        "package main; import \"fmt\"; func modify(arr [3]int) { arr[0] = 99; }; func main() { arr := [3]int{1, 2, 3}; modify(arr); fmt.Println(arr[0]); }",
    );
    assert_eq!(out, vec!["1"]);
}
#[test]
fn array_equality() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { a := [2]int{1, 2}; b := [2]int{1, 2}; fmt.Println(a == b); }",
    );
    assert_eq!(out, vec!["true"]);
}
#[test]
fn array_inequality() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { a := [2]int{1, 2}; b := [2]int{1, 3}; fmt.Println(a != b); }",
    );
    assert_eq!(out, vec!["true"]);
}
#[test]
fn array_copy() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { a := [2]int{1, 2}; b := a; b[0] = 9; fmt.Println(a[0]); fmt.Println(b[0]); }",
    );
    assert_eq!(out, vec!["1", "9"]);
}
#[test]
fn array_partial_init() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { arr := [3]int{1}; fmt.Println(arr[0]); fmt.Println(arr[1]); }",
    );
    assert_eq!(out, vec!["1", "0"]);
}
#[test]
fn array_bools() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { arr := [3]bool{true, false, true}; fmt.Println(arr[1]); }",
    );
    assert_eq!(out, vec!["false"]);
}
