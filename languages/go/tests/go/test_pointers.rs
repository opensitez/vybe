use crate::helpers::*;

#[test]
fn pointer_basic() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 10; p := &x; fmt.Println(*p); }",
    );
    assert_eq!(out, vec!["10"]);
}
#[test]
fn pointer_modify() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 10; p := &x; *p = 20; fmt.Println(x); }",
    );
    assert_eq!(out, vec!["20"]);
}
#[test]
fn pointer_to_pointer() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 5; p := &x; pp := &p; fmt.Println(**pp); }",
    );
    assert_eq!(out, vec!["5"]);
}
#[test]
fn pointer_pass_to_func() {
    let out = run_prints(
        "package main; import \"fmt\"; func modify(p *int) { *p = 99 } func main() { x := 1; modify(&x); fmt.Println(x); }",
    );
    assert_eq!(out, vec!["99"]);
}
#[test]
fn pointer_return_from_func() {
    let out = run_prints(
        "package main; import \"fmt\"; func create() *int { x := 42; return &x } func main() { p := create(); fmt.Println(*p); }",
    );
    assert_eq!(out, vec!["42"]);
}
#[test]
fn pointer_to_struct() {
    let out = run_prints(
        "package main; import \"fmt\"; type Point struct { X int; Y int }; func main() { p := &Point{X: 1, Y: 2}; fmt.Println(p.X); }",
    );
    assert_eq!(out, vec!["1"]);
}
#[test]
fn pointer_to_struct_modify() {
    let out = run_prints(
        "package main; import \"fmt\"; type Point struct { X int; Y int }; func main() { p := &Point{X: 1, Y: 2}; p.X = 10; fmt.Println(p.X); }",
    );
    assert_eq!(out, vec!["10"]);
}
#[test]
fn pointer_struct_method_receiver() {
    let out = run_prints(
        "package main; import \"fmt\"; type Counter struct { N int }; func (c *Counter) Inc() { c.N++ }; func main() { c := Counter{N: 0}; c.Inc(); fmt.Println(c.N); }",
    );
    assert_eq!(out, vec!["1"]);
}
#[test]
fn pointer_struct_method_value() {
    let out = run_prints(
        "package main; import \"fmt\"; type Counter struct { N int }; func (c Counter) Inc() { c.N++ }; func main() { c := Counter{N: 0}; c.Inc(); fmt.Println(c.N); }",
    );
    assert_eq!(out, vec!["0"]);
}
#[test]
fn pointer_equality() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 5; p1 := &x; p2 := &x; fmt.Println(p1 == p2); }",
    );
    assert_eq!(out, vec!["true"]);
}
#[test]
fn pointer_inequality() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 5; y := 5; p1 := &x; p2 := &y; fmt.Println(p1 == p2); }",
    );
    assert_eq!(out, vec!["false"]);
}
#[test]
fn pointer_nil_check() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { var p *int; fmt.Println(p == nil); }",
    );
    assert_eq!(out, vec!["true"]);
}
#[test]
fn pointer_to_string() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := \"hello\"; p := &s; *p = \"world\"; fmt.Println(s); }",
    );
    assert_eq!(out, vec!["world"]);
}
#[test]
fn pointer_to_bool() {
    let out = run_prints(
        "package main; import \"fmt\"; func toggle(b *bool) { *b = !*b }; func main() { b := true; toggle(&b); fmt.Println(b); }",
    );
    assert_eq!(out, vec!["false"]);
}
#[test]
fn pointer_array_modify() {
    let out = run_prints(
        "package main; import \"fmt\"; func modify(arr *[3]int) { arr[0] = 99 }; func main() { a := [3]int{1, 2, 3}; modify(&a); fmt.Println(a[0]); }",
    );
    assert_eq!(out, vec!["99"]);
}
