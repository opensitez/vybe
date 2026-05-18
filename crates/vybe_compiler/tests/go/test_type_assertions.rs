use crate::helpers::*;

#[test] fn type_assert_success() {
    let out = run_prints("package main; import \"fmt\"; func main() { var x interface{} = \"hello\"; s := x.(string); fmt.Println(s); }");
    assert_eq!(out, vec!["hello"]);
}
#[test] fn type_assert_int() {
    let out = run_prints("package main; import \"fmt\"; func main() { var x interface{} = 42; n := x.(int); fmt.Println(n); }");
    assert_eq!(out, vec!["42"]);
}
#[test] fn type_assert_bool() {
    let out = run_prints("package main; import \"fmt\"; func main() { var x interface{} = true; b := x.(bool); fmt.Println(b); }");
    assert_eq!(out, vec!["true"]);
}
#[test] fn type_assert_struct() {
    let out = run_prints("package main; import \"fmt\"; type Point struct { X int; Y int }; func main() { var x interface{} = Point{X: 1, Y: 2}; p := x.(Point); fmt.Println(p.X); }");
    assert_eq!(out, vec!["1"]);
}
#[test] fn type_assert_ok_idiom_success() {
    let out = run_prints("package main; import \"fmt\"; func main() { var x interface{} = \"hello\"; s, ok := x.(string); fmt.Println(s); fmt.Println(ok); }");
    assert_eq!(out, vec!["hello", "true"]);
}
#[test] fn type_assert_ok_idiom_fail() {
    let out = run_prints("package main; import \"fmt\"; func main() { var x interface{} = 42; s, ok := x.(string); fmt.Println(s); fmt.Println(ok); }");
    assert_eq!(out, vec!["", "false"]); // "" is default string
}
#[test] fn type_switch_int() {
    let out = run_prints("package main; import \"fmt\"; func printType(x interface{}) { switch x.(type) { case int: fmt.Println(\"int\"); case string: fmt.Println(\"string\"); default: fmt.Println(\"other\"); } }; func main() { printType(42); }");
    assert_eq!(out, vec!["int"]);
}
#[test] fn type_switch_string() {
    let out = run_prints("package main; import \"fmt\"; func printType(x interface{}) { switch x.(type) { case int: fmt.Println(\"int\"); case string: fmt.Println(\"string\"); default: fmt.Println(\"other\"); } }; func main() { printType(\"hello\"); }");
    assert_eq!(out, vec!["string"]);
}
#[test] fn type_switch_default() {
    let out = run_prints("package main; import \"fmt\"; func printType(x interface{}) { switch x.(type) { case int: fmt.Println(\"int\"); case string: fmt.Println(\"string\"); default: fmt.Println(\"other\"); } }; func main() { printType(true); }");
    assert_eq!(out, vec!["other"]);
}
#[test] fn type_switch_assign() {
    let out = run_prints("package main; import \"fmt\"; func check(x interface{}) { switch v := x.(type) { case int: fmt.Println(v + 1); case string: fmt.Println(v + \"!\"); default: fmt.Println(\"none\"); } }; func main() { check(10); check(\"hi\"); }");
    assert_eq!(out, vec!["11", "hi!"]);
}
#[test] fn interface_to_interface_assert() {
    compile_ok("package main; type Reader interface { Read() }; type Writer interface { Write() }; type ReadWriter interface { Reader; Writer }; func main() { var rw ReadWriter; var r Reader = rw; _ = r.(Writer); }");
}
#[test] fn type_assert_nil_interface() {
    let out = run_prints("package main; import \"fmt\"; func main() { var x interface{}; s, ok := x.(string); fmt.Println(s); fmt.Println(ok); }");
    assert_eq!(out, vec!["", "false"]);
}
