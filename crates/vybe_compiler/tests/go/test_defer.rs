use crate::helpers::*;

#[test] fn defer_simple() {
    let out = run_prints("package main; import \"fmt\"; func main() { defer fmt.Println(\"deferred\"); fmt.Println(\"first\"); }");
    assert_eq!(out, vec!["first", "deferred"]);
}
#[test] fn defer_multiple_LIFO() {
    let out = run_prints("package main; import \"fmt\"; func main() { defer fmt.Println(\"1\"); defer fmt.Println(\"2\"); defer fmt.Println(\"3\"); }");
    assert_eq!(out, vec!["3", "2", "1"]);
}
#[test] fn defer_with_params() {
    let out = run_prints("package main; import \"fmt\"; func printVal(n int) { fmt.Println(n) }; func main() { defer printVal(10); fmt.Println(5); }");
    assert_eq!(out, vec!["5", "10"]);
}
#[test] fn defer_eval_params_immediately() {
    let out = run_prints("package main; import \"fmt\"; func main() { n := 1; defer fmt.Println(n); n = 2; }");
    assert_eq!(out, vec!["1"]);
}
#[test] fn defer_in_loop() {
    let out = run_prints("package main; import \"fmt\"; func main() { for i := 0; i < 3; i++ { defer fmt.Println(i); }; }");
    assert_eq!(out, vec!["2", "1", "0"]);
}
#[test] fn defer_modifies_named_return() {
    let out = run_prints("package main; import \"fmt\"; func getVal() (result int) { defer func() { result++ }(); return 1; }; func main() { fmt.Println(getVal()); }");
    assert_eq!(out, vec!["2"]);
}
#[test] fn defer_nested_func() {
    let out = run_prints("package main; import \"fmt\"; func inner() { defer fmt.Println(\"inner def\"); fmt.Println(\"inner run\"); }; func main() { defer fmt.Println(\"main def\"); inner(); fmt.Println(\"main run\"); }");
    assert_eq!(out, vec!["inner run", "inner def", "main run", "main def"]);
}
#[test] fn defer_recover_simple() {
    compile_ok("package main; import \"fmt\"; func main() { defer func() { recover() }(); panic(\"test\"); }");
}
#[test] fn defer_closure_capture() {
    let out = run_prints("package main; import \"fmt\"; func main() { n := 1; defer func() { fmt.Println(n) }(); n = 2; }");
    assert_eq!(out, vec!["2"]);
}
#[test] fn defer_method_call() {
    let out = run_prints("package main; import \"fmt\"; type T struct { msg string }; func (t T) Print() { fmt.Println(t.msg) }; func main() { t := T{msg: \"hello\"}; defer t.Print(); t.msg = \"world\"; }");
    assert_eq!(out, vec!["hello"]);
}
