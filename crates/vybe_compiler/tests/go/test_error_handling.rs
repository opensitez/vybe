use crate::helpers::*;

#[test] fn error_interface_compile() {
    compile_ok("package main; type error interface { Error() string }; func main() {}");
}
#[test] fn custom_error_struct() {
    let out = run_prints("package main; import \"fmt\"; type MyErr struct { msg string }; func (e MyErr) Error() string { return e.msg }; func main() { var err error = MyErr{msg: \"failed\"}; fmt.Println(err.Error()); }");
    assert_eq!(out, vec!["failed"]);
}
#[test] fn return_error_nil() {
    let out = run_prints("package main; import \"fmt\"; func doWork() error { return nil }; func main() { err := doWork(); fmt.Println(err == nil); }");
    assert_eq!(out, vec!["true"]);
}
#[test] fn return_error_value() {
    let out = run_prints("package main; import \"fmt\"; type BaseErr struct{}; func (BaseErr) Error() string { return \"err\" }; func doWork() error { return BaseErr{} }; func main() { err := doWork(); fmt.Println(err != nil); }");
    assert_eq!(out, vec!["true"]);
}
#[test] fn error_check_idiom() {
    let out = run_prints("package main; import \"fmt\"; type BasicErr struct{}; func (BasicErr) Error() string { return \"err\" }; func divide(a int, b int) (int, error) { if b == 0 { return 0, BasicErr{} }; return a / b, nil }; func main() { res, err := divide(10, 2); if err != nil { fmt.Println(\"error\") } else { fmt.Println(res) } }");
    assert_eq!(out, vec!["5"]);
}
#[test] fn error_check_idiom_fail() {
    let out = run_prints("package main; import \"fmt\"; type BasicErr struct{}; func (BasicErr) Error() string { return \"err\" }; func divide(a int, b int) (int, error) { if b == 0 { return 0, BasicErr{} }; return a / b, nil }; func main() { res, err := divide(10, 0); if err != nil { fmt.Println(\"error\") } else { fmt.Println(res) } }");
    assert_eq!(out, vec!["error"]);
}
#[test] fn panic_compile() {
    compile_ok("package main; func main() { panic(\"fatal error\") }");
}
#[test] fn recover_compile() {
    compile_ok("package main; func main() { recover() }");
}
#[test] fn defer_recover_pattern() {
    compile_ok("package main; import \"fmt\"; func safeCall() { defer func() { if r := recover(); r != nil { fmt.Println(\"recovered\") } }(); panic(\"oops\"); }; func main() { safeCall() }");
}
#[test] fn multiple_panics_in_defer() {
    compile_ok("package main; func main() { defer func() { recover() }(); panic(\"1\") }");
}
