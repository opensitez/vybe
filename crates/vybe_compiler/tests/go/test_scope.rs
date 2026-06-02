use crate::helpers::*;

#[test]
fn scope_inner_shadows_outer() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 1; { x := 2; fmt.Println(x); }; fmt.Println(x); }",
    );
    assert_eq!(out, vec!["2", "1"]);
}
#[test]
fn scope_outer_visible_in_inner() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { x := 42; { fmt.Println(x); } }");
    assert_eq!(out, vec!["42"]);
}
#[test]
fn scope_for_loop_var() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { for i := 0; i < 3; i++ { fmt.Println(i); } }",
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}
#[test]
fn scope_if_init_var() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { if x := 10; x > 5 { fmt.Println(\"big\"); } }",
    );
    assert_eq!(out, vec!["big"]);
}
#[test]
fn global_var_in_func() {
    let out = run_prints(
        "package main; import \"fmt\"; var globalVal = 99; func getGlobal() int { return globalVal } func main() { fmt.Println(getGlobal()); }",
    );
    assert_eq!(out, vec!["99"]);
}
#[test]
fn const_in_func_scope() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { const limit = 100; fmt.Println(limit); }",
    );
    assert_eq!(out, vec!["100"]);
}
#[test]
fn multiple_assignments_swap() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { a := 1; b := 2; a, b = b, a; fmt.Println(a); fmt.Println(b); }",
    );
    assert_eq!(out, vec!["2", "1"]);
}
#[test]
fn blank_identifier_ignore() {
    compile_ok(
        "package main; func pair() (int, int) { return 1, 2 } func main() { _, b := pair(); _ = b }",
    );
}
#[test]
fn short_decl_in_if_init() {
    let out = run_prints(
        "package main; import \"fmt\"; func compute() int { return 7 } func main() { if v := compute(); v > 5 { fmt.Println(\"yes\"); } }",
    );
    assert_eq!(out, vec!["yes"]);
}
#[test]
fn var_block_declaration() {
    let out = run_prints(
        "package main; import \"fmt\"; var ( x = 10; y = 20 ); func main() { fmt.Println(x + y); }",
    );
    assert_eq!(out, vec!["30"]);
}
#[test]
fn const_block_declaration() {
    let out = run_prints(
        "package main; import \"fmt\"; const ( A = 1; B = 2; C = 3 ); func main() { fmt.Println(A + B + C); }",
    );
    assert_eq!(out, vec!["6"]);
}
#[test]
fn type_conversion_int_float() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { i := 5; f := float64(i); fmt.Println(f); }",
    );
    assert_eq!(out, vec!["5"]);
}
#[test]
fn type_conversion_float_int() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { f := 3.7; i := int(f); fmt.Println(i); }",
    );
    assert_eq!(out, vec!["3"]);
}
#[test]
fn type_conversion_string_bytes() {
    compile_ok("package main; func main() { s := \"hello\"; b := []byte(s); _ = b }");
}
#[test]
fn named_return_values_compile() {
    compile_ok(
        "package main; func divide(a, b float64) (result float64, err bool) { if b == 0 { err = true; return }; result = a / b; return } func main() { _, _ = divide(10, 2) }",
    );
}
#[test]
fn multiple_assignment_no_decl() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { a := 0; b := 0; a, b = 5, 10; fmt.Println(a); fmt.Println(b); }",
    );
    assert_eq!(out, vec!["5", "10"]);
}
#[test]
fn shadowing_in_for_loop() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 100; for i := 0; i < 2; i++ { x := i; fmt.Println(x); }; fmt.Println(x); }",
    );
    assert_eq!(out, vec!["0", "1", "100"]);
}
#[test]
fn defer_compile() {
    compile_ok(
        "package main; import \"fmt\"; func main() { defer fmt.Println(\"deferred\"); fmt.Println(\"first\"); }",
    );
}
#[test]
fn goto_compile() {
    compile_ok(
        "package main; import \"fmt\"; func main() { goto done; fmt.Println(\"skip\"); done: fmt.Println(\"done\"); }",
    );
}
#[test]
fn labeled_break_compile() {
    compile_ok(
        "package main; import \"fmt\"; func main() { outer: for i := 0; i < 3; i++ { for j := 0; j < 3; j++ { if j == 1 { break outer }; fmt.Println(i, j); } } }",
    );
}
