use crate::helpers::*;

#[test]
fn hello_world() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { fmt.Println(\"hello world\"); }");
    assert_eq!(out, vec!["hello world"]);
}
#[test]
fn variable_declaration() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { var x int = 42; fmt.Println(x); }");
    assert_eq!(out, vec!["42"]);
}
#[test]
fn short_var_declaration() {
    let out = run_prints("package main; import \"fmt\"; func main() { x := 42; fmt.Println(x); }");
    assert_eq!(out, vec!["42"]);
}
#[test]
fn multiple_variables() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { a, b := 1, 2; fmt.Println(a); fmt.Println(b); }",
    );
    assert_eq!(out, vec!["1", "2"]);
}
#[test]
fn const_declaration() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { const Pi = 3.14; fmt.Println(Pi); }",
    );
    assert_eq!(out, vec!["3.14"]);
}
#[test]
fn string_literal() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { s := \"hello\"; fmt.Println(s); }");
    assert_eq!(out, vec!["hello"]);
}
#[test]
fn numeric_literals() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { a := 42; b := 3.14; c := 0xFF; fmt.Println(a); fmt.Println(b); fmt.Println(c); }",
    );
    assert_eq!(out, vec!["42", "3.14", "255"]);
}
#[test]
fn arithmetic() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { a := 10; b := 3; fmt.Println(a + b); fmt.Println(a - b); fmt.Println(a * b); fmt.Println(a / b); }",
    );
    assert_eq!(out, vec!["13", "7", "30", "3"]);
}
#[test]
fn assignment() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { x := 5; x = 10; fmt.Println(x); }");
    assert_eq!(out, vec!["10"]);
}
#[test]
fn compound_assignment() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { x := 5; x += 3; fmt.Println(x); }");
    assert_eq!(out, vec!["8"]);
}
#[test]
fn compound_sub_assign() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 10; x -= 3; fmt.Println(x); }",
    );
    assert_eq!(out, vec!["7"]);
}
#[test]
fn compound_mul_assign() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { x := 4; x *= 3; fmt.Println(x); }");
    assert_eq!(out, vec!["12"]);
}
#[test]
fn compound_div_assign() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 12; x /= 4; fmt.Println(x); }",
    );
    assert_eq!(out, vec!["3"]);
}
#[test]
fn compound_mod_assign() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 10; x %= 3; fmt.Println(x); }",
    );
    assert_eq!(out, vec!["1"]);
}
#[test]
fn increment() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { x := 5; x++; fmt.Println(x); }");
    assert_eq!(out, vec!["6"]);
}
#[test]
fn decrement() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { x := 5; x--; fmt.Println(x); }");
    assert_eq!(out, vec!["4"]);
}
#[test]
fn nil_value() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { var x interface{} = nil; fmt.Println(x); }",
    );
    assert_eq!(out, vec!["null"]);
}
#[test]
fn bool_true() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { b := true; fmt.Println(b); }");
    assert_eq!(out, vec!["true"]);
}
#[test]
fn bool_false() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { b := false; fmt.Println(b); }");
    assert_eq!(out, vec!["false"]);
}
#[test]
fn bool_not() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { b := true; fmt.Println(!b); }");
    assert_eq!(out, vec!["false"]);
}
#[test]
fn comparison_eq() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(5 == 5); }");
    assert_eq!(out, vec!["true"]);
}
#[test]
fn comparison_neq() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(5 != 6); }");
    assert_eq!(out, vec!["true"]);
}
#[test]
fn comparison_lt() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(3 < 5); }");
    assert_eq!(out, vec!["true"]);
}
#[test]
fn comparison_lte() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(5 <= 5); }");
    assert_eq!(out, vec!["true"]);
}
#[test]
fn comparison_gt() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(7 > 3); }");
    assert_eq!(out, vec!["true"]);
}
#[test]
fn comparison_gte() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(5 >= 5); }");
    assert_eq!(out, vec!["true"]);
}
#[test]
fn and_operator() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { fmt.Println(true && true); }");
    assert_eq!(out, vec!["true"]);
}
#[test]
fn and_operator_false() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { fmt.Println(true && false); }");
    assert_eq!(out, vec!["false"]);
}
#[test]
fn or_operator() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { fmt.Println(false || true); }");
    assert_eq!(out, vec!["true"]);
}
#[test]
fn three_vars_same_type() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { var a, b, c int = 1, 2, 3; fmt.Println(a + b + c); }",
    );
    assert_eq!(out, vec!["6"]);
}
#[test]
fn var_no_initializer() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { var x int; fmt.Println(x); }");
    assert_eq!(out, vec!["0"]);
}
#[test]
fn hex_literal() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(0xFF); }");
    assert_eq!(out, vec!["255"]);
}
#[test]
fn octal_literal() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(0o10); }");
    assert_eq!(out, vec!["8"]);
}
#[test]
fn binary_literal() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(0b1010); }");
    assert_eq!(out, vec!["10"]);
}
