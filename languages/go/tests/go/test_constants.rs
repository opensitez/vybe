use crate::helpers::*;

#[test]
fn const_integer() {
    let out =
        run_prints("package main; import \"fmt\"; const N = 100; func main() { fmt.Println(N); }");
    assert_eq!(out, vec!["100"]);
}
#[test]
fn const_string() {
    let out = run_prints(
        "package main; import \"fmt\"; const Msg = \"hello\"; func main() { fmt.Println(Msg); }",
    );
    assert_eq!(out, vec!["hello"]);
}
#[test]
fn const_bool() {
    let out = run_prints(
        "package main; import \"fmt\"; const Flag = true; func main() { fmt.Println(Flag); }",
    );
    assert_eq!(out, vec!["true"]);
}
#[test]
fn const_float() {
    let out = run_prints(
        "package main; import \"fmt\"; const Pi = 3.14159; func main() { fmt.Println(Pi); }",
    );
    assert_eq!(out, vec!["3.14159"]);
}
#[test]
fn const_typed() {
    let out = run_prints(
        "package main; import \"fmt\"; const N int = 42; func main() { fmt.Println(N); }",
    );
    assert_eq!(out, vec!["42"]);
}
#[test]
fn const_block() {
    let out = run_prints(
        "package main; import \"fmt\"; const ( A = 1; B = 2; C = 3 ); func main() { fmt.Println(A + B + C); }",
    );
    assert_eq!(out, vec!["6"]);
}
#[test]
fn const_iota_basic() {
    let out = run_prints(
        "package main; import \"fmt\"; const ( A = iota; B = iota; C = iota ); func main() { fmt.Println(A); fmt.Println(B); fmt.Println(C); }",
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}
#[test]
fn const_iota_implicit() {
    let out = run_prints(
        "package main; import \"fmt\"; const ( A = iota; B; C ); func main() { fmt.Println(A); fmt.Println(B); fmt.Println(C); }",
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}
#[test]
fn const_iota_expression() {
    let out = run_prints(
        "package main; import \"fmt\"; const ( A = iota * 2; B; C ); func main() { fmt.Println(A); fmt.Println(B); fmt.Println(C); }",
    );
    assert_eq!(out, vec!["0", "2", "4"]);
}
#[test]
fn const_iota_skip() {
    let out = run_prints(
        "package main; import \"fmt\"; const ( A = iota; _; C ); func main() { fmt.Println(A); fmt.Println(C); }",
    );
    assert_eq!(out, vec!["0", "2"]);
}
#[test]
fn const_shadowing() {
    let out = run_prints(
        "package main; import \"fmt\"; const N = 10; func main() { N := 20; fmt.Println(N); }",
    );
    assert_eq!(out, vec!["20"]);
}
#[test]
fn const_arithmetic() {
    let out = run_prints(
        "package main; import \"fmt\"; const A = 10; const B = 20; const C = A + B; func main() { fmt.Println(C); }",
    );
    assert_eq!(out, vec!["30"]);
}
#[test]
fn const_string_concat() {
    let out = run_prints(
        "package main; import \"fmt\"; const A = \"hello \"; const B = \"world\"; const C = A + B; func main() { fmt.Println(C); }",
    );
    assert_eq!(out, vec!["hello world"]);
}
#[test]
fn const_shift() {
    let out = run_prints(
        "package main; import \"fmt\"; const KB = 1 << 10; const MB = 1 << 20; func main() { fmt.Println(KB); fmt.Println(MB); }",
    );
    assert_eq!(out, vec!["1024", "1048576"]);
}
