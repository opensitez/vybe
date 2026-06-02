use crate::helpers::*;

#[test]
fn bitwise_and() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(12 & 10); }"); // 1100 & 1010 = 1000 (8)
    assert_eq!(out, vec!["8"]);
}
#[test]
fn bitwise_or() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(12 | 10); }"); // 1100 | 1010 = 1110 (14)
    assert_eq!(out, vec!["14"]);
}
#[test]
fn bitwise_xor() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(12 ^ 10); }"); // 1100 ^ 1010 = 0110 (6)
    assert_eq!(out, vec!["6"]);
}
#[test]
fn bitwise_and_not() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(12 &^ 10); }"); // 1100 &^ 1010 = 0100 (4)
    assert_eq!(out, vec!["4"]);
}
#[test]
fn bitwise_left_shift() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(3 << 2); }"); // 3 * 4 = 12
    assert_eq!(out, vec!["12"]);
}
#[test]
fn bitwise_right_shift() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(12 >> 2); }"); // 12 / 4 = 3
    assert_eq!(out, vec!["3"]);
}
#[test]
fn bitwise_not() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(^2); }"); // ~2 = -3
    assert_eq!(out, vec!["-3"]);
}
#[test]
fn bitwise_compound_and() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 12; x &= 10; fmt.Println(x); }",
    );
    assert_eq!(out, vec!["8"]);
}
#[test]
fn bitwise_compound_or() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 12; x |= 10; fmt.Println(x); }",
    );
    assert_eq!(out, vec!["14"]);
}
#[test]
fn bitwise_compound_xor() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 12; x ^= 10; fmt.Println(x); }",
    );
    assert_eq!(out, vec!["6"]);
}
#[test]
fn bitwise_compound_and_not() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 12; x &^= 10; fmt.Println(x); }",
    );
    assert_eq!(out, vec!["4"]);
}
#[test]
fn bitwise_compound_left_shift() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 3; x <<= 2; fmt.Println(x); }",
    );
    assert_eq!(out, vec!["12"]);
}
#[test]
fn bitwise_compound_right_shift() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 12; x >>= 2; fmt.Println(x); }",
    );
    assert_eq!(out, vec!["3"]);
}
