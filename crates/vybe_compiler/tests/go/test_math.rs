use crate::helpers::*;

#[test]
fn abs_positive() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Abs(5.0)); }",
    );
    assert_eq!(out, vec!["5"]);
}
#[test]
fn abs_negative() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Abs(-3.0)); }",
    );
    assert_eq!(out, vec!["3"]);
}
#[test]
fn floor_positive() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Floor(3.7)); }",
    );
    assert_eq!(out, vec!["3"]);
}
#[test]
fn ceil_positive() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Ceil(3.2)); }",
    );
    assert_eq!(out, vec!["4"]);
}
#[test]
fn sqrt_four() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Sqrt(4.0)); }",
    );
    assert_eq!(out, vec!["2"]);
}
#[test]
fn sqrt_nine() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Sqrt(9.0)); }",
    );
    assert_eq!(out, vec!["3"]);
}
#[test]
fn math_min() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Min(3.0, 7.0)); }",
    );
    assert_eq!(out, vec!["3"]);
}
#[test]
fn math_max() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Max(3.0, 7.0)); }",
    );
    assert_eq!(out, vec!["7"]);
}
#[test]
fn integer_add() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(10 + 5); }");
    assert_eq!(out, vec!["15"]);
}
#[test]
fn integer_sub() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(10 - 3); }");
    assert_eq!(out, vec!["7"]);
}
#[test]
fn integer_mul() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(4 * 6); }");
    assert_eq!(out, vec!["24"]);
}
#[test]
fn integer_div() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(15 / 3); }");
    assert_eq!(out, vec!["5"]);
}
#[test]
fn integer_mod() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(17 % 5); }");
    assert_eq!(out, vec!["2"]);
}
#[test]
fn integer_mod_zero() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(10 % 2); }");
    assert_eq!(out, vec!["0"]);
}
#[test]
fn negative_numbers() {
    let out = run_prints("package main; import \"fmt\"; func main() { x := -5; fmt.Println(x); }");
    assert_eq!(out, vec!["-5"]);
}
#[test]
fn operator_precedence_mul_first() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(2 + 3 * 4); }");
    assert_eq!(out, vec!["14"]);
}
#[test]
fn operator_precedence_parens() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println((2 + 3) * 4); }");
    assert_eq!(out, vec!["20"]);
}
#[test]
fn bitwise_and() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(6 & 3); }");
    assert_eq!(out, vec!["2"]);
}
#[test]
fn bitwise_or() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(6 | 3); }");
    assert_eq!(out, vec!["7"]);
}
#[test]
fn bitwise_xor() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(6 ^ 3); }");
    assert_eq!(out, vec!["5"]);
}
#[test]
fn left_shift() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(1 << 3); }");
    assert_eq!(out, vec!["8"]);
}
#[test]
fn right_shift() {
    let out = run_prints("package main; import \"fmt\"; func main() { fmt.Println(16 >> 2); }");
    assert_eq!(out, vec!["4"]);
}
#[test]
fn sum_in_loop() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { sum := 0; i := 1; for i <= 5 { sum = sum + i; i++ }; fmt.Println(sum); }",
    );
    assert_eq!(out, vec!["15"]);
}
#[test]
fn product_in_loop() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { prod := 1; i := 1; for i <= 5 { prod = prod * i; i++ }; fmt.Println(prod); }",
    );
    assert_eq!(out, vec!["120"]);
}
#[test]
fn max_of_three() {
    let out = run_prints(
        "package main; import \"fmt\"; func maxOf(a int, b int, c int) int { m := a; if b > m { m = b }; if c > m { m = c }; return m } func main() { fmt.Println(maxOf(3, 7, 5)); }",
    );
    assert_eq!(out, vec!["7"]);
}
#[test]
fn min_of_three() {
    let out = run_prints(
        "package main; import \"fmt\"; func minOf(a int, b int, c int) int { m := a; if b < m { m = b }; if c < m { m = c }; return m } func main() { fmt.Println(minOf(3, 7, 1)); }",
    );
    assert_eq!(out, vec!["1"]);
}
#[test]
fn power_manual() {
    let out = run_prints(
        "package main; import \"fmt\"; func pow(base int, exp int) int { result := 1; i := 0; for i < exp { result = result * base; i++ }; return result } func main() { fmt.Println(pow(2, 8)); }",
    );
    assert_eq!(out, vec!["256"]);
}
#[test]
fn gcd() {
    let out = run_prints(
        "package main; import \"fmt\"; func gcd(a int, b int) int { for b != 0 { a, b = b, a % b }; return a } func main() { fmt.Println(gcd(48, 18)); }",
    );
    assert_eq!(out, vec!["6"]);
}
#[test]
fn is_even() {
    let out = run_prints(
        "package main; import \"fmt\"; func isEven(n int) bool { return n % 2 == 0 } func main() { fmt.Println(isEven(4)); }",
    );
    assert_eq!(out, vec!["true"]);
}
#[test]
fn is_odd() {
    let out = run_prints(
        "package main; import \"fmt\"; func isOdd(n int) bool { return n % 2 != 0 } func main() { fmt.Println(isOdd(7)); }",
    );
    assert_eq!(out, vec!["true"]);
}
#[test]
fn count_divisors() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { n := 12; count := 0; i := 1; for i <= n { if n % i == 0 { count++ }; i++ }; fmt.Println(count); }",
    );
    assert_eq!(out, vec!["6"]);
}
#[test]
fn math_pi_constant() {
    compile_ok("package main; import \"math\"; func main() { _ := math.Pi }");
}
