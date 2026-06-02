use crate::helpers::*;

#[test]
fn factorial_base_case() {
    let out = run_prints(
        "package main; import \"fmt\"; func fact(n int) int { if n <= 1 { return 1 }; return n * fact(n-1) } func main() { fmt.Println(fact(1)); }",
    );
    assert_eq!(out, vec!["1"]);
}
#[test]
fn factorial_five() {
    let out = run_prints(
        "package main; import \"fmt\"; func fact(n int) int { if n <= 1 { return 1 }; return n * fact(n-1) } func main() { fmt.Println(fact(5)); }",
    );
    assert_eq!(out, vec!["120"]);
}
#[test]
fn factorial_ten() {
    let out = run_prints(
        "package main; import \"fmt\"; func fact(n int) int { if n <= 1 { return 1 }; return n * fact(n-1) } func main() { fmt.Println(fact(10)); }",
    );
    assert_eq!(out, vec!["3628800"]);
}
#[test]
fn fibonacci_base_zero() {
    let out = run_prints(
        "package main; import \"fmt\"; func fib(n int) int { if n <= 0 { return 0 }; if n == 1 { return 1 }; return fib(n-1) + fib(n-2) } func main() { fmt.Println(fib(0)); }",
    );
    assert_eq!(out, vec!["0"]);
}
#[test]
fn fibonacci_one() {
    let out = run_prints(
        "package main; import \"fmt\"; func fib(n int) int { if n <= 0 { return 0 }; if n == 1 { return 1 }; return fib(n-1) + fib(n-2) } func main() { fmt.Println(fib(1)); }",
    );
    assert_eq!(out, vec!["1"]);
}
#[test]
fn fibonacci_seven() {
    let out = run_prints(
        "package main; import \"fmt\"; func fib(n int) int { if n <= 0 { return 0 }; if n == 1 { return 1 }; return fib(n-1) + fib(n-2) } func main() { fmt.Println(fib(7)); }",
    );
    assert_eq!(out, vec!["13"]);
}
#[test]
fn fibonacci_ten() {
    let out = run_prints(
        "package main; import \"fmt\"; func fib(n int) int { if n <= 0 { return 0 }; if n == 1 { return 1 }; return fib(n-1) + fib(n-2) } func main() { fmt.Println(fib(10)); }",
    );
    assert_eq!(out, vec!["55"]);
}
#[test]
fn sum_recursive() {
    let out = run_prints(
        "package main; import \"fmt\"; func sumTo(n int) int { if n <= 0 { return 0 }; return n + sumTo(n-1) } func main() { fmt.Println(sumTo(5)); }",
    );
    assert_eq!(out, vec!["15"]);
}
#[test]
fn sum_recursive_ten() {
    let out = run_prints(
        "package main; import \"fmt\"; func sumTo(n int) int { if n <= 0 { return 0 }; return n + sumTo(n-1) } func main() { fmt.Println(sumTo(10)); }",
    );
    assert_eq!(out, vec!["55"]);
}
#[test]
fn power_recursive() {
    let out = run_prints(
        "package main; import \"fmt\"; func pow(base int, exp int) int { if exp == 0 { return 1 }; return base * pow(base, exp-1) } func main() { fmt.Println(pow(2, 10)); }",
    );
    assert_eq!(out, vec!["1024"]);
}
#[test]
fn countdown_recursive() {
    let out = run_prints(
        "package main; import \"fmt\"; func countdown(n int) { if n < 0 { return }; fmt.Println(n); countdown(n-1); } func main() { countdown(3); }",
    );
    assert_eq!(out, vec!["3", "2", "1", "0"]);
}
#[test]
fn count_down_to_one() {
    let out = run_prints(
        "package main; import \"fmt\"; func countDown(n int) { if n == 0 { return }; fmt.Println(n); countDown(n-1); } func main() { countDown(3); }",
    );
    assert_eq!(out, vec!["3", "2", "1"]);
}
#[test]
fn mutual_recursion_even() {
    let out = run_prints(
        "package main; import \"fmt\"; func isEven(n int) bool { if n == 0 { return true }; return isOdd(n-1) } func isOdd(n int) bool { if n == 0 { return false }; return isEven(n-1) } func main() { fmt.Println(isEven(4)); }",
    );
    assert_eq!(out, vec!["true"]);
}
#[test]
fn mutual_recursion_odd() {
    let out = run_prints(
        "package main; import \"fmt\"; func isEven(n int) bool { if n == 0 { return true }; return isOdd(n-1) } func isOdd(n int) bool { if n == 0 { return false }; return isEven(n-1) } func main() { fmt.Println(isOdd(5)); }",
    );
    assert_eq!(out, vec!["true"]);
}
#[test]
fn recursive_string_reverse() {
    let out = run_prints(
        "package main; import \"fmt\"; func rev(s string) string { if len(s) == 0 { return \"\" }; return rev(s[1:]) + string(s[0]) } func main() { fmt.Println(rev(\"abc\")); }",
    );
    assert_eq!(out, vec!["cba"]);
}
#[test]
fn flattened_recursion_gcd() {
    let out = run_prints(
        "package main; import \"fmt\"; func gcd(a int, b int) int { if b == 0 { return a }; return gcd(b, a % b) } func main() { fmt.Println(gcd(48, 18)); }",
    );
    assert_eq!(out, vec!["6"]);
}
#[test]
fn triangular_number() {
    let out = run_prints(
        "package main; import \"fmt\"; func tri(n int) int { if n <= 0 { return 0 }; return n + tri(n-1) } func main() { fmt.Println(tri(4)); }",
    );
    assert_eq!(out, vec!["10"]);
}
#[test]
fn recursive_len_of_slice() {
    let out = run_prints(
        "package main; import \"fmt\"; func myLen(s []int, acc int) int { if len(s) == 0 { return acc }; return myLen(s[1:], acc+1) } func main() { fmt.Println(myLen([]int{1,2,3,4,5}, 0)); }",
    );
    assert_eq!(out, vec!["5"]);
}
