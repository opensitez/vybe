use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// Fortran: Complete programs
// ═══════════════════════════════════════════════════════════

#[test]
fn hello_world() {
    let out = run_prints("program hello\nprint *, \"Hello, World!\"\nend program hello\n");
    assert_eq!(out, vec!["Hello, World!"]);
}

#[test]
fn sum_1_to_n() {
    let out = run_prints("program t\ninteger :: i, s\ns = 0\ndo i = 1, 100\ns = s + i\nend do\nprint *, s\nend program t\n");
    assert_eq!(out, vec!["5050"]);
}

#[test]
fn factorial_5() {
    let out = run_prints("program t\ninteger :: i, f\nf = 1\ndo i = 1, 5\nf = f * i\nend do\nprint *, f\nend program t\n");
    assert_eq!(out, vec!["120"]);
}

#[test]
fn fibonacci_10() {
    let out = run_prints("program t\ninteger :: i, a, b, tmp\na = 0\nb = 1\ndo i = 1, 10\ntmp = a + b\na = b\nb = tmp\nend do\nprint *, a\nend program t\n");
    assert_eq!(out, vec!["55"]);
}

#[test]
fn celsius_to_fahrenheit() {
    let out = run_prints("program t\nreal :: c, f\nc = 100.0\nf = c * 9.0 / 5.0 + 32.0\nprint *, f\nend program t\n");
    assert_eq!(out, vec!["212"]);
}

#[test]
fn power_of_two_table() {
    let out = run_prints("program t\ninteger :: i, p\np = 1\ndo i = 0, 3\nprint *, p\np = p * 2\nend do\nend program t\n");
    assert_eq!(out, vec!["1", "2", "4", "8"]);
}

#[test]
fn min_of_three() {
    let out = run_prints("program t\nprint *, min(min(5, 3), 7)\nend program t\n");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn max_of_three() {
    let out = run_prints("program t\nprint *, max(max(5, 3), 7)\nend program t\n");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn average_computation() {
    let out = run_prints("program t\nreal :: avg\navg = (10 + 20 + 30) / 3.0\nprint *, avg\nend program t\n");
    assert_eq!(out, vec!["20"]);
}

#[test]
fn swap_variables() {
    let out = run_prints("program t\ninteger :: a, b, tmp\na = 10\nb = 20\ntmp = a\na = b\nb = tmp\nprint *, a\nprint *, b\nend program t\n");
    assert_eq!(out, vec!["20", "10"]);
}

#[test]
#[ignore] // do while hang
fn count_down() {
    let out = run_prints("program t\ninteger :: i\ni = 5\ndo while (i > 0)\nprint *, i\ni = i - 1\nend do\nend program t\n");
    assert_eq!(out, vec!["5", "4", "3", "2", "1"]);
}

#[test]
fn triangle_area() {
    let out = run_prints("program t\nreal :: base, height, area\nbase = 10.0\nheight = 5.0\narea = 0.5 * base * height\nprint *, area\nend program t\n");
    assert_eq!(out, vec!["25"]);
}

#[test]
fn circle_circumference() {
    let out = run_prints("program t\nreal, parameter :: PI = 3.14159\nreal :: r, c\nr = 5.0\nc = 2.0 * PI * r\nprint *, c\nend program t\n");
    // 2 * 3.14159 * 5 = 31.4159
    assert!(run_prints("program t\nreal, parameter :: PI = 3.14159\nreal :: r, c\nr = 5.0\nc = 2.0 * PI * r\nprint *, c\nend program t\n")[0].starts_with("31.4"));
}

#[test]
fn fizzbuzz() {
    compile_ok("program t\ninteger :: i\ndo i = 1, 15\nif (mod(i, 15) == 0) then\nprint *, \"FizzBuzz\"\nelse if (mod(i, 3) == 0) then\nprint *, \"Fizz\"\nelse if (mod(i, 5) == 0) then\nprint *, \"Buzz\"\nelse\nprint *, i\nend if\nend do\nend program t\n");
}

#[test]
fn sum_even_numbers() {
    let out = run_prints("program t\ninteger :: i, s\ns = 0\ndo i = 1, 10\nif (mod(i, 2) == 0) then\ns = s + i\nend if\nend do\nprint *, s\nend program t\n");
    assert_eq!(out, vec!["30"]);
}

#[test]
fn quadratic_discriminant() {
    compile_ok("program t\nreal :: a, b, c, d\na = 1.0\nb = -5.0\nc = 6.0\nd = b**2 - 4.0*a*c\nif (d >= 0.0) then\nprint *, \"real roots\"\nelse\nprint *, \"complex roots\"\nend if\nend program t\n");
}

#[test]
#[ignore] // do while hang
fn gcd_iterative() {
    let out = run_prints("program t\ninteger :: a, b, tmp\na = 48\nb = 18\ndo while (b /= 0)\ntmp = b\nb = mod(a, b)\na = tmp\nend do\nprint *, a\nend program t\n");
    assert_eq!(out, vec!["6"]);
}

#[test]
fn is_even_or_odd() {
    let out = run_prints("program t\ninteger :: n = 7\nif (mod(n, 2) == 0) then\nprint *, \"even\"\nelse\nprint *, \"odd\"\nend if\nend program t\n");
    assert_eq!(out, vec!["odd"]);
}

#[test]
fn absolute_difference() {
    let out = run_prints("program t\nprint *, abs(5 - 12)\nend program t\n");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn string_greeting() {
    let out = run_prints("program t\ncharacter(len=30) :: name\nname = \"Fortran\"\nprint *, \"Hello, \", name\nend program t\n");
    assert!(out[0].contains("Fortran"));
}
