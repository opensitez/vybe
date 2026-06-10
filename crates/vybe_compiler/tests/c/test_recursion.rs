use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { declarations: $decls:expr, body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>"], $decls, $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    factorial_recursive => {
        declarations: "int fact(int n) { return n <= 1 ? 1 : n * fact(n - 1); }",
        body: "printf(\"%d\\n\", fact(5));\nreturn 0;",
        expect: ["120"]
    },
    fibonacci_recursive => {
        declarations: "int fib(int n) { return n <= 1 ? n : fib(n-1) + fib(n-2); }",
        body: "printf(\"%d %d %d\\n\", fib(5), fib(7), fib(10));\nreturn 0;",
        expect: ["5 13 55"]
    },
    gcd_recursive => {
        declarations: "int gcd(int a, int b) { return b == 0 ? a : gcd(b, a % b); }",
        body: "printf(\"%d\\n\", gcd(48, 18));\nreturn 0;",
        expect: ["6"]
    },
    power_recursive => {
        declarations: "int power(int base, int exp) { return exp == 0 ? 1 : base * power(base, exp - 1); }",
        body: "printf(\"%d %d\\n\", power(2, 10), power(3, 4));\nreturn 0;",
        expect: ["1024 81"]
    },
    sum_recursive => {
        declarations: "int sum(int n) { return n == 0 ? 0 : n + sum(n - 1); }",
        body: "printf(\"%d\\n\", sum(10));\nreturn 0;",
        expect: ["55"]
    },
    mutual_recursion => {
        declarations: "int is_even(int n); int is_odd(int n) { return n == 0 ? 0 : is_even(n - 1); } int is_even(int n) { return n == 0 ? 1 : is_odd(n - 1); }",
        body: "printf(\"%d %d\\n\", is_even(4), is_odd(7));\nreturn 0;",
        expect: ["1 1"]
    },
    count_digits_recursive => {
        declarations: "int digits(int n) { return n < 10 ? 1 : 1 + digits(n / 10); }",
        body: "printf(\"%d %d %d\\n\", digits(5), digits(42), digits(1000));\nreturn 0;",
        expect: ["1 2 4"]
    }
}
