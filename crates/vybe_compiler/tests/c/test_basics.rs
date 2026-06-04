use super::helpers::*;

#[test]
fn hello_world() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    puts("Hello, World!");
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["Hello, World!"]);
}

#[test]
fn integer_arithmetic() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    int a = 10;
    int b = 3;
    int sum = a + b;
    int diff = a - b;
    int prod = a * b;
    int quot = a / b;
    int rem = a % b;
    printf("%d\n", sum);
    printf("%d\n", diff);
    printf("%d\n", prod);
    printf("%d\n", quot);
    printf("%d\n", rem);
    return 0;
}
"#,
    );
    assert_eq!(out[0], "13");
    assert_eq!(out[1], "7");
    assert_eq!(out[2], "30");
    assert_eq!(out[3], "3");
    assert_eq!(out[4], "1");
}

#[test]
fn variables_and_assignment() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    int x = 5;
    x = x + 1;
    printf("%d\n", x);
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn float_arithmetic() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    double a = 3.14;
    double b = 2.0;
    double result = a * b;
    printf("%.2f\n", result);
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["6.28"]);
}

#[test]
fn if_else() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    int x = 10;
    if (x > 5) {
        puts("greater");
    } else {
        puts("not greater");
    }
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["greater"]);
}

#[test]
fn while_loop() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    int i = 0;
    while (i < 3) {
        printf("%d\n", i);
        i++;
    }
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn for_loop() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    for (int i = 0; i < 5; i++) {
        printf("%d\n", i);
    }
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["0", "1", "2", "3", "4"]);
}

#[test]
fn function_call() {
    let out = run_prints(
        r#"
#include <stdio.h>
int add(int a, int b) {
    return a + b;
}
int main() {
    int result = add(3, 4);
    printf("%d\n", result);
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn recursive_factorial() {
    let out = run_prints(
        r#"
#include <stdio.h>
int factorial(int n) {
    if (n <= 1) return 1;
    return n * factorial(n - 1);
}
int main() {
    printf("%d\n", factorial(5));
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["120"]);
}

#[test]
fn string_output() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    char *s = "hello";
    puts(s);
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn ternary_operator() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    int x = 7;
    char *msg = x > 5 ? "big" : "small";
    puts(msg);
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["big"]);
}

#[test]
fn do_while() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    int i = 0;
    do {
        printf("%d\n", i);
        i++;
    } while (i < 3);
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn nested_if() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    int x = 5;
    if (x > 0) {
        if (x > 3) {
            puts("big positive");
        } else {
            puts("small positive");
        }
    } else {
        puts("negative");
    }
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["big positive"]);
}

#[test]
fn compound_assignment() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    int x = 10;
    x += 5;
    x -= 2;
    x *= 3;
    printf("%d\n", x);
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["39"]);
}

#[test]
fn break_in_loop() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    for (int i = 0; i < 10; i++) {
        if (i == 3) break;
        printf("%d\n", i);
    }
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn continue_in_loop() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    for (int i = 0; i < 5; i++) {
        if (i == 2) continue;
        printf("%d\n", i);
    }
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["0", "1", "3", "4"]);
}
