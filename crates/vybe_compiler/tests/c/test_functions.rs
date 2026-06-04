use super::helpers::*;

#[test]
fn multiple_params() {
    let out = run_prints(
        r#"
#include <stdio.h>
int max(int a, int b) {
    return a > b ? a : b;
}
int main() {
    printf("%d\n", max(3, 7));
    printf("%d\n", max(9, 2));
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["7", "9"]);
}

#[test]
fn recursive_fibonacci() {
    let out = run_prints(
        r#"
#include <stdio.h>
int fib(int n) {
    if (n <= 1) return n;
    return fib(n - 1) + fib(n - 2);
}
int main() {
    printf("%d\n", fib(0));
    printf("%d\n", fib(1));
    printf("%d\n", fib(7));
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["0", "1", "13"]);
}

#[test]
fn void_function() {
    let out = run_prints(
        r#"
#include <stdio.h>
void greet(char *name) {
    printf("Hello %s\n", name);
}
int main() {
    greet("Alice");
    greet("Bob");
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["Hello Alice", "Hello Bob"]);
}

#[test]
fn function_pointers_via_variable() {
    let out = run_prints(
        r#"
#include <stdio.h>
int double_it(int x) { return x * 2; }
int main() {
    int result = double_it(5);
    printf("%d\n", result);
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["10"]);
}
