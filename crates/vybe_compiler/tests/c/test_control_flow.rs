use super::helpers::*;

#[test]
fn switch_basic() {
    let out = run_prints(r#"
#include <stdio.h>
int main() {
    int x = 2;
    switch (x) {
        case 1: puts("one"); break;
        case 2: puts("two"); break;
        case 3: puts("three"); break;
        default: puts("other"); break;
    }
    return 0;
}
"#);
    assert_eq!(out, vec!["two"]);
}

#[test]
fn switch_default() {
    let out = run_prints(r#"
#include <stdio.h>
int main() {
    int x = 99;
    switch (x) {
        case 1: puts("one"); break;
        default: puts("other"); break;
    }
    return 0;
}
"#);
    assert_eq!(out, vec!["other"]);
}

#[test]
fn nested_loops() {
    let out = run_prints(r#"
#include <stdio.h>
int main() {
    for (int i = 0; i < 2; i++) {
        for (int j = 0; j < 2; j++) {
            printf("%d%d\n", i, j);
        }
    }
    return 0;
}
"#);
    assert_eq!(out, vec!["00", "01", "10", "11"]);
}

#[test]
fn early_return() {
    let out = run_prints(r#"
#include <stdio.h>
int sign(int n) {
    if (n > 0) return 1;
    if (n < 0) return -1;
    return 0;
}
int main() {
    printf("%d\n", sign(5));
    printf("%d\n", sign(-3));
    printf("%d\n", sign(0));
    return 0;
}
"#);
    assert_eq!(out, vec!["1", "-1", "0"]);
}
