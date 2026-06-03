use super::helpers::*;

#[test]
fn bitwise_and() {
    let out = run_prints(r#"
#include <stdio.h>
int main() {
    int a = 0b1100;
    int b = 0b1010;
    printf("%d\n", a & b);
    return 0;
}
"#);
    assert_eq!(out, vec!["8"]);
}

#[test]
fn bitwise_or() {
    let out = run_prints(r#"
#include <stdio.h>
int main() {
    int a = 0b1100;
    int b = 0b1010;
    printf("%d\n", a | b);
    return 0;
}
"#);
    assert_eq!(out, vec!["14"]);
}

#[test]
fn bitwise_xor() {
    let out = run_prints(r#"
#include <stdio.h>
int main() {
    int a = 0b1100;
    int b = 0b1010;
    printf("%d\n", a ^ b);
    return 0;
}
"#);
    assert_eq!(out, vec!["6"]);
}

#[test]
fn bitwise_shift() {
    let out = run_prints(r#"
#include <stdio.h>
int main() {
    printf("%d\n", 1 << 4);
    printf("%d\n", 32 >> 2);
    return 0;
}
"#);
    assert_eq!(out, vec!["16", "8"]);
}

#[test]
fn bitwise_not() {
    let out = run_prints(r#"
#include <stdio.h>
int main() {
    int x = 0;
    printf("%d\n", ~x + 1);
    return 0;
}
"#);
    assert_eq!(out, vec!["0"]);
}
