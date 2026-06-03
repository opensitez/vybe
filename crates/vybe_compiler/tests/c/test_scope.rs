use super::helpers::*;

#[test]
fn block_scope() {
    let out = run_prints(r#"
#include <stdio.h>
int main() {
    int x = 1;
    {
        int x = 2;
        printf("%d\n", x);
    }
    printf("%d\n", x);
    return 0;
}
"#);
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn global_constant() {
    let out = run_prints(r#"
#include <stdio.h>
#define MAX 100
int main() {
    printf("%d\n", MAX);
    return 0;
}
"#);
    assert_eq!(out, vec!["100"]);
}

#[test]
fn multiple_functions() {
    let out = run_prints(r#"
#include <stdio.h>
int square(int n) { return n * n; }
int cube(int n) { return n * n * n; }
int main() {
    printf("%d\n", square(4));
    printf("%d\n", cube(3));
    return 0;
}
"#);
    assert_eq!(out, vec!["16", "27"]);
}
