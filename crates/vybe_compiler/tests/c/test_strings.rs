use super::helpers::*;

#[test]
fn string_length() {
    let out = run_prints(r#"
#include <stdio.h>
#include <string.h>
int main() {
    char *s = "hello";
    printf("%d\n", strlen(s));
    return 0;
}
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn string_compare() {
    let out = run_prints(r#"
#include <stdio.h>
#include <string.h>
int main() {
    char *a = "hello";
    char *b = "hello";
    char *c = "world";
    if (strcmp(a, b) == 0) puts("equal");
    else puts("not equal");
    if (strcmp(a, c) == 0) puts("equal");
    else puts("not equal");
    return 0;
}
"#);
    assert_eq!(out, vec!["equal", "not equal"]);
}

#[test]
fn string_concat() {
    let out = run_prints(r#"
#include <stdio.h>
#include <string.h>
int main() {
    char *a = "hello";
    char *b = " world";
    char *c = strcat(a, b);
    puts(c);
    return 0;
}
"#);
    assert_eq!(out, vec!["hello world"]);
}
