use super::helpers::*;

#[test]
fn enum_basic() {
    let out = run_prints(
        r#"
#include <stdio.h>
enum Color { RED, GREEN, BLUE };
int main() {
    enum Color c = GREEN;
    printf("%d\n", c);
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn enum_with_values() {
    let out = run_prints(
        r#"
#include <stdio.h>
enum Status { OK = 200, NOT_FOUND = 404, ERROR = 500 };
int main() {
    printf("%d\n", OK);
    printf("%d\n", NOT_FOUND);
    printf("%d\n", ERROR);
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["200", "404", "500"]);
}
