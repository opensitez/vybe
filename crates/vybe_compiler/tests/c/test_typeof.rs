use super::helpers::*;

// typeof / __typeof__ (GCC extension, adopted in C23)
#[test]
fn typeof_basic_variable() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    int x = 42;
    __typeof__(x) y = x + 1;
    printf("%d\n", y);
    return 0;
}
"#,
        &["43"],
    );
}

#[test]
fn typeof_in_macro() {
    assert_outputs(
        r#"
#include <stdio.h>
#define SWAP(a, b) do { __typeof__(a) _t = a; a = b; b = _t; } while(0)
int main() {
    int x = 1, y = 2;
    SWAP(x, y);
    printf("%d %d\n", x, y);
    return 0;
}
"#,
        &["2 1"],
    );
}

#[test]
fn typeof_preserves_pointer_type() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    int x = 99;
    int *p = &x;
    __typeof__(p) q = p;
    printf("%d\n", *q);
    return 0;
}
"#,
        &["99"],
    );
}

#[test]
fn typeof_double_expression() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    double d = 3.14;
    __typeof__(d) e = d * 2;
    printf("%.2f\n", e);
    return 0;
}
"#,
        &["6.28"],
    );
}
