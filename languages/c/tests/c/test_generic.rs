use super::helpers::*;

// C11 _Generic selection expression
#[test]
fn generic_selects_int_branch() {
    assert_outputs(
        r#"
#include <stdio.h>
#define TYPE_NAME(x) _Generic((x), int: "int", float: "float", double: "double", default: "other")
int main() {
    int x = 0;
    printf("%s\n", TYPE_NAME(x));
    return 0;
}
"#,
        &["int"],
    );
}

#[test]
fn generic_selects_float_branch() {
    assert_outputs(
        r#"
#include <stdio.h>
#define TYPE_NAME(x) _Generic((x), int: "int", float: "float", double: "double", default: "other")
int main() {
    float x = 0.0f;
    printf("%s\n", TYPE_NAME(x));
    return 0;
}
"#,
        &["float"],
    );
}

#[test]
fn generic_selects_double_branch() {
    assert_outputs(
        r#"
#include <stdio.h>
#define TYPE_NAME(x) _Generic((x), int: "int", float: "float", double: "double", default: "other")
int main() {
    double x = 0.0;
    printf("%s\n", TYPE_NAME(x));
    return 0;
}
"#,
        &["double"],
    );
}

#[test]
fn generic_default_branch() {
    assert_outputs(
        r#"
#include <stdio.h>
#define TYPE_NAME(x) _Generic((x), int: "int", float: "float", default: "other")
int main() {
    char x = 'a';
    printf("%s\n", TYPE_NAME(x));
    return 0;
}
"#,
        &["other"],
    );
}

#[test]
fn generic_as_expression_returns_value() {
    assert_outputs(
        r#"
#include <stdio.h>
#define ABS(x) _Generic((x), int: abs((int)(x)), double: fabs((double)(x)))
#include <math.h>
#include <stdlib.h>
int main() {
    int n = -5;
    printf("%d\n", ABS(n));
    return 0;
}
"#,
        &["5"],
    );
}
