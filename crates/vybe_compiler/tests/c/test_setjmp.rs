use super::helpers::*;

#[test]
fn setjmp_basic_returns_zero_initially() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <setjmp.h>
static jmp_buf buf;
int main() {
    int v = setjmp(buf);
    if (v == 0) {
        printf("first\n");
        longjmp(buf, 1);
    } else {
        printf("jumped %d\n", v);
    }
    return 0;
}
"#,
        &["first", "jumped 1"],
    );
}

#[test]
fn setjmp_longjmp_skips_code() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <setjmp.h>
static jmp_buf env;
int main() {
    if (setjmp(env) == 0) {
        printf("a\n");
        longjmp(env, 42);
        printf("b\n");
    } else {
        printf("c\n");
    }
    return 0;
}
"#,
        &["a", "c"],
    );
}

#[test]
fn longjmp_passes_value() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <setjmp.h>
static jmp_buf env;
int main() {
    int code = setjmp(env);
    if (code == 0) {
        longjmp(env, 99);
    }
    printf("%d\n", code);
    return 0;
}
"#,
        &["99"],
    );
}

#[test]
fn setjmp_nested_function_longjmp() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <setjmp.h>
static jmp_buf env;
void inner() {
    printf("inner\n");
    longjmp(env, 1);
    printf("unreachable\n");
}
int main() {
    if (setjmp(env) == 0) {
        inner();
    } else {
        printf("returned\n");
    }
    return 0;
}
"#,
        &["inner", "returned"],
    );
}
