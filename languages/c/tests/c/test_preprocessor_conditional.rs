use super::helpers::*;

#[test]
fn define_check_defined() {
    assert_outputs(
        r#"
#include <stdio.h>
#define MY_FLAG
int main() {
#if defined(MY_FLAG)
    printf("defined\n");
#else
    printf("not defined\n");
#endif
    return 0;
}
"#,
        &["defined"],
    );
}

#[test]
fn define_value_comparison() {
    assert_outputs(
        r#"
#include <stdio.h>
#define VERSION 2
int main() {
#if VERSION >= 2
    printf("v2+\n");
#else
    printf("v1\n");
#endif
    return 0;
}
"#,
        &["v2+"],
    );
}

#[test]
fn elif_chain() {
    assert_outputs(
        r#"
#include <stdio.h>
#define PLATFORM 2
int main() {
#if PLATFORM == 1
    printf("platform1\n");
#elif PLATFORM == 2
    printf("platform2\n");
#else
    printf("other\n");
#endif
    return 0;
}
"#,
        &["platform2"],
    );
}

#[test]
fn nested_conditional_compilation() {
    assert_outputs(
        r#"
#include <stdio.h>
#define A
#define B
int main() {
#ifdef A
  #ifdef B
    printf("both\n");
  #else
    printf("a only\n");
  #endif
#endif
    return 0;
}
"#,
        &["both"],
    );
}

#[test]
fn conditional_define_selects_value() {
    assert_outputs(
        r#"
#include <stdio.h>
#define RELEASE
#ifdef RELEASE
#define LOG_LEVEL 0
#else
#define LOG_LEVEL 3
#endif
int main() {
    printf("%d\n", LOG_LEVEL);
    return 0;
}
"#,
        &["0"],
    );
}
