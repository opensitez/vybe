use super::helpers::*;

// Stringize operator (#)
#[test]
fn stringize_operator() {
    assert_outputs(
        r#"
#include <stdio.h>
#define STRINGIFY(x) #x
int main() {
    printf("%s\n", STRINGIFY(hello));
    return 0;
}
"#,
        &["hello"],
    );
}

// Token-pasting operator (##)
#[test]
fn token_paste_operator() {
    assert_outputs(
        r#"
#include <stdio.h>
#define CONCAT(a, b) a##b
int main() {
    int xy = 42;
    printf("%d\n", CONCAT(x, y));
    return 0;
}
"#,
        &["42"],
    );
}

// Variadic macros (__VA_ARGS__)
#[test]
fn variadic_macro_basic() {
    assert_outputs(
        r#"
#include <stdio.h>
#define LOG(fmt, ...) printf(fmt, __VA_ARGS__)
int main() {
    LOG("%d %s\n", 42, "test");
    return 0;
}
"#,
        &["42 test"],
    );
}

// Multi-line macro with backslash continuation
#[test]
fn multiline_macro() {
    assert_outputs(
        r#"
#include <stdio.h>
#define SWAP(a, b) \
    do { int _tmp = a; a = b; b = _tmp; } while(0)
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

// #undef
#[test]
fn undef_redefine_macro() {
    assert_outputs(
        r#"
#include <stdio.h>
#define VALUE 10
#undef VALUE
#define VALUE 20
int main() {
    printf("%d\n", VALUE);
    return 0;
}
"#,
        &["20"],
    );
}

// #ifdef / #ifndef
#[test]
fn ifdef_defined_macro() {
    assert_outputs(
        r#"
#include <stdio.h>
#define FEATURE
int main() {
#ifdef FEATURE
    printf("enabled\n");
#else
    printf("disabled\n");
#endif
    return 0;
}
"#,
        &["enabled"],
    );
}

#[test]
fn ifndef_undefined_macro() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
#ifndef MISSING
    printf("not defined\n");
#endif
    return 0;
}
"#,
        &["not defined"],
    );
}

// Nested macro expansion
#[test]
fn nested_macro_expansion() {
    assert_outputs(
        r#"
#include <stdio.h>
#define DOUBLE(x) ((x) * 2)
#define QUAD(x) DOUBLE(DOUBLE(x))
int main() {
    printf("%d\n", QUAD(5));
    return 0;
}
"#,
        &["20"],
    );
}
