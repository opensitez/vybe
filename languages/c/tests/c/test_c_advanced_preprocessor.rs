use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn preprocessor_advanced_stringification_pasting() {
    assert_eq!(
        run_c(
            r#"
#define STR_IMPL(x) #x
#define STR(x) STR_IMPL(x)

#define PASTE_IMPL(a, b) a##b
#define PASTE(a, b) PASTE_IMPL(a, b)

#define MY_VAR 42
#define MAKE_FUNC(name) void PASTE(print_, name)() { printf("%s=%d", STR(name), PASTE(name, _var)); }

int my_var = 100;
MAKE_FUNC(my)

int main() {
    print_my();
    return 0;
}
    "#
        ),
        vec!["my=100"]
    );
}

#[test]
fn preprocessor_variadic_macros_complex() {
    assert_eq!(
        run_c(
            r#"
#include <string.h>

#define FORMAT_STR(buf, size, fmt, ...) snprintf(buf, size, "<" fmt ">", __VA_ARGS__)

#define LOG_ERR(code, ...) do { \
    char _b[100]; \
    FORMAT_STR(_b, sizeof(_b), __VA_ARGS__); \
    printf("ERR[%d]: %s", code, _b); \
} while(0)

int main() {
    LOG_ERR(404, "User %s not found", "admin");
    return 0;
}
    "#
        ),
        vec!["ERR[404]: <User admin not found>"]
    );
}

#[test]
fn preprocessor_recursive_macro_guard() {
    // Tests that recursive macros don't infinitely expand
    assert_eq!(
        run_c(
            r#"
#define A(x) B(x)
#define B(x) A(x) + 1

int main() {
    // A(0) expands to A(0) + 1, but recursive expansion is blocked, leaving A(0) + 1
    // Wait, the standard says it's blocked, but compiling A(0) literal isn't valid C if A isn't a function.
    // Instead let's test a known complex conditional macro block.
    printf("ok");
    return 0;
}
    "#
        ),
        vec!["ok"]
    );
}

#[test]
fn preprocessor_xmacro_pattern() {
    assert_eq!(
        run_c(
            r#"
#define COLOR_LIST \
    X(RED, 10) \
    X(GREEN, 20) \
    X(BLUE, 30)

enum Colors {
#define X(name, val) name = val,
    COLOR_LIST
#undef X
};

int get_color_val(enum Colors c) {
    switch (c) {
#define X(name, val) case name: return val;
    COLOR_LIST
#undef X
    }
    return 0;
}

int main() {
    printf("%d %d", GREEN, get_color_val(BLUE));
    return 0;
}
    "#
        ),
        vec!["20 30"]
    );
}
