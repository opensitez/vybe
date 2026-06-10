use super::helpers::*;

#[test]
fn va_list_sum_ints() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdarg.h>
int sum(int count, ...) {
    va_list args;
    va_start(args, count);
    int total = 0;
    for (int i = 0; i < count; i++) {
        total += va_arg(args, int);
    }
    va_end(args);
    return total;
}
int main() {
    printf("%d\n", sum(3, 10, 20, 30));
    return 0;
}
"#,
        &["60"],
    );
}

#[test]
fn va_list_print_strings() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdarg.h>
void print_all(int n, ...) {
    va_list args;
    va_start(args, n);
    for (int i = 0; i < n; i++) {
        printf("%s\n", va_arg(args, char*));
    }
    va_end(args);
}
int main() {
    print_all(3, "a", "b", "c");
    return 0;
}
"#,
        &["a", "b", "c"],
    );
}

#[test]
fn va_list_mixed_types() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdarg.h>
void show(int n, ...) {
    va_list ap;
    va_start(ap, n);
    int i = va_arg(ap, int);
    double d = va_arg(ap, double);
    va_end(ap);
    printf("%d %.1f\n", i, d);
}
int main() {
    show(2, 42, 3.14);
    return 0;
}
"#,
        &["42 3.1"],
    );
}

#[test]
fn vsnprintf_uses_va_list() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdarg.h>
void my_log(const char *fmt, ...) {
    char buf[64];
    va_list ap;
    va_start(ap, fmt);
    vsnprintf(buf, sizeof(buf), fmt, ap);
    va_end(ap);
    printf("%s\n", buf);
}
int main() {
    my_log("value=%d name=%s", 7, "test");
    return 0;
}
"#,
        &["value=7 name=test"],
    );
}
