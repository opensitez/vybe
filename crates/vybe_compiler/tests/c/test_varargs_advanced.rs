use super::helpers::*;

#[test]
fn va_copy_duplicates_list() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdarg.h>
int sum_va(int n, va_list ap) {
    int total = 0;
    for (int i = 0; i < n; i++) total += va_arg(ap, int);
    return total;
}
int double_sum(int n, ...) {
    va_list ap, ap2;
    va_start(ap, n);
    va_copy(ap2, ap);
    int s1 = sum_va(n, ap);
    int s2 = sum_va(n, ap2);
    va_end(ap2);
    va_end(ap);
    return s1 + s2;
}
int main() {
    printf("%d\n", double_sum(3, 10, 20, 30));
    return 0;
}
"#,
        &["120"],
    );
}

#[test]
fn va_arg_multiple_types() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdarg.h>
void show(const char *fmt, ...) {
    va_list ap;
    va_start(ap, fmt);
    while (*fmt) {
        if (*fmt == 'i') printf("%d ", va_arg(ap, int));
        else if (*fmt == 's') printf("%s ", va_arg(ap, char*));
        else if (*fmt == 'f') printf("%.1f ", va_arg(ap, double));
        fmt++;
    }
    printf("\n");
    va_end(ap);
}
int main() {
    show("isf", 42, "hi", 3.14);
    return 0;
}
"#,
        &["42 hi 3.1 "],
    );
}

#[test]
fn variadic_min_function() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdarg.h>
int vmin(int n, ...) {
    va_list ap;
    va_start(ap, n);
    int m = va_arg(ap, int);
    for (int i = 1; i < n; i++) {
        int v = va_arg(ap, int);
        if (v < m) m = v;
    }
    va_end(ap);
    return m;
}
int main() {
    printf("%d\n", vmin(4, 5, 2, 8, 1));
    return 0;
}
"#,
        &["1"],
    );
}
