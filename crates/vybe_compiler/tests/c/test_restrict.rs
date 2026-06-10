use super::helpers::*;

// restrict is a type qualifier hinting no aliasing; semantics same as without it
#[test]
fn restrict_pointer_parameter() {
    assert_outputs(
        r#"
#include <stdio.h>
void add_arrays(int n, int * restrict a, const int * restrict b) {
    for (int i = 0; i < n; i++) a[i] += b[i];
}
int main() {
    int a[3] = {1, 2, 3};
    int b[3] = {10, 20, 30};
    add_arrays(3, a, b);
    printf("%d %d %d\n", a[0], a[1], a[2]);
    return 0;
}
"#,
        &["11 22 33"],
    );
}

#[test]
fn restrict_pointer_local_variable() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    int x = 42;
    int * restrict p = &x;
    printf("%d\n", *p);
    return 0;
}
"#,
        &["42"],
    );
}

#[test]
fn restrict_in_memcpy_like_function() {
    assert_outputs(
        r#"
#include <stdio.h>
void my_copy(int n, int * restrict dst, const int * restrict src) {
    for (int i = 0; i < n; i++) dst[i] = src[i];
}
int main() {
    int src[4] = {1, 2, 3, 4};
    int dst[4];
    my_copy(4, dst, src);
    printf("%d %d %d %d\n", dst[0], dst[1], dst[2], dst[3]);
    return 0;
}
"#,
        &["1 2 3 4"],
    );
}
