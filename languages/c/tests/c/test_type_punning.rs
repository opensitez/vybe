use super::helpers::*;

// Type punning via unions (well-defined in C)
#[test]
fn union_int_float_same_storage() {
    assert_outputs(
        r#"
#include <stdio.h>
union FloatBits {
    float f;
    unsigned int i;
};
int main() {
    union FloatBits u;
    u.f = 0.0f;
    printf("%d\n", u.i == 0 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn union_size_is_max_member() {
    assert_outputs(
        r#"
#include <stdio.h>
union U { char c; short s; int i; double d; };
int main() {
    printf("%d\n", (int)sizeof(union U) >= 8 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn union_shared_bytes_read_write() {
    assert_outputs(
        r#"
#include <stdio.h>
union Raw { int n; char b[4]; };
int main() {
    union Raw r;
    r.n = 0;
    printf("%d\n", r.b[0]);
    return 0;
}
"#,
        &["0"],
    );
}

#[test]
fn union_last_write_wins() {
    assert_outputs(
        r#"
#include <stdio.h>
union Val { int i; float f; };
int main() {
    union Val v;
    v.i = 42;
    v.i = 99;
    printf("%d\n", v.i);
    return 0;
}
"#,
        &["99"],
    );
}

#[test]
fn cast_char_ptr_to_int_ptr() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    int x = 0x01020304;
    char *p = (char*)&x;
    int same = *p != 0;
    printf("%d\n", same);
    return 0;
}
"#,
        &["1"],
    );
}

// `(char*)&int` exposes the integer's object representation as a little-endian
// byte view, so p[i] reads byte i (0x01020304 → 04 03 02 01).
#[test]
fn char_ptr_indexes_object_representation_bytes() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    int x = 0x01020304;
    char *p = (char*)&x;
    printf("%d %d %d %d\n", p[0], p[1], p[2], p[3]);
    return 0;
}
"#,
        &["4 3 2 1"],
    );
}
