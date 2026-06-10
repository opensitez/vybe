use super::helpers::*;

// Anonymous structs/unions (C11)
#[test]
fn anonymous_union_in_struct() {
    assert_outputs(
        r#"
#include <stdio.h>
struct Value {
    int type;
    union {
        int i;
        float f;
    };
};
int main() {
    struct Value v;
    v.type = 0;
    v.i = 42;
    printf("%d\n", v.i);
    return 0;
}
"#,
        &["42"],
    );
}

#[test]
fn anonymous_struct_in_union() {
    assert_outputs(
        r#"
#include <stdio.h>
union Data {
    struct { int lo; int hi; };
    long long full;
};
int main() {
    union Data d;
    d.lo = 1; d.hi = 0;
    printf("%d\n", d.lo);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn named_union_with_shared_storage() {
    assert_outputs(
        r#"
#include <stdio.h>
union Num {
    int i;
    float f;
    char bytes[4];
};
int main() {
    union Num n;
    n.i = 0x41424344;
    printf("%d\n", sizeof(union Num));
    return 0;
}
"#,
        &["4"],
    );
}

#[test]
fn anonymous_union_float_and_int() {
    assert_outputs(
        r#"
#include <stdio.h>
struct Tagged {
    int tag;
    union { int as_int; float as_float; };
};
int main() {
    struct Tagged t;
    t.tag = 1;
    t.as_float = 1.5f;
    printf("%.1f\n", t.as_float);
    return 0;
}
"#,
        &["1.5"],
    );
}
