use super::helpers::*;

// _Alignof returns alignment requirement; _Alignas sets it
#[test]
fn alignof_fundamental_types() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    printf("%d\n", (int)_Alignof(char));
    printf("%d\n", (int)_Alignof(int));
    printf("%d\n", (int)_Alignof(double));
    return 0;
}
"#,
        &["1", "4", "8"],
    );
}

#[test]
fn alignof_pointer() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    printf("%d\n", (int)_Alignof(void*) >= 4 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn alignof_struct() {
    assert_outputs(
        r#"
#include <stdio.h>
struct S { int x; char c; };
int main() {
    printf("%d\n", (int)_Alignof(struct S) >= 4 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn alignas_variable_alignment() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    _Alignas(16) char buf[16];
    printf("%d\n", (unsigned long long)buf % 16 == 0 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn alignas_struct_member() {
    assert_outputs(
        r#"
#include <stdio.h>
struct Aligned {
    char a;
    _Alignas(8) int b;
};
int main() {
    printf("%d\n", (int)_Alignof(struct Aligned) >= 4 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn stdalign_header_macros() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdalign.h>
int main() {
    printf("%d\n", (int)alignof(int));
    return 0;
}
"#,
        &["4"],
    );
}
