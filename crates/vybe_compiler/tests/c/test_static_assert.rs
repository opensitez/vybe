use super::helpers::*;

#[test]
fn static_assert_true_condition() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <assert.h>
_Static_assert(sizeof(int) >= 4, "int must be at least 4 bytes");
int main() {
    printf("ok\n");
    return 0;
}
"#,
        &["ok"],
    );
}

#[test]
fn static_assert_sizeof_char_is_one() {
    assert_outputs(
        r#"
#include <stdio.h>
_Static_assert(sizeof(char) == 1, "char must be 1 byte");
int main() {
    printf("ok\n");
    return 0;
}
"#,
        &["ok"],
    );
}

#[test]
fn static_assert_c11_keyword() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <assert.h>
static_assert(1 == 1, "trivially true");
int main() {
    printf("ok\n");
    return 0;
}
"#,
        &["ok"],
    );
}

#[test]
fn static_assert_in_function_scope() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    _Static_assert(sizeof(long) >= sizeof(int), "long must be at least int size");
    printf("ok\n");
    return 0;
}
"#,
        &["ok"],
    );
}
