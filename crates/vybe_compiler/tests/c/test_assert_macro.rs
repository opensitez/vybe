use super::helpers::*;

#[test]
fn assert_passes_on_true_condition() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <assert.h>
int main() {
    assert(1 == 1);
    printf("ok\n");
    return 0;
}
"#,
        &["ok"],
    );
}

#[test]
fn assert_passes_on_nonzero_value() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <assert.h>
int main() {
    assert(42);
    printf("passed\n");
    return 0;
}
"#,
        &["passed"],
    );
}

#[test]
fn assert_with_complex_expression() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <assert.h>
int main() {
    int x = 5;
    assert(x > 0 && x < 10);
    printf("%d\n", x);
    return 0;
}
"#,
        &["5"],
    );
}

#[test]
fn assert_multiple_assertions() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <assert.h>
int main() {
    assert(1);
    assert(2 + 2 == 4);
    assert('A' == 65);
    printf("all ok\n");
    return 0;
}
"#,
        &["all ok"],
    );
}

#[test]
fn ndebug_disables_assert() {
    assert_outputs(
        r#"
#include <stdio.h>
#define NDEBUG
#include <assert.h>
int main() {
    assert(0);
    printf("not aborted\n");
    return 0;
}
"#,
        &["not aborted"],
    );
}
