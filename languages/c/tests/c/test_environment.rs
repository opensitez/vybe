use super::helpers::*;

#[test]
fn getenv_path_is_set() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdlib.h>
int main() {
    char *path = getenv("PATH");
    printf("%d\n", path != NULL ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn getenv_missing_returns_null() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdlib.h>
int main() {
    char *val = getenv("__VYBE_UNDEFINED_VAR__");
    printf("%d\n", val == NULL ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn exit_zero_is_success() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdlib.h>
int main() {
    printf("before\n");
    exit(EXIT_SUCCESS);
    printf("after\n");
    return 0;
}
"#,
        &["before"],
    );
}

#[test]
fn exit_constants_defined() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdlib.h>
int main() {
    printf("%d %d\n", EXIT_SUCCESS, EXIT_FAILURE != 0 ? 1 : 0);
    return 0;
}
"#,
        &["0 1"],
    );
}

#[test]
fn exit_runs_atexit_handlers() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdlib.h>
void goodbye(void) { printf("bye\n"); }
int main() {
    atexit(goodbye);
    printf("hi\n");
    exit(0);
}
"#,
        &["hi", "bye"],
    );
}
