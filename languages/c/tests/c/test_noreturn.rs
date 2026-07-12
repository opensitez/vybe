use super::helpers::*;

// _Noreturn / [[noreturn]] attribute
#[test]
fn noreturn_function_attribute() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdlib.h>
_Noreturn void fatal(const char *msg) {
    printf("%s\n", msg);
    exit(0);
}
int main() {
    fatal("bye");
}
"#,
        &["bye"],
    );
}

#[test]
fn noreturn_function_called_conditionally() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdlib.h>
_Noreturn void die(void) {
    printf("dead\n");
    exit(0);
}
int main() {
    int x = 1;
    if (x > 100) die();
    printf("alive\n");
    return 0;
}
"#,
        &["alive"],
    );
}

// noreturn-style with stdnoreturn.h
#[test]
fn noreturn_macro_in_header() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdnoreturn.h>
#include <stdlib.h>
noreturn void stop(void) {
    exit(0);
}
int main() {
    printf("go\n");
    return 0;
}
"#,
        &["go"],
    );
}
