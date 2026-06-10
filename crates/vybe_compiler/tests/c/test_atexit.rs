use super::helpers::*;

#[test]
fn atexit_single_handler() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdlib.h>
void cleanup() {
    printf("cleanup\n");
}
int main() {
    atexit(cleanup);
    printf("main\n");
    return 0;
}
"#,
        &["main", "cleanup"],
    );
}

#[test]
fn atexit_multiple_handlers_lifo_order() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdlib.h>
void first() { printf("first\n"); }
void second() { printf("second\n"); }
void third() { printf("third\n"); }
int main() {
    atexit(first);
    atexit(second);
    atexit(third);
    printf("main\n");
    return 0;
}
"#,
        &["main", "third", "second", "first"],
    );
}

#[test]
fn atexit_runs_after_main_body() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdlib.h>
void done() { printf("done\n"); }
int main() {
    atexit(done);
    printf("a\n");
    printf("b\n");
    return 0;
}
"#,
        &["a", "b", "done"],
    );
}
