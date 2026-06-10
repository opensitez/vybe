use super::helpers::*;

// C11 _Thread_local storage duration
#[test]
fn thread_local_basic_declaration() {
    assert_outputs(
        r#"
#include <stdio.h>
_Thread_local int tls_var = 0;
int main() {
    tls_var = 42;
    printf("%d\n", tls_var);
    return 0;
}
"#,
        &["42"],
    );
}

#[test]
fn thread_local_initialized_value() {
    assert_outputs(
        r#"
#include <stdio.h>
_Thread_local int counter = 100;
int main() {
    counter += 5;
    printf("%d\n", counter);
    return 0;
}
"#,
        &["105"],
    );
}

#[test]
fn thread_local_keyword_stdc_threads() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <threads.h>
thread_local int x = 7;
int main() {
    printf("%d\n", x);
    return 0;
}
"#,
        &["7"],
    );
}
