use super::helpers::*;

#[test]
fn signal_constants_defined() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <signal.h>
int main() {
    printf("%d\n", SIGINT > 0 ? 1 : 0);
    printf("%d\n", SIGTERM > 0 ? 1 : 0);
    printf("%d\n", SIGABRT > 0 ? 1 : 0);
    return 0;
}
"#,
        &["1", "1", "1"],
    );
}

#[test]
fn signal_handler_registered() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <signal.h>
static int caught = 0;
void handler(int sig) { caught = sig; }
int main() {
    signal(SIGUSR1, handler);
    raise(SIGUSR1);
    printf("%d\n", caught == SIGUSR1 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn signal_sig_ign_ignores() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <signal.h>
int main() {
    signal(SIGUSR1, SIG_IGN);
    raise(SIGUSR1);
    printf("ok\n");
    return 0;
}
"#,
        &["ok"],
    );
}

#[test]
fn signal_sig_dfl_constant() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <signal.h>
int main() {
    printf("%d\n", SIG_DFL == SIG_DFL ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn raise_returns_zero_on_success() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <signal.h>
static int count = 0;
void handler(int sig) { count++; }
int main() {
    signal(SIGUSR1, handler);
    int ret = raise(SIGUSR1);
    printf("%d %d\n", ret, count);
    return 0;
}
"#,
        &["0 1"],
    );
}
