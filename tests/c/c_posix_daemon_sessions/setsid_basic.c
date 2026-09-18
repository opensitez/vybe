// vybe-test: c/c_posix_daemon_sessions/setsid_basic
// origin: languages/c/tests/c/test_c_posix_daemon_sessions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <sys/wait.h>
int main() {
    pid_t p = fork();
    if (p == 0) {
        pid_t sid = setsid();
        if (sid <= 0) _exit(1);
        _exit(0);
    }
    int st = 0;
    waitpid(p, &st, 0);
    if (!WIFEXITED(st) || WEXITSTATUS(st) != 0) {
        printf("FAIL: child exited with non-zero\n");
        assert(0);
    }
    return 0;
}

