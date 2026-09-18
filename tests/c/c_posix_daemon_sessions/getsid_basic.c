// vybe-test: c/c_posix_daemon_sessions/getsid_basic
// origin: languages/c/tests/c/test_c_posix_daemon_sessions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _XOPEN_SOURCE 500
#include <unistd.h>
#include <sys/wait.h>
int main() {
    pid_t p = fork();
    if (p == 0) {
        setsid();
        if (getsid(0) != getpid()) _exit(1);
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

