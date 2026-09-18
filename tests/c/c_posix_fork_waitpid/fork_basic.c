// vybe-test: c/c_posix_fork_waitpid/fork_basic
// origin: languages/c/tests/c/test_c_posix_fork_waitpid.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <sys/wait.h>
int main() {
    pid_t p = fork();
    if (p == 0) {
        char __t[512];
        snprintf(__t, sizeof(__t), "child");
        if (strcmp(__t, "child") != 0) { printf("FAIL child: got [%s]\n", __t); _exit(1); }
        _exit(0);
    } else if (p > 0) {
        int st = 0;
        waitpid(p, &st, 0);
        if (!WIFEXITED(st) || WEXITSTATUS(st) != 0) { printf("FAIL child status\n"); assert(0); }
        char __t[512];
        snprintf(__t, sizeof(__t), "parent");
        if (strcmp(__t, "parent") != 0) { printf("FAIL parent: got [%s]\n", __t); assert(0); }
    } else {
        assert(0);
    }
    return 0;
}

