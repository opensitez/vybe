// vybe-test: c/c_posix_fork_waitpid/fork_flush_stdout
// origin: languages/c/tests/c/test_c_posix_fork_waitpid.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <sys/wait.h>
int main() {
    char __t[512];
    snprintf(__t, sizeof(__t), "A");
    if (strcmp(__t, "A") != 0) { printf("FAIL A: got [%s]\n", __t); assert(0); }
    fflush(stdout);
    pid_t p = fork();
    if (p == 0) {
        snprintf(__t, sizeof(__t), "B");
        if (strcmp(__t, "B") != 0) { printf("FAIL B: got [%s]\n", __t); _exit(1); }
        _exit(0);
    }
    int st = 0;
    waitpid(p, &st, 0);
    if (!WIFEXITED(st) || WEXITSTATUS(st) != 0) { printf("FAIL child status\n"); assert(0); }
    snprintf(__t, sizeof(__t), "C");
    if (strcmp(__t, "C") != 0) { printf("FAIL C: got [%s]\n", __t); assert(0); }
    return 0;
}

