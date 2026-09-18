// vybe-test: c/c_posix_sigaction_masks/sigwait_basic
// origin: languages/c/tests/c/test_c_posix_sigaction_masks.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <signal.h>
#include <unistd.h>
#include <sys/wait.h>
int main() {const char *__w[] = {"caught"};
int __n = 1, __i = 0;
 pid_t p = fork(); if (p==0) { sigset_t s; sigemptyset(&s); sigaddset(&s, SIGUSR1); sigprocmask(SIG_BLOCK, &s, NULL); int sig = 0; sigwait(&s, &sig); _exit(sig == SIGUSR1 ? 0 : 1); } sleep(1); kill(p, SIGUSR1); int status = 0; wait(&status); int passed = (WIFEXITED(status) && WEXITSTATUS(status) == 0); { char __t[512]; snprintf(__t, sizeof(__t), "%s", passed ? "caught" : "fail");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

