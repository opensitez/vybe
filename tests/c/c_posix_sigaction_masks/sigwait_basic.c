// vybe-test: c/c_posix_sigaction_masks/sigwait_basic
// origin: languages/c/tests/c/test_c_posix_sigaction_masks.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <signal.h>
#include <unistd.h>
#include <sys/wait.h>
int main() {const char *__w[] = {"caught 30", "caught 10", "caught 16"};
int __n = 3, __i = 0;
 pid_t p = fork(); if (p==0) { sigset_t s; sigemptyset(&s); sigaddset(&s, SIGUSR1); sigprocmask(SIG_BLOCK, &s, NULL); int sig; sigwait(&s, &sig); { char __t[512]; snprintf(__t, sizeof(__t), "caught %d", sig);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } _exit(0); } sleep(1); kill(p, SIGUSR1); wait(NULL); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

