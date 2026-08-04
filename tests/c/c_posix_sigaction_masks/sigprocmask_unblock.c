// vybe-test: c/c_posix_sigaction_masks/sigprocmask_unblock
// origin: languages/c/tests/c/test_c_posix_sigaction_masks.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <signal.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 sigset_t s; sigemptyset(&s); sigaddset(&s, SIGUSR1); sigprocmask(SIG_BLOCK, &s, NULL); int r = sigprocmask(SIG_UNBLOCK, &s, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

