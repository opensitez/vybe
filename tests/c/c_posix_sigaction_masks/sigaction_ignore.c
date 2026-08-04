// vybe-test: c/c_posix_sigaction_masks/sigaction_ignore
// origin: languages/c/tests/c/test_c_posix_sigaction_masks.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <signal.h>
int main() {const char *__w[] = {"ignored"};
int __n = 1, __i = 0;
 struct sigaction sa; sa.sa_handler = SIG_IGN; sigemptyset(&sa.sa_mask); sa.sa_flags = 0; sigaction(SIGUSR1, &sa, NULL); raise(SIGUSR1); { char __t[512]; snprintf(__t, sizeof(__t), "ignored");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

