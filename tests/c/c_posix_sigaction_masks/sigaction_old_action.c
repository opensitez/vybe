// vybe-test: c/c_posix_sigaction_masks/sigaction_old_action
// origin: languages/c/tests/c/test_c_posix_sigaction_masks.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <signal.h>
void h(int s) {}
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 struct sigaction sa, old; sa.sa_handler = h; sigemptyset(&sa.sa_mask); sa.sa_flags = 0; sigaction(SIGUSR1, &sa, NULL); sigaction(SIGUSR1, NULL, &old); { char __t[512]; snprintf(__t, sizeof(__t), "%d", old.sa_handler == h);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

