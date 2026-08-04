// vybe-test: c/c_posix_sigaction_masks/sigaction_basic
// origin: languages/c/tests/c/test_c_posix_sigaction_masks.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"act"};
static int __n = 1, __i = 0;
#define _POSIX_C_SOURCE 200809L
#include <signal.h>
#include <stdlib.h>
void h(int s) { { char __t[512]; snprintf(__t, sizeof(__t), "act");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } exit(0); }
int main() { struct sigaction sa; sa.sa_handler = h; sigemptyset(&sa.sa_mask); sa.sa_flags = 0; sigaction(SIGUSR1, &sa, NULL); raise(SIGUSR1); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

