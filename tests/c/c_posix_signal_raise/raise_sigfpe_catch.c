// vybe-test: c/c_posix_signal_raise/raise_sigfpe_catch
// origin: languages/c/tests/c/test_c_posix_signal_raise.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"fpe"};
static int __n = 1, __i = 0;
#include <signal.h>
#include <stdlib.h>
void h(int s) { { char __t[512]; snprintf(__t, sizeof(__t), "fpe");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } exit(0); }
int main() { signal(SIGFPE, h); raise(SIGFPE); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

