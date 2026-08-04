// vybe-test: c/c_posix_signal_raise/signal_multiple
// origin: languages/c/tests/c/test_c_posix_signal_raise.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"21"};
static int __n = 1, __i = 0;
#include <signal.h>
#include <stdlib.h>
void h1(int s) { { char __t[512]; snprintf(__t, sizeof(__t), "1");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } exit(0); }
void h2(int s) { { char __t[512]; snprintf(__t, sizeof(__t), "2");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } }
int main() { signal(SIGUSR1, h1); signal(SIGUSR2, h2); raise(SIGUSR2); raise(SIGUSR1); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

