// vybe-test: c/c_posix_signal_raise/kill_self
// origin: languages/c/tests/c/test_c_posix_signal_raise.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"killed"};
static int __n = 1, __i = 0;
#define _POSIX_C_SOURCE 200809L
#include <signal.h>
#include <unistd.h>
void h(int s) { { char __t[512]; snprintf(__t, sizeof(__t), "killed");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } _exit(0); }
int main() { signal(SIGUSR1, h); kill(getpid(), SIGUSR1); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

