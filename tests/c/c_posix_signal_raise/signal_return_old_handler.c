// vybe-test: c/c_posix_signal_raise/signal_return_old_handler
// origin: languages/c/tests/c/test_c_posix_signal_raise.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <signal.h>
void h(int s) {}
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 void (*old)(int) = signal(SIGUSR1, h); void (*old2)(int) = signal(SIGUSR1, SIG_IGN); { char __t[512]; snprintf(__t, sizeof(__t), "%d", old2 == h);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

