// vybe-test: c/lang_run_breadth2/signal_raise_ignored
// origin: languages/c/tests/c/test_lang_run_breadth2.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <signal.h>
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
signal(SIGUSR2,SIG_IGN); raise(SIGUSR2); { char __t[512]; snprintf(__t, sizeof(__t), "1\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

