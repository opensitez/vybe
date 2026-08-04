// vybe-test: c/signal/signal_handler_registered
// origin: languages/c/tests/c/test_signal.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <signal.h>
static int caught = 0;
void handler(int sig) { caught = sig; }
int main() {const char *__w[] = {"1\n"};
int __n = 1, __i = 0;

    signal(SIGUSR1, handler);
    raise(SIGUSR1);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", caught == SIGUSR1 ? 1 : 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

