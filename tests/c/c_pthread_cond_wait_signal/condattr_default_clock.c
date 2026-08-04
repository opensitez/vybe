// vybe-test: c/c_pthread_cond_wait_signal/condattr_default_clock
// origin: languages/c/tests/c/test_c_pthread_cond_wait_signal.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <pthread.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 pthread_condattr_t a; pthread_condattr_init(&a); clockid_t c; pthread_condattr_getclock(&a, &c); { char __t[512]; snprintf(__t, sizeof(__t), "%d", c == CLOCK_REALTIME);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

