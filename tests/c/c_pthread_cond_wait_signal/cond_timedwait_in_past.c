// vybe-test: c/c_pthread_cond_wait_signal/cond_timedwait_in_past
// origin: languages/c/tests/c/test_c_pthread_cond_wait_signal.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <pthread.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 pthread_mutex_t m = PTHREAD_MUTEX_INITIALIZER; pthread_cond_t c = PTHREAD_COND_INITIALIZER; pthread_mutex_lock(&m); struct timespec ts = {0, 0}; int r = pthread_cond_timedwait(&c, &m, &ts); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

