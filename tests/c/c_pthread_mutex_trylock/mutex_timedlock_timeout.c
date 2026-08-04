// vybe-test: c/c_pthread_mutex_trylock/mutex_timedlock_timeout
// origin: languages/c/tests/c/test_c_pthread_mutex_trylock.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <pthread.h>
#include <time.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 pthread_mutex_t m = PTHREAD_MUTEX_INITIALIZER; pthread_mutex_lock(&m); struct timespec ts; clock_gettime(CLOCK_REALTIME, &ts); ts.tv_nsec += 50000000; if(ts.tv_nsec >= 1000000000) { ts.tv_sec++; ts.tv_nsec -= 1000000000; } int r = pthread_mutex_timedlock(&m, &ts); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

