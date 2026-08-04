// vybe-test: c/c_posix_timers/clock_gettime_invalid_clock
// origin: languages/c/tests/c/test_c_posix_timers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <time.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 struct timespec ts; int r = clock_gettime(99999, &ts); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == -1);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

