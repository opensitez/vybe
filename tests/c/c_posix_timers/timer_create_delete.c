// vybe-test: c/c_posix_timers/timer_create_delete
// origin: languages/c/tests/c/test_c_posix_timers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <time.h>
#include <signal.h>
int main() {const char *__w[] = {"1 1"};
int __n = 1, __i = 0;
 timer_t t; struct sigevent sev = {0}; sev.sigev_notify = SIGEV_NONE; int r1 = timer_create(CLOCK_REALTIME, &sev, &t); int r2 = 0; if (r1 == 0) r2 = timer_delete(t); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", r1 == 0, r2 == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

