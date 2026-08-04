// vybe-test: c/c_posix_timers/timer_settime_gettime
// origin: languages/c/tests/c/test_c_posix_timers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <time.h>
#include <signal.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 timer_t t; struct sigevent sev = {0}; sev.sigev_notify = SIGEV_NONE; timer_create(CLOCK_REALTIME, &sev, &t); struct itimerspec its = {{0,0}, {10,0}}; timer_settime(t, 0, &its, NULL); struct itimerspec curr; timer_gettime(t, &curr); { char __t[512]; snprintf(__t, sizeof(__t), "%d", curr.it_value.tv_sec > 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } timer_delete(t); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

