// vybe-test: c/c_posix_timers/timer_create_sigev_thread
// origin: languages/c/tests/c/test_c_posix_timers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <time.h>
#include <signal.h>
void f(union sigval v) {}
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 timer_t t; struct sigevent sev = {0}; sev.sigev_notify = SIGEV_THREAD; sev.sigev_notify_function = f; int r = timer_create(CLOCK_REALTIME, &sev, &t); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if(r==0) timer_delete(t); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

