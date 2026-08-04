// vybe-test: c/c_posix_kill_alarm/getitimer_basic
// origin: languages/c/tests/c/test_c_posix_kill_alarm.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/time.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 struct itimerval it = {0}; it.it_value.tv_sec = 10; setitimer(ITIMER_REAL, &it, NULL); struct itimerval old; getitimer(ITIMER_REAL, &old); { char __t[512]; snprintf(__t, sizeof(__t), "%d", old.it_value.tv_sec > 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

