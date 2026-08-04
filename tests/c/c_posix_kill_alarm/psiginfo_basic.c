// vybe-test: c/c_posix_kill_alarm/psiginfo_basic
// origin: languages/c/tests/c/test_c_posix_kill_alarm.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <signal.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 siginfo_t info = {0}; info.si_signo = SIGINT; psiginfo(&info, "Info"); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

