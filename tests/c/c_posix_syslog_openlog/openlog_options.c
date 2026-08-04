// vybe-test: c/c_posix_syslog_openlog/openlog_options
// origin: languages/c/tests/c/test_c_posix_syslog_openlog.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <syslog.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 openlog("opt_test", LOG_PID | LOG_CONS | LOG_NDELAY | LOG_ODELAY | LOG_NOWAIT, LOG_USER); closelog(); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

