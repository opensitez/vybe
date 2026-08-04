// vybe-test: c/c_posix_syslog_openlog/vsyslog_basic
// origin: languages/c/tests/c/test_c_posix_syslog_openlog.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <syslog.h>
#include <stdarg.h>
void log_it(int prio, const char *fmt, ...) { va_list args; va_start(args, fmt); vsyslog(prio, fmt, args); va_end(args); }
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 log_it(LOG_INFO, "vsyslog %d", 42); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

