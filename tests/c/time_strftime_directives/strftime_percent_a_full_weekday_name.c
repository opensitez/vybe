// vybe-test: c/time_strftime_directives/strftime_percent_a_full_weekday_name
// origin: languages/c/tests/c/test_time_strftime_directives.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <time.h>
int main() {
const char *__w[] = {"Saturday\n"};
int __n = 1, __i = 0;
struct tm t={.tm_year=124,.tm_mon=5,.tm_mday=15,.tm_wday=6}; char b[16]; strftime(b,sizeof(b),"%A",&t); { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

