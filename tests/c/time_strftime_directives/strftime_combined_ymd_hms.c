// vybe-test: c/time_strftime_directives/strftime_combined_ymd_hms
// origin: languages/c/tests/c/test_time_strftime_directives.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <time.h>
int main() {
const char *__w[] = {"19700101000000\n"};
int __n = 1, __i = 0;
struct tm t={.tm_year=70,.tm_mon=0,.tm_mday=1,.tm_hour=0,.tm_min=0,.tm_sec=0}; char b[32]; strftime(b,sizeof(b),"%Y%m%d%H%M%S",&t); { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

