// vybe-test: c/time_epoch_calendar_math/mktime_march_first_after_leap_feb
// origin: languages/c/tests/c/test_time_epoch_calendar_math.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <time.h>
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
struct tm t={.tm_year=124,.tm_mon=2,.tm_mday=1,.tm_hour=0,.tm_min=0,.tm_sec=0}; time_t v=mktime(&t); { char __t[512]; snprintf(__t, sizeof(__t), "%lld\n", (long long)difftime(v, 0) > 1700000000);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

