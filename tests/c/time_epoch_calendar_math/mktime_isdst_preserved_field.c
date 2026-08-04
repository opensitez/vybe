// vybe-test: c/time_epoch_calendar_math/mktime_isdst_preserved_field
// origin: languages/c/tests/c/test_time_epoch_calendar_math.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <time.h>
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
struct tm t={.tm_year=124,.tm_mon=6,.tm_mday=1,.tm_hour=12,.tm_isdst=1}; mktime(&t); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", t.tm_isdst);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

