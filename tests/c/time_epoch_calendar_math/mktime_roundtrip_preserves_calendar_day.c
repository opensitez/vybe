// vybe-test: c/time_epoch_calendar_math/mktime_roundtrip_preserves_calendar_day
// origin: languages/c/tests/c/test_time_epoch_calendar_math.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <time.h>
int main() {
const char *__w[] = {"7 4\n"};
int __n = 1, __i = 0;
struct tm t={.tm_year=120,.tm_mon=6,.tm_mday=4,.tm_hour=15,.tm_min=30,.tm_sec=0}; time_t v=mktime(&t); struct tm *p=gmtime(&v); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", p->tm_mon+1, p->tm_mday);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

