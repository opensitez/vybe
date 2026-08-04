// vybe-test: c/cover_time_inttypes/asctime_format
// origin: languages/c/tests/c/test_cover_time_inttypes.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <time.h>
int main() {
const char *__w[] = {"0\n"};
int __n = 1, __i = 0;
struct tm t={.tm_year=70,.tm_mon=0,.tm_mday=1,.tm_hour=0,.tm_min=0,.tm_sec=0,.tm_wday=4}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", asctime(&t)[0]=='T');
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

