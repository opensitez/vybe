// vybe-test: c/time_epoch_calendar_math/gmtime_sec_field_at_offset
// origin: languages/c/tests/c/test_time_epoch_calendar_math.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <time.h>
int main() {
const char *__w[] = {"45\n"};
int __n = 1, __i = 0;
time_t e=45; struct tm *p=gmtime(&e); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", p->tm_sec);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

