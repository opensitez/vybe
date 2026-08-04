// vybe-test: c/time_epoch_calendar_math/time_t_compare_chain_ordering
// origin: languages/c/tests/c/test_time_epoch_calendar_math.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <time.h>
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
time_t a=10,b=20,c=30; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a<b && b<c);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

