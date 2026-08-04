// vybe-test: c/time_functions/difftime_positive_for_later
// origin: languages/c/tests/c/test_time_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <time.h>
int main() {const char *__w[] = {"1000\n"};
int __n = 1, __i = 0;

    time_t t1 = 1000;
    time_t t2 = 2000;
    double diff = difftime(t2, t1);
    { char __t[512]; snprintf(__t, sizeof(__t), "%.0f\n", diff);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

