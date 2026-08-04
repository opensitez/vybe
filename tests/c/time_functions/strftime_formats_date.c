// vybe-test: c/time_functions/strftime_formats_date
// origin: languages/c/tests/c/test_time_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <time.h>
#include <string.h>
int main() {const char *__w[] = {"1970-01-01\n"};
int __n = 1, __i = 0;

    time_t epoch = 0;
    struct tm *t = gmtime(&epoch);
    char buf[64];
    strftime(buf, sizeof(buf), "%Y-%m-%d", t);
    { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

