// vybe-test: c/c_posix_time/posix_time_complex_formatting
// origin: languages/c/tests/c/test_c_posix_time.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#define _POSIX_C_SOURCE 200809L
#include <time.h>
#include <string.h>
#include <stdlib.h>

int main() {const char *__w[] = {"2023-10-15T14:30:45Z"};
int __n = 1, __i = 0;

    setenv("TZ", "UTC", 1);
    tzset();

    struct tm timeinfo = {0};
    timeinfo.tm_year = 2023 - 1900;
    timeinfo.tm_mon = 10 - 1; // October
    timeinfo.tm_mday = 15;
    timeinfo.tm_hour = 14;
    timeinfo.tm_min = 30;
    timeinfo.tm_sec = 45;
    timeinfo.tm_isdst = 0;
    
    time_t t = mktime(&timeinfo);
    
    char buf[128];
    // ISO 8601 format
    strftime(buf, sizeof(buf), "%Y-%m-%dT%H:%M:%SZ", gmtime(&t));
    { char __t[512]; snprintf(__t, sizeof(__t), "%s", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

